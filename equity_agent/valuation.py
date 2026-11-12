"""Single valuation engine used by the Excel model, the PDF and the summary JSON.

Cost of capital
    Ke = Rf + beta x ERP (Rf, ERP from ``config/market.json``; beta = Blume-adjusted
    2-yr weekly regression vs Nifty 50). WACC on market-cap and book-borrowing weights.

Three-stage DCF (FCFF, mid-year discounting)
    1. Explicit years 1-5 from the 3-statement model:
       FCFF = EBIT x (1 - t) + D&A - capex - change in NWC (EBIT excludes other income,
       because cash and investments are added in the equity bridge).
    2. Fade years 6-10: revenue growth falls linearly from the year-5 rate to g; the
       year-5 EBIT margin and tax rate are held; ROIC moves linearly from current ROIC to
       the terminal return on new capital (RONIC); reinvestment = NOPAT x growth / ROIC.
    3. Terminal value = NOPAT_10 x (1 + g) x (1 - g / RONIC) / (WACC - g), discounted with
       the year-10 mid-year factor.
    Current ROIC = NOPAT_5 / average operating invested capital (net block + CWIP +
    operating NWC) of years 4-5, capped at ``roic_cap`` and floored at g + 1%.
    RONIC = max(WACC, min(ROIC, WACC + share x (ROIC - WACC))), share from the
    ``ronic_share`` driver (scenario dependent) unless an analyst override is given.

12-month target
    * DCF rolled forward 12 months at Ke minus dividend yield.
    * P/E: mean of the last 5 fiscal-year-end P/Es (actual traded prices, per-year share
      count) x next-twelve-months EPS.
    * EV/EBITDA: peer median trailing EV/EBITDA x next-twelve-months EBITDA.
    Weights from ``target_weights``, renormalised over the methods available. A sourced
    sum-of-the-parts (analyst input) replaces the blend.
    Rating on expected total return = target / price - 1 + dividend yield.
"""
from __future__ import annotations

import copy
import statistics
import sys
from dataclasses import dataclass, field
from datetime import date
from typing import Dict, List, Optional

from .config import CONFIG_PATH, load_market_config, with_defaults  # noqa: F401  (re-exported)
from .drivers import Drivers, build_drivers
from .financials import CompanyFinancials
from .market_data import MarketData
from .model_engine import Forecast, run_forecast

_MONTH = {m: i for i, m in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)}
WACC_STEPS = [-0.01, -0.005, 0.0, 0.005, 0.01]
G_STEPS = [-0.01, -0.005, 0.0, 0.005, 0.01]
GROWTH_STEPS = [-0.02, -0.01, 0.0, 0.01, 0.02]
MARGIN_STEPS = [-0.02, -0.01, 0.0, 0.01, 0.02]
METHODS = ("dcf", "pe", "ev_ebitda")
METHOD_LABELS = {"dcf": "DCF (3-stage, rolled forward 12 months)", "pe": "P/E vs own 5-yr history",
                 "ev_ebitda": "EV/EBITDA vs peer median", "sotp": "Sum of the parts (analyst input)"}


def capm(rf: float, beta: float, erp: float) -> float:
    return rf + beta * erp


def wacc_formula(ke: float, kd_post: float, e: float, d: float) -> float:
    total = e + d
    return ke * e / total + kd_post * d / total


def mid_year_factor(rate: float, t: int) -> float:
    return 1.0 / (1.0 + rate) ** (t - 0.5)


def gordon_tv(fcff_last: float, g: float, rate: float) -> float:
    return fcff_last * (1 + g) / (rate - g)


def terminal_value(nopat_last: float, g: float, wacc: float, ronic: float) -> float:
    """Value-driver terminal value: NOPAT x (1 + g) x (1 - g / RONIC) / (WACC - g)."""
    return nopat_last * (1 + g) * (1 - g / ronic) / (wacc - g)


def rating_for(etr: float, cfg: dict) -> str:
    for threshold, label in cfg["rating_thresholds"]:
        if etr >= threshold:
            return label
    return cfg.get("rating_floor", "SELL")


def verdict_for(upside: float, cfg: dict) -> str:  # backwards-compatible name
    return rating_for(upside, with_defaults(cfg))


def ronic_rule(roic: float, wacc: float, share: float) -> float:
    return max(wacc, min(roic, wacc + share * (roic - wacc)))


# --------------------------------------------------------------------------- #
# three-stage DCF
# --------------------------------------------------------------------------- #
@dataclass
class DCFInputs:
    """Everything the fade and terminal stages need from the explicit forecast."""
    fcff: List[float]
    sales_last: float
    growth_last: float
    margin_last: float  # operating EBIT / revenue, final explicit year
    tax_last: float
    nopat_last: float
    ebitda_margin_last: float
    roic_raw: Optional[float]
    roic_cap: float
    n_fade: int


def dcf_inputs(forecast: Forecast, drivers: Drivers, cfg: dict) -> DCFInputs:
    F = forecast.lines
    s = F["sales"]
    growth_last = s[-1] / s[-2] - 1 if len(s) > 1 and s[-2] else drivers.value("sales_growth", forecast.scenario, len(s) - 1)
    ic = F["invested_capital"]
    avg_ic = (ic[-1] + ic[-2]) / 2 if len(ic) > 1 else ic[-1]
    roic_raw = F["nopat"][-1] / avg_ic if avg_ic > 0 else None
    return DCFInputs(fcff=list(F["fcff"]), sales_last=s[-1], growth_last=growth_last,
                     margin_last=F["ebit_operating"][-1] / s[-1] if s[-1] else 0.0,
                     tax_last=drivers.value("tax_rate", forecast.scenario, len(s) - 1), nopat_last=F["nopat"][-1],
                     ebitda_margin_last=F["ebitda"][-1] / s[-1] if s[-1] else 0.0, roic_raw=roic_raw,
                     roic_cap=cfg.get("roic_cap", 1.0), n_fade=int(cfg.get("fade_years", 5)))


@dataclass
class DCFResult:
    wacc: float
    g: float
    roic: float
    ronic: float
    fcff: List[float]
    revenue_fade: List[float]
    growth_fade: List[float]
    nopat_fade: List[float]
    roic_fade: List[float]
    reinvestment_fade: List[float]
    periods: List[float]
    dfs: List[float]
    pvs: List[float]
    sum_pv_explicit: float
    sum_pv_fade: float
    tv: float
    pv_tv: float
    ev: float
    nopat_terminal: float

    @property
    def sum_pv(self) -> float:
        return self.sum_pv_explicit + self.sum_pv_fade


def effective_roic(inp: DCFInputs, g: float) -> float:
    if inp.roic_raw is None:
        return inp.roic_cap
    return min(inp.roic_cap, max(inp.roic_raw, g + 0.01))


def three_stage(inp: DCFInputs, wacc: float, g: float, share: float, ronic_override: Optional[float] = None) -> DCFResult:
    roic = effective_roic(inp, g)
    ronic = ronic_override if ronic_override is not None else ronic_rule(roic, wacc, share)
    ronic = max(ronic, g + 0.01)
    n_exp, n_fade = len(inp.fcff), inp.n_fade
    rev, rev_f, g_f, np_f, roic_f, re_f, fc_f = inp.sales_last, [], [], [], [], [], []
    for k in range(1, n_fade + 1):
        gk = inp.growth_last + (g - inp.growth_last) * k / n_fade
        rev = rev * (1 + gk)
        nopat = rev * inp.margin_last * (1 - inp.tax_last)
        rk = roic + (ronic - roic) * k / n_fade
        reinv = nopat * gk / rk
        rev_f.append(rev), g_f.append(gk), np_f.append(nopat), roic_f.append(rk), re_f.append(reinv)
        fc_f.append(nopat - reinv)
    fcff = list(inp.fcff) + fc_f
    periods = [t - 0.5 for t in range(1, len(fcff) + 1)]
    dfs = [mid_year_factor(wacc, t) for t in range(1, len(fcff) + 1)]
    pvs = [f * d for f, d in zip(fcff, dfs)]
    nopat_last = np_f[-1] if np_f else inp.nopat_last
    tv = terminal_value(nopat_last, g, wacc, ronic)
    pv_tv = tv * dfs[-1]
    sum_e, sum_f = sum(pvs[:n_exp]), sum(pvs[n_exp:])
    return DCFResult(wacc=wacc, g=g, roic=roic, ronic=ronic, fcff=fcff, revenue_fade=rev_f, growth_fade=g_f,
                     nopat_fade=np_f, roic_fade=roic_f, reinvestment_fade=re_f, periods=periods, dfs=dfs, pvs=pvs,
                     sum_pv_explicit=sum_e, sum_pv_fade=sum_f, tv=tv, pv_tv=pv_tv, ev=sum_e + sum_f + pv_tv,
                     nopat_terminal=nopat_last * (1 + g))


def dcf_per_share(inp: DCFInputs, wacc: float, g: float, share: float, bridge_net: float, shares: float,
                  ronic_override: Optional[float] = None) -> Optional[float]:
    if wacc - g < 0.005:
        return None
    r = three_stage(inp, wacc, g, share, ronic_override)
    return (r.ev + bridge_net) / shares


# --------------------------------------------------------------------------- #
# relative valuation inputs
# --------------------------------------------------------------------------- #
def fy_end_month(fin: CompanyFinancials) -> int:
    return _MONTH.get(fin.period_labels[-1][:3], 3)


def ntm_weight(fin: CompanyFinancials, valuation_date: date) -> float:
    """Share of the next twelve months that falls in the first forecast year."""
    import calendar

    m = fy_end_month(fin)
    y = fin.years[-1] + 1
    fy1_end = date(y, m, calendar.monthrange(y, m)[1])
    return min(1.0, max(0.0, (fy1_end - valuation_date).days / 365.0))


def historical_pe(fin: CompanyFinancials, md: Optional[MarketData], max_years: int = 5, notes: Optional[list] = None):
    """Fiscal-year-end P/E = actual traded price / EPS on that year's share count.

    Share count per year = net profit / reported EPS when it is within 0.67x-1.5x of today's count (buybacks,
    issuance); otherwise (splits / bonus issues not restated) today's count is used.
    """
    if md is None or not (md.month_end_prices_unadj or md.month_end_prices):
        return {}
    prices = md.month_end_prices_unadj or md.month_end_prices
    if notes is not None and not md.month_end_prices_unadj:
        notes.append("Historical P/E uses dividend-adjusted prices (unadjusted prices not cached); P/E band is "
                     "slightly understated.")
    out: Dict[str, float] = {}
    shares_now = fin.shares_cr
    for k in range(fin.n - 1, -1, -1):
        lab = fin.period_labels[k]
        mon = _MONTH.get(lab[:3])
        key = f"{fin.years[k]}-{mon:02d}" if mon else None
        npf = fin.lines["net_profit"][k]
        if key in prices and npf > 0:
            eps = fin.eps[k] if k < len(fin.eps) else None
            shares = shares_now
            if eps and eps > 0:
                implied = npf / eps
                if 0.67 < implied / shares_now < 1.5:
                    shares = implied
            out[f"FY{fin.years[k] % 100:02d}"] = prices[key] / (npf / shares)
        if len(out) >= max_years:
            break
    return dict(reversed(list(out.items())))


def ntm(series_fy1: float, series_fy2: float, w: float) -> float:
    return w * series_fy1 + (1 - w) * series_fy2


def _bridge(fin: CompanyFinancials) -> dict:
    L = fin.lines
    return {"cash": L["cash"][-1], "investments": L["investments"][-1],
            "borrowings": L["borrowings"][-1], "minority_interest": L["minority_interest"][-1]}


def bridge_net(fin: CompanyFinancials) -> float:
    b = _bridge(fin)
    return b["cash"] + b["investments"] - b["borrowings"] - b["minority_interest"]


def blend(values: Dict[str, Optional[float]], weights: Dict[str, float]) -> (float, Dict[str, float]):
    avail = {m: w for m, w in weights.items() if values.get(m) is not None and w > 0}
    tot = sum(avail.values())
    if tot <= 0:
        return values.get("dcf") or 0.0, {"dcf": 1.0}
    used = {m: w / tot for m, w in avail.items()}
    return sum(values[m] * w for m, w in used.items()), used


# --------------------------------------------------------------------------- #
# result containers
# --------------------------------------------------------------------------- #
@dataclass
class ScenarioValue:
    scenario: str
    sales_growth: float
    ebitda_margin: float
    capex_pct: float
    ronic: float
    dcf_value: float
    dcf_value_12m: float
    pe_value: Optional[float]
    ev_ebitda_value: Optional[float]
    target: float
    expected_total_return: float
    probability: float
    revenue_final: float
    ebitda_final: float
    eps_ntm: float


@dataclass
class Valuation:
    # cost of capital
    rf: float
    erp: float
    beta: float
    beta_raw: Optional[float]
    beta_source: str
    ke: float
    kd_pre: float
    tax_rate: float
    kd_post: float
    market_cap: float
    debt: float
    weight_equity: float
    weight_debt: float
    wacc: float
    terminal_growth: float
    # DCF
    years: List[int]
    stage: List[str]
    fcff: List[float]
    revenue_fade: List[float]
    growth_fade: List[float]
    nopat_fade: List[float]
    roic_fade: List[float]
    reinvestment_fade: List[float]
    periods: List[float]
    discount_factors: List[float]
    pv_fcff: List[float]
    sum_pv_explicit: float
    sum_pv_fade: float
    sum_pv: float
    roic_raw: Optional[float]
    roic: float
    ronic: float
    ronic_source: str
    terminal_value: float
    pv_terminal_value: float
    enterprise_value: float
    tv_share: float
    tv_nopat_multiple: Optional[float]
    tv_ebitda_multiple: Optional[float]
    cash: float
    investments: float
    borrowings: float
    minority_interest: float
    equity_value: float
    shares_cr: float
    dcf_value: float
    dividend_yield: float
    dcf_value_12m: float
    # relative
    pe_history: Dict[str, float]
    pe_mean: Optional[float]
    pe_sd: Optional[float]
    pe_source: str
    ntm_weight: float
    ntm_eps: float
    ntm_ebitda: float
    pe_value: Optional[float]
    pe_low: Optional[float]
    pe_high: Optional[float]
    peer_ev_ebitda: Optional[float]
    peer_count: int
    ev_ebitda_value: Optional[float]
    ev_source: str
    current_ev: float
    current_ev_ebitda: Optional[float]
    current_pe: Optional[float]
    # target
    sotp_value: Optional[float]
    sotp: List[dict]
    method_values: Dict[str, Optional[float]]
    weights: Dict[str, float]
    target_method: str
    target_price: float
    price: float
    upside: float
    expected_total_return: float
    rating: str
    valuation_date: str
    # cross-checks and sensitivity
    exit_multiple: Optional[float]
    exit_value: Optional[float]
    sensitivity_wacc: List[float]
    sensitivity_g: List[float]
    sensitivity_grid: List[List[Optional[float]]]
    sens_growth: List[float]
    sens_margin: List[float]
    sens_gm_grid: List[List[Optional[float]]]
    scenarios: Dict[str, ScenarioValue]
    prob_weighted_value: Optional[float]
    reverse_dcf_growth: Optional[float]
    warnings: List[str] = field(default_factory=list)

    @property
    def verdict(self) -> str:  # kept for callers written against the old engine
        return self.rating

    @property
    def pe_median(self) -> Optional[float]:  # old name
        return self.pe_mean


# --------------------------------------------------------------------------- #
# engine
# --------------------------------------------------------------------------- #
def _dividend_yield(fin: CompanyFinancials) -> float:
    return (fin.dividend_yield or 0.0) / 100.0


def _relative_values(fc: Forecast, w: float, shares: float, pe_mean, peer_mult, bnet):
    F = fc.lines
    eps = ntm(F["net_profit"][0], F["net_profit"][1], w) / shares
    ebitda = ntm(F["ebitda"][0], F["ebitda"][1], w)
    pe_v = pe_mean * eps if pe_mean is not None and eps > 0 else None
    ev_v = (peer_mult * ebitda + bnet) / shares if peer_mult is not None and ebitda > 0 else None
    return eps, ebitda, pe_v, ev_v


def compute_valuation(fin: CompanyFinancials, drivers: Drivers, forecast: Forecast, md: Optional[MarketData],
                      cfg: Optional[dict] = None, compute_reverse: bool = True, peers=None, analyst=None,
                      valuation_date: Optional[date] = None, compute_extras: bool = True) -> Valuation:
    cfg = with_defaults(cfg)
    warnings: List[str] = []
    vdate = valuation_date or (date.fromisoformat(cfg["valuation_date"]) if cfg.get("valuation_date") else date.today())
    scen = forecast.scenario
    rf, erp = cfg["risk_free_rate"], cfg["equity_risk_premium"]
    beta = md.beta if md is not None and md.beta is not None else cfg.get("beta_default", 1.0)
    beta_source = md.beta_source if md is not None else f"Default {beta:.2f}"
    beta_raw = md.beta_raw if md is not None else None
    tax = drivers.value("tax_rate", scen)
    kd_pre = drivers.value("interest_rate", scen)
    ke = capm(rf, beta, erp)
    kd_post = kd_pre * (1 - tax)
    e = fin.market_cap
    dbt = fin.lines["borrowings"][-1]
    wacc = wacc_formula(ke, kd_post, e, dbt)
    g = cfg["terminal_growth"]
    if analyst is not None and analyst.terminal_growth is not None:
        g = analyst.terminal_growth
    if wacc - g < 0.01:
        new_g = wacc - 0.01
        warnings.append(f"Terminal growth {g:.2%} too close to WACC {wacc:.2%}; reduced to {new_g:.2%}.")
        g = new_g
    ronic_override = analyst.terminal_ronic if analyst is not None else None
    share = drivers.value("ronic_share", scen)

    inp = dcf_inputs(forecast, drivers, cfg)
    r = three_stage(inp, wacc, g, share, ronic_override)
    br = _bridge(fin)
    bnet = bridge_net(fin)
    shares = fin.shares_cr
    equity = r.ev + bnet
    dcf_value = equity / shares
    dy = _dividend_yield(fin)
    dcf_12m = dcf_value * (1 + ke - dy)
    if forecast.lines["fcff"][-1] <= 0 or r.fcff[-1] <= 0:
        warnings.append("Final-year FCFF is negative, so the terminal value is negative; DCF value is not meaningful.")
    if inp.roic_raw is None:
        warnings.append(f"Operating invested capital is not positive; current ROIC set to the cap {inp.roic_cap:.0%}.")
    ebitda_10 = inp.ebitda_margin_last * (r.revenue_fade[-1] if r.revenue_fade else inp.sales_last)
    tv_nopat_mult = r.tv / r.nopat_terminal if r.nopat_terminal > 0 else None
    tv_ebitda_mult = r.tv / (ebitda_10 * (1 + g)) if ebitda_10 > 0 else None

    # relative valuation
    pe_notes: List[str] = []
    pe_hist = historical_pe(fin, md, notes=pe_notes)
    warnings += pe_notes
    if len(pe_hist) >= 2:
        pe_mean = statistics.mean(pe_hist.values())
        pe_sd = statistics.stdev(pe_hist.values())
        pe_source = (f"Mean of fiscal-year-end P/E for {', '.join(pe_hist)} (yfinance month-end traded price / EPS on "
                     f"that year's share count)")
    elif fin.pe:
        pe_mean, pe_sd = fin.pe, None
        pe_source = "Current trailing P/E from Screener (historical prices unavailable)"
    else:
        pe_mean, pe_sd, pe_source = None, None, "Not available"
    wgt = ntm_weight(fin, vdate)
    peer_mult = peers.median("ev_ebitda") if peers is not None else None
    n_peers = peers.count("ev_ebitda") if peers is not None else 0
    ev_source = (f"Median trailing EV/EBITDA of {n_peers} Screener.in peers" + (f" ({peers.source})" if peers is not None
                 and peers.source else "")) if peer_mult is not None else "Not available (no peer data)"
    eps_ntm, ebitda_ntm, pe_value, ev_value = _relative_values(forecast, wgt, shares, pe_mean, peer_mult, bnet)
    pe_low = (pe_mean - pe_sd) * eps_ntm if pe_value is not None and pe_sd is not None else None
    pe_high = (pe_mean + pe_sd) * eps_ntm if pe_value is not None and pe_sd is not None else None
    ebitda_ttm = fin.ttm.get("ebitda") if fin.ttm.get("ebitda") else fin.latest("ebitda")
    cur_ev = e + br["borrowings"] - br["cash"] - br["investments"]
    cur_ev_ebitda = cur_ev / ebitda_ttm if ebitda_ttm and ebitda_ttm > 0 else None

    # target
    weights_cfg = dict(cfg["target_weights"])
    values = {"dcf": dcf_12m, "pe": pe_value, "ev_ebitda": ev_value}
    blended, used = blend(values, weights_cfg)
    if pe_value is None:
        warnings.append("P/E value unavailable (no P/E history or negative forward EPS); weight moved to other methods.")
    if ev_value is None:
        warnings.append("EV/EBITDA value unavailable (no peer multiples or negative forward EBITDA); weight moved to "
                        "other methods.")
    sotp_rows, sotp_value = [], None
    if analyst is not None and analyst.sotp:
        sotp_rows, sotp_ev = analyst.sotp_rows()
        sotp_value = (sotp_ev + bnet - analyst.sotp_other_adjustment_cr) / shares
    if sotp_value is not None:
        target, method = sotp_value, "sotp"
        used = {"sotp": 1.0}
    else:
        target, method = blended, "blend"
    upside = target / fin.price - 1
    etr = upside + dy
    rating = rating_for(etr, cfg)

    # exit-multiple cross-check: current EV/EBITDA (incl. minority) on final explicit EBITDA, discounted at t = 5
    exit_multiple = exit_value = None
    ebitda_now = fin.latest("ebitda")
    n_exp = len(forecast.years)
    if ebitda_now > 0:
        mkt_ev = e + br["borrowings"] + br["minority_interest"] - br["cash"] - br["investments"]
        exit_multiple = mkt_ev / ebitda_now
        exit_ev = r.sum_pv_explicit + forecast.lines["ebitda"][-1] * exit_multiple / (1 + wacc) ** n_exp
        exit_value = (exit_ev + bnet) / shares

    w_list = [wacc + s for s in WACC_STEPS]
    g_list = [g + s for s in G_STEPS]
    grid = [[dcf_per_share(inp, w, gg, share, bnet, shares, ronic_override) for gg in g_list] for w in w_list]

    gm_grid: List[List[Optional[float]]] = []
    if compute_extras:
        for gs in GROWTH_STEPS:
            row = []
            for ms in MARGIN_STEPS:
                d2 = copy.deepcopy(drivers)
                for key, step in (("sales_growth", gs), ("ebitda_margin", ms)):
                    it = d2.items[key]
                    it.base, it.bull, it.bear = it.base + step, it.bull + step, it.bear + step
                fc2 = run_forecast(fin, d2, scen, len(forecast.years), cfg.get("cwip_years"))
                row.append(dcf_per_share(dcf_inputs(fc2, d2, cfg), wacc, g, share, bnet, shares, ronic_override))
            gm_grid.append(row)

    scen_vals: Dict[str, ScenarioValue] = {}
    pw = None
    if compute_extras:
        probs = cfg["scenarios"].get("probabilities", {"bull": 0.25, "base": 0.5, "bear": 0.25})
        for s in ("bull", "base", "bear"):
            fc_s = forecast if s == scen else run_forecast(fin, drivers, s, len(forecast.years), cfg.get("cwip_years"))
            inp_s = dcf_inputs(fc_s, drivers, cfg)
            r_s = three_stage(inp_s, wacc, g, drivers.value("ronic_share", s), ronic_override)
            dcf_s = (r_s.ev + bnet) / shares
            dcf_s12 = dcf_s * (1 + ke - dy)
            eps_s, ebitda_s, pe_s, ev_s = _relative_values(fc_s, wgt, shares, pe_mean, peer_mult, bnet)
            tgt_s, _ = blend({"dcf": dcf_s12, "pe": pe_s, "ev_ebitda": ev_s}, weights_cfg)
            scen_vals[s] = ScenarioValue(
                scenario=s, sales_growth=drivers.value("sales_growth", s), ebitda_margin=drivers.value("ebitda_margin", s),
                capex_pct=drivers.value("capex_pct", s), ronic=r_s.ronic, dcf_value=dcf_s, dcf_value_12m=dcf_s12,
                pe_value=pe_s, ev_ebitda_value=ev_s, target=tgt_s, expected_total_return=tgt_s / fin.price - 1 + dy,
                probability=probs.get(s, 0.0), revenue_final=fc_s.lines["sales"][-1],
                ebitda_final=fc_s.lines["ebitda"][-1], eps_ntm=eps_s)
        pw = sum(v.target * v.probability for v in scen_vals.values())

    rev = None
    if compute_reverse:
        rev = reverse_dcf_growth(fin, drivers, scen, wacc, g, fin.price, cfg, ronic_override)

    years10 = list(forecast.years) + [forecast.years[-1] + k for k in range(1, inp.n_fade + 1)]
    return Valuation(
        rf=rf, erp=erp, beta=beta, beta_raw=beta_raw, beta_source=beta_source, ke=ke, kd_pre=kd_pre, tax_rate=tax,
        kd_post=kd_post, market_cap=e, debt=dbt, weight_equity=e / (e + dbt), weight_debt=dbt / (e + dbt), wacc=wacc,
        terminal_growth=g, years=years10, stage=["explicit"] * n_exp + ["fade"] * inp.n_fade, fcff=r.fcff,
        revenue_fade=r.revenue_fade, growth_fade=r.growth_fade, nopat_fade=r.nopat_fade, roic_fade=r.roic_fade,
        reinvestment_fade=r.reinvestment_fade, periods=r.periods, discount_factors=r.dfs, pv_fcff=r.pvs,
        sum_pv_explicit=r.sum_pv_explicit, sum_pv_fade=r.sum_pv_fade, sum_pv=r.sum_pv, roic_raw=inp.roic_raw,
        roic=r.roic, ronic=r.ronic,
        ronic_source=("Analyst input" if ronic_override is not None else
                      f"WACC + {share:.0%} x (ROIC - WACC), floored at WACC, capped at ROIC"),
        terminal_value=r.tv, pv_terminal_value=r.pv_tv, enterprise_value=r.ev,
        tv_share=r.pv_tv / r.ev if r.ev else 0.0, tv_nopat_multiple=tv_nopat_mult, tv_ebitda_multiple=tv_ebitda_mult,
        cash=br["cash"], investments=br["investments"], borrowings=br["borrowings"],
        minority_interest=br["minority_interest"], equity_value=equity, shares_cr=shares, dcf_value=dcf_value,
        dividend_yield=dy, dcf_value_12m=dcf_12m, pe_history=pe_hist, pe_mean=pe_mean, pe_sd=pe_sd,
        pe_source=pe_source, ntm_weight=wgt, ntm_eps=eps_ntm, ntm_ebitda=ebitda_ntm, pe_value=pe_value,
        pe_low=pe_low, pe_high=pe_high, peer_ev_ebitda=peer_mult, peer_count=n_peers, ev_ebitda_value=ev_value,
        ev_source=ev_source, current_ev=cur_ev, current_ev_ebitda=cur_ev_ebitda, current_pe=fin.pe,
        sotp_value=sotp_value, sotp=sotp_rows, method_values=values, weights=used, target_method=method,
        target_price=target, price=fin.price, upside=upside, expected_total_return=etr, rating=rating,
        valuation_date=vdate.isoformat(), exit_multiple=exit_multiple, exit_value=exit_value,
        sensitivity_wacc=w_list, sensitivity_g=g_list, sensitivity_grid=grid, sens_growth=list(GROWTH_STEPS),
        sens_margin=list(MARGIN_STEPS), sens_gm_grid=gm_grid, scenarios=scen_vals, prob_weighted_value=pw,
        reverse_dcf_growth=rev, warnings=warnings,
    )


def reverse_dcf_growth(fin: CompanyFinancials, drivers: Drivers, scenario: str, wacc: float, g: float,
                       price: float, cfg: Optional[dict] = None, ronic_override: Optional[float] = None
                       ) -> Optional[float]:
    """Revenue growth for the 5 explicit years (all other drivers fixed, fade and terminal stages as in the
    base case) at which the 3-stage DCF value today equals the market price."""
    cfg = with_defaults(cfg)
    bnet = bridge_net(fin)
    share = drivers.value("ronic_share", scenario)

    def value_at(growth: float) -> Optional[float]:
        d2 = copy.deepcopy(drivers)
        it = d2.items["sales_growth"]
        it.base = it.bull = it.bear = growth
        d2.overrides.pop("sales_growth", None)
        fc = run_forecast(fin, d2, scenario, cfg.get("forecast_years", 5), cfg.get("cwip_years"))
        return dcf_per_share(dcf_inputs(fc, d2, cfg), wacc, g, share, bnet, fin.shares_cr, ronic_override)

    grid = [x / 100 for x in range(-30, 81)]
    vals = [value_at(gr) for gr in grid]
    bracket = None
    for k in range(len(grid) - 1):
        a, b = vals[k], vals[k + 1]
        if a is not None and b is not None and (a - price) * (b - price) <= 0:
            bracket = k
            break
    if bracket is None:
        return None
    lo, hi = grid[bracket], grid[bracket + 1]
    vlo = vals[bracket]
    for _ in range(40):
        mid = (lo + hi) / 2
        vm = value_at(mid)
        if vm is None:
            return None
        if (vlo - price) * (vm - price) <= 0:
            hi = mid
        else:
            lo, vlo = mid, vm
    return (lo + hi) / 2


def _main(argv):
    import argparse

    from .pipeline import analyse

    ap = argparse.ArgumentParser(description="Print WACC, DCF, relative values, target and rating.")
    ap.add_argument("symbol")
    ap.add_argument("--cache-dir", default="cache")
    ap.add_argument("--offline", action="store_true")
    a = ap.parse_args(argv)
    res = analyse(a.symbol, cache_dir=a.cache_dir, offline=a.offline)
    v = res.valuation
    print(f"{res.fin.name}: beta {v.beta:.2f} ({v.beta_source})")
    print(f"  Ke {v.ke:.2%}  Kd(post-tax) {v.kd_post:.2%}  WACC {v.wacc:.2%}  g {v.terminal_growth:.2%}")
    print(f"  ROIC {v.roic:.1%}  RONIC {v.ronic:.1%}  TV share {v.tv_share:.0%}")
    print(f"  EV {v.enterprise_value:,.0f} Cr  Equity {v.equity_value:,.0f} Cr  Shares {v.shares_cr:,.2f} Cr")
    print(f"  DCF Rs {v.dcf_value:,.2f} (12m {v.dcf_value_12m:,.2f})  P/E Rs {v.pe_value or 0:,.2f}  "
          f"EV/EBITDA Rs {v.ev_ebitda_value or 0:,.2f}")
    print(f"  Target Rs {v.target_price:,.2f} vs CMP Rs {v.price:,.2f} -> ETR {v.expected_total_return:+.1%} {v.rating}")
    for w in res.warnings:
        print("  warning:", w)
    return 0


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
