"""Valuation: method summary -> target, cost of capital, 3-stage DCF, sensitivities, football field, scenarios, flags."""
from __future__ import annotations

from reportlab.platypus import Paragraph, Spacer

from ..common import R, pair
from ..layout import section_title
from ..styles import CONTENT_W, HALF_W, RS, data_table, fmt_num, fmt_pct, fmt_rs, fmt_x
from ...valuation import METHOD_LABELS


def method_table(r: R, width: float = CONTENT_W):
    v = r.res.valuation
    rows = [[r.P("Method", "thl"), r.P("Basis", "thl"), r.P(f"Value ({RS})", "th"), r.P("Weight", "th"),
             r.P(f"Contribution ({RS})", "th")]]
    basis = {"dcf": f"WACC {v.wacc:.1%}, g {v.terminal_growth:.1%}, RONIC {v.ronic:.0%}; rolled forward at Ke - yield",
             "pe": (f"{v.pe_mean:.1f}x mean P/E" + (f" (+/-1 sd {v.pe_sd:.1f}x)" if v.pe_sd else "") +
                    f" on NTM EPS {RS}{v.ntm_eps:,.1f}") if v.pe_mean else "n/a",
             "ev_ebitda": (f"{v.peer_ev_ebitda:.1f}x peer median ({v.peer_count} peers) on NTM EBITDA "
                           f"{RS}{v.ntm_ebitda:,.0f} Cr") if v.peer_ev_ebitda else "no peer data"}
    if v.target_method == "sotp":
        for s in v.sotp:
            rows.append([r.P(f"SOTP: {s['segment']}", "tdl"), r.P(s.get("method", ""), "tdl"),
                         r.P(f"{s['ev_cr']:,.0f} Cr", "td"), r.P("", "td"), r.P("", "td")])
        rows.append([r.P("<b>Sum of the parts (analyst input)</b>", "tdl"), r.P("EV + net cash - minority", "tdl"),
                     r.P(fmt_rs(v.sotp_value), "tdb"), r.P("100%", "tdb"), r.P(fmt_rs(v.sotp_value), "tdb")])
    else:
        for m in ("dcf", "pe", "ev_ebitda"):
            val = v.method_values.get(m)
            w = v.weights.get(m, 0.0)
            rows.append([r.P(METHOD_LABELS[m], "tdl"), r.P(basis[m], "tdl"), r.P(fmt_rs(val), "td"),
                         r.P(fmt_pct(w, 0) if val is not None else "-", "td"),
                         r.P(fmt_rs(val * w) if val is not None else "-", "td")])
    rows.append([r.P("<b>12-month target price</b>", "tdl"), r.P("", "tdl"), r.P(f"<b>{fmt_rs(v.target_price)}</b>", "tdb"),
                 r.P("", "td"), r.P(f"<b>{fmt_rs(v.target_price)}</b>", "tdb")])
    rows.append([r.P("Expected total return / rating", "tdl"),
                 r.P(f"upside {v.upside:+.1%} + dividend yield {v.dividend_yield:.1%}", "tdl"),
                 r.P(fmt_pct(v.expected_total_return, 1, True), "tdb"), r.P("", "td"), r.P(f"<b>{v.rating}</b>", "tdb")])
    w = width
    return data_table(r.theme, rows, [w * 0.27, w * 0.39, w * 0.11, w * 0.09, w * 0.14], total_rows=(len(rows) - 2,))


def _wacc_table(r: R, width: float):
    v, cfg = r.res.valuation, r.res.cfg
    rows = [[r.P("Cost of capital", "thl"), r.P("Value", "th"), r.P("Source / basis", "thl")],
            [r.P("Risk-free rate", "tdl"), r.P(fmt_pct(v.rf, 2), "td"), r.P(cfg.get("risk_free_source", ""), "small")],
            [r.P("Equity risk premium", "tdl"), r.P(fmt_pct(v.erp, 2), "td"), r.P(cfg.get("equity_risk_premium_source", ""), "small")],
            [r.P("Beta", "tdl"), r.P(f"{v.beta:.2f}", "td"), r.P(v.beta_source, "small")],
            [r.P("Cost of equity", "tdl"), r.P(fmt_pct(v.ke, 2), "td"), r.P("CAPM", "small")],
            [r.P("Post-tax cost of debt", "tdl"), r.P(fmt_pct(v.kd_post, 2), "td"), r.P("Historical interest / borrowings x (1 - t)", "small")],
            [r.P("Weights E / D", "tdl"), r.P(f"{v.weight_equity:.0%} / {v.weight_debt:.0%}", "td"), r.P("Market cap; book borrowings", "small")],
            [r.P("<b>WACC</b>", "tdl"), r.P(f"<b>{fmt_pct(v.wacc, 2)}</b>", "tdb"), r.P("", "small")],
            [r.P("Terminal growth", "tdl"), r.P(fmt_pct(v.terminal_growth, 2), "td"), r.P(cfg.get("terminal_growth_source", ""), "small")],
            [r.P("Current ROIC / terminal RONIC", "tdl"), r.P(f"{v.roic:.0%} / {v.ronic:.0%}", "td"), r.P(v.ronic_source, "small")]]
    return data_table(r.theme, rows, [width * 0.33, width * 0.16, width * 0.51], total_rows=(7,))


def _dcf_table(r: R):
    v, fc = r.res.valuation, r.res.forecast
    F = fc.lines
    n_e = len(fc.years)
    heads = fc.year_labels + [f"FY{y % 100:02d}F" for y in v.years[n_e:]]
    rev = list(F["sales"]) + v.revenue_fade
    nopat = list(F["nopat"]) + v.nopat_fade
    reinv = [cx - d + dn for cx, d, dn in zip(F["capex"], F["depreciation"], F["delta_nwc"])] + v.reinvestment_fade
    rows = [[r.P(f"{RS} Cr", "thl")] + [r.P(h, "th") for h in heads]]
    spec = [("Revenue", rev, fmt_num), ("NOPAT", nopat, fmt_num), ("Net reinvestment", reinv, fmt_num),
            ("FCFF", v.fcff, fmt_num), ("Discount factor", v.discount_factors, lambda x: f"{x:.3f}"),
            ("PV of FCFF", v.pv_fcff, fmt_num)]
    for lab, vals, fmt in spec:
        rows.append([r.P(lab, "tdl")] + [r.P(fmt(x), "tdb" if lab == "FCFF" else "td") for x in vals])
    lw = CONTENT_W * 0.15
    cw = (CONTENT_W - lw) / len(heads)
    return data_table(r.theme, rows, [lw] + [cw] * len(heads), estimate_cols=range(n_e + 1, len(heads) + 1),
                      total_rows=(4,))


def _bridge(r: R, width: float):
    v = r.res.valuation
    rows = [[r.P("Equity bridge", "thl"), r.P(f"{RS} Cr", "th")],
            [r.P("PV of FCFF, explicit years", "tdl"), r.P(fmt_num(v.sum_pv_explicit), "td")],
            [r.P("PV of FCFF, fade years", "tdl"), r.P(fmt_num(v.sum_pv_fade), "td")],
            [r.P("PV of terminal value", "tdl"), r.P(fmt_num(v.pv_terminal_value), "td")],
            [r.P("<b>Enterprise value</b>", "tdl"), r.P(f"<b>{fmt_num(v.enterprise_value)}</b>", "tdb")],
            [r.P("  terminal value share of EV", "tdl"), r.P(fmt_pct(v.tv_share, 0), "td")],
            [r.P("  implied terminal multiple of NOPAT", "tdl"), r.P(fmt_x(v.tv_nopat_multiple), "td")],
            [r.P("+ cash + investments", "tdl"), r.P(fmt_num(v.cash + v.investments), "td")],
            [r.P("- borrowings - minority interest", "tdl"), r.P(fmt_num(-(v.borrowings + v.minority_interest)), "td")],
            [r.P("<b>Equity value</b>", "tdl"), r.P(f"<b>{fmt_num(v.equity_value)}</b>", "tdb")],
            [r.P(f"DCF per share today ({RS})", "tdl"), r.P(f"{v.dcf_value:,.2f}", "td")],
            [r.P(f"<b>DCF per share in 12 months ({RS})</b>", "tdl"), r.P(f"<b>{v.dcf_value_12m:,.2f}</b>", "tdb")]]
    return data_table(r.theme, rows, [width * 0.7, width * 0.3], total_rows=(4, 9, 11))


def _grid(r: R, rows_lab, cols_lab, grid, width, corner, center=(2, 2)):
    rows = [[r.P(corner, "thl")] + [r.P(h, "th") for h in cols_lab]]
    for lab, row in zip(rows_lab, grid):
        rows.append([r.P(lab, "tdlb")] + [r.P(fmt_num(x) if x is not None else "n/a", "td") for x in row])
    t = data_table(r.theme, rows, [width * 0.2] + [width * 0.16] * len(cols_lab))
    from reportlab.platypus import TableStyle
    from ..styles import c
    t.setStyle(TableStyle([("BACKGROUND", (center[1] + 1, center[0] + 1), (center[1] + 1, center[0] + 1), c(r.theme.light))]))
    return t


def valuation(r: R, num: int) -> list:
    res = r.res
    v, f = res.valuation, res.fin
    out = section_title(r.theme, "Valuation", num)
    out.append(r.ex.make("From methods to the 12-month target", method_table(r),
                         "Model; P/E history from yfinance month-end prices and Screener.in EPS; peers from Screener.in"))
    out.append(pair(r.ex.make("Cost of capital and long-run returns", _wacc_table(r, HALF_W), "config/market.json, yfinance, model"),
                    r.ex.make("Enterprise value to equity value", _bridge(r, HALF_W), "Model; Screener.in balance sheet")))
    out.append(r.ex.make("Three-stage DCF: explicit years (E), fade years (F), then terminal value", _dcf_table(r),
                         "Model. Fade: growth falls linearly to g, ROIC moves to RONIC, reinvestment = NOPAT x g / ROIC."))
    g1 = _grid(r, [f"{w:.1%}" for w in v.sensitivity_wacc], [f"{g:.1%}" for g in v.sensitivity_g], v.sensitivity_grid,
               HALF_W, "WACC \\ g")
    left = r.ex.make(f"DCF value today ({RS}/share): WACC vs terminal growth", g1, "Model")
    if v.sens_gm_grid:
        d = res.drivers
        g2 = _grid(r, [f"{d.value('sales_growth') + s:.1%}" for s in v.sens_growth],
                   [f"{d.value('ebitda_margin') + s:.1%}" for s in v.sens_margin], v.sens_gm_grid, HALF_W, "Growth \\ margin")
        right = r.ex.make(f"DCF value today ({RS}/share): revenue growth vs EBITDA margin", g2, "Model")
    else:
        right = Spacer(1, 1)
    out.append(pair(left, right))
    rows = []
    if f.low_52 and f.high_52:
        rows.append(("52-week range", f.low_52, f.high_52))
    grid = [x for row in v.sensitivity_grid for x in row if x is not None]
    if grid:
        k = 1 + v.ke - v.dividend_yield
        rows.append(("DCF, 12m (WACC, g +/-1pp)", min(grid) * k, max(grid) * k))
    if v.pe_low is not None:
        rows.append((f"P/E {v.pe_mean - v.pe_sd:.0f}x-{v.pe_mean + v.pe_sd:.0f}x NTM EPS", v.pe_low, v.pe_high))
    if v.ev_ebitda_value is not None and res.peers is not None:
        vals = res.peers.values("ev_ebitda")
        if len(vals) >= 2:
            lo_m, hi_m = min(vals), max(vals)
            to_ps = lambda m: (m * v.ntm_ebitda + v.cash + v.investments - v.borrowings - v.minority_interest) / v.shares_cr
            rows.append((f"Peer EV/EBITDA {lo_m:.0f}x-{hi_m:.0f}x", to_ps(lo_m), to_ps(hi_m)))
    if v.scenarios:
        rows.append(("Bear - bull scenarios", v.scenarios["bear"].target, v.scenarios["bull"].target))
    p = r.charts.football("football", rows, f.price, v.target_price, CONTENT_W, 40 + 24 * len(rows))
    if p:
        out.append(r.ex.make(f"Valuation football field ({RS} per share)", r.img(p, CONTENT_W, 40 + 24 * len(rows))))
    out += scenarios(r)
    out += flags_block(r)
    return out


def scenarios(r: R) -> list:
    v, res = r.res.valuation, r.res
    if not v.scenarios:
        return []
    order = ("bear", "base", "bull")
    sc = v.scenarios
    rows = [[r.P("", "thl")] + [r.P(s.title(), "th") for s in order]]
    spec = [("Revenue growth (driver)", lambda s: fmt_pct(s.sales_growth)), ("EBITDA margin (driver)", lambda s: fmt_pct(s.ebitda_margin)),
            ("Capex % of revenue", lambda s: fmt_pct(s.capex_pct)), ("Terminal RONIC", lambda s: fmt_pct(s.ronic, 0)),
            (f"Revenue {res.forecast.year_labels[-1]} ({RS} Cr)", lambda s: fmt_num(s.revenue_final)),
            (f"DCF value, 12m ({RS})", lambda s: fmt_rs(s.dcf_value_12m)), (f"P/E value ({RS})", lambda s: fmt_rs(s.pe_value)),
            (f"EV/EBITDA value ({RS})", lambda s: fmt_rs(s.ev_ebitda_value)),
            (f"Scenario value ({RS})", lambda s: fmt_rs(s.target)),
            ("Expected total return", lambda s: fmt_pct(s.expected_total_return, 1, True)),
            ("Probability", lambda s: fmt_pct(s.probability, 0))]
    for lab, fn in spec:
        bold = lab.startswith("Scenario value")
        rows.append([r.P(lab, "tdl")] + [r.P(fn(sc[s]), "tdb" if bold else "td") for s in order])
    t = data_table(r.theme, rows, [CONTENT_W * 0.34] + [CONTENT_W * 0.22] * 3, total_rows=(9,))
    pw = v.prob_weighted_value
    note = (f"Probability-weighted value {RS}{pw:,.0f}; the target is the base case"
            + (" (sum of the parts)" if v.target_method == "sotp" else "") + ". ")
    if v.reverse_dcf_growth is not None:
        note += (f"Reverse DCF: the current price implies {v.reverse_dcf_growth:.1%} revenue growth a year for five years "
                 f"(base case {res.drivers.value('sales_growth'):.1%}), all else unchanged.")
    else:
        note += "Reverse DCF: no five-year growth rate between -30% and +80% reaches the current price."
    return [r.ex.make("Bull / base / bear (each through forecast, DCF and multiples)", [t, Spacer(1, 2), r.P(note, "small")],
                      "Model; scenario offsets in config/market.json")]


def flags_block(r: R) -> list:
    res = r.res
    if not res.flags:
        return []
    st = r.st
    out = [Paragraph("Model credibility checks", st["h2"])]
    nar = res.narrative
    for fl in res.flags:
        if fl.severity == "low":
            continue
        resp = nar.flag_responses.get(fl.code) if nar else None
        txt = f"<b>[{fl.severity.upper()}] {fl.code.replace('_', ' ').title()}.</b> {fl.message}"
        if resp:
            txt += f" <b>Analyst response:</b> {r.N(resp)}"
        elif fl.severity == "high":
            txt += " <i>Awaiting analyst response.</i>"
        out.append(Paragraph(txt, st["bul"], bulletText="•"))
    return out
