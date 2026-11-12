"""Financial statements appendix: income statement, balance sheet, cash flow and ratios (5 reported + 5 forecast years)."""
from __future__ import annotations

from typing import Callable, List, Optional, Sequence

from ...financials import cost_of_revenue_label
from ...ratios import annual
from ..common import R
from ..layout import section_title
from ..metrics import key_metrics
from ..styles import CONTENT_W, RS, data_table, fmt_num, fmt_pct

N_ACT = 5

Spec = Sequence[tuple]  # (label, values or None for a section header, formatter, bold)


def _cat(r: R, hist: List[Optional[float]], fcast: List[Optional[float]]) -> List[Optional[float]]:
    f = r.res.fin
    return list(hist[max(0, f.n - N_ACT):]) + list(fcast)


def _div(a, b):
    return [x / y if x is not None and y not in (None, 0) else None for x, y in zip(a, b)]


def _table(r: R, spec: Spec, label_share: float = 0.25):
    f, fc = r.res.fin, r.res.forecast
    heads = f.year_labels[max(0, f.n - N_ACT):] + fc.year_labels
    na = min(N_ACT, f.n)
    rows = [[r.P(f"{RS} Cr", "thl")] + [r.P(h, "th") for h in heads]]
    sections, totals = [], []
    for k, (lab, vals, fmt, bold) in enumerate(spec, 1):
        if vals is None:
            sections.append(k)
            rows.append([r.P(f"<b>{lab}</b>", "tdl")] + [r.P("", "td") for _ in heads])
            continue
        if bold:
            totals.append(k)
        rows.append([r.P(f"<b>{lab}</b>" if bold else lab, "tdl")] +
                    [r.P(fmt(v) if v is not None else "n/a", "tdb" if bold else "td") for v in vals])
    lw = CONTENT_W * label_share
    return data_table(r.theme, rows, [lw] + [(CONTENT_W - lw) / len(heads)] * len(heads),
                      estimate_cols=range(na + 1, len(heads) + 1), total_rows=totals, section_rows=sections)


def income_statement(r: R) -> list:
    f, fc = r.res.fin, r.res.forecast
    L, F = f.lines, fc.lines
    c = lambda k, fk=None: _cat(r, L[k], F[fk or k])
    sales, ebitda, npf = c("sales"), c("ebitda"), c("net_profit")
    ebit_h = [e + o - d for e, o, d in zip(L["ebitda"], L["other_income"], L["depreciation"])]
    ebit = _cat(r, ebit_h, F["ebit"])
    m = key_metrics(r.res)
    k0 = max(0, f.n - N_ACT)
    eps = m["eps"][k0:]
    dps = m["dps"][k0:]
    # growth on the full history, so the first column shown has its growth too (as in the Excel model)
    full = lambda k: list(L[k]) + list(F[k])
    g = lambda k: ([None] + [b / a - 1 if a and a > 0 and b is not None else None
                             for a, b in zip(full(k), full(k)[1:])])[k0:]
    pg = lambda x: fmt_pct(x, 1, True)
    ratios = annual(r.res)
    spec = [("Revenue", sales, fmt_num, True), ("  growth", g("sales"), pg, False),
            (cost_of_revenue_label(f, short=True), c("cogs"), fmt_num, False),
            ("Gross profit", c("gross_profit"), fmt_num, False), ("  gross margin", ratios["gross_margin"][k0:], fmt_pct, False),
            ("Operating expenses", c("opex"), fmt_num, False), ("EBITDA", ebitda, fmt_num, True),
            ("  EBITDA margin", _div(ebitda, sales), fmt_pct, False), ("  growth", g("ebitda"), pg, False),
            ("Other income", c("other_income"), fmt_num, False), ("Depreciation", c("depreciation"), fmt_num, False),
            ("EBIT (incl. other income)", ebit, fmt_num, False), ("Interest", c("interest"), fmt_num, False),
            ("Profit before tax", c("pbt"), fmt_num, True), ("Tax", c("tax"), fmt_num, False),
            ("  effective tax rate", _div(c("tax"), c("pbt")), fmt_pct, False),
            ("Minority, associates & other", c("minority_other"), fmt_num, False), ("Net profit", npf, fmt_num, True),
            ("  net margin", _div(npf, sales), fmt_pct, False), ("  growth", g("net_profit"), pg, False),
            (f"EPS ({RS})", eps, lambda x: fmt_num(x, 2), False), (f"DPS ({RS})", dps, lambda x: fmt_num(x, 2), False)]
    return [r.ex.make(f"Income statement ({RS} Cr)", _table(r, spec),
                      "Screener.in (reported, " + ("consolidated" if f.consolidated else "standalone") + "), model "
                      "estimates. Reported tax = PBT x reported tax %. Reported DPS = net profit x payout / today's shares.")]


def balance_sheet(r: R) -> list:
    f, fc = r.res.fin, r.res.forecast
    L, F = f.lines, fc.lines
    c = lambda k, fk=None: _cat(r, L[k], F[fk or k])
    other_liab = _cat(r, [a + b for a, b in zip(L["other_liability_items"], L["advance_from_customers"])], F["other_liab"])
    spec = [("Equity and liabilities", None, None, False),
            ("Equity capital", c("equity_capital"), fmt_num, False), ("Reserves", c("reserves"), fmt_num, False),
            ("Total equity", c("total_equity"), fmt_num, True), ("Minority interest", c("minority_interest"), fmt_num, False),
            ("Long-term borrowings", c("lt_borrowings"), fmt_num, False),
            ("Short-term borrowings", c("st_borrowings"), fmt_num, False),
            ("Lease liabilities", c("lease_liabilities"), fmt_num, False),
            ("Other borrowings", c("other_borrowings"), fmt_num, False),
            ("Total borrowings", c("borrowings"), fmt_num, False), ("Trade payables", c("trade_payables"), fmt_num, False),
            ("Other liabilities (incl. customer advances)", other_liab, fmt_num, False),
            ("Total equity and liabilities", c("total_liabilities"), fmt_num, True),
            ("Assets", None, None, False),
            ("Net fixed assets", c("fixed_assets"), fmt_num, False), ("Capital work in progress", c("cwip"), fmt_num, False),
            ("Investments", c("investments"), fmt_num, False), ("Inventory", c("inventory"), fmt_num, False),
            ("Trade receivables", c("receivables"), fmt_num, False), ("Cash and equivalents", c("cash"), fmt_num, False),
            ("Loans & advances", c("loans_advances"), fmt_num, False),
            ("Other assets", _cat(r, L["other_asset_items"], F["other_ca"]), fmt_num, False),
            ("Total assets", c("total_assets"), fmt_num, True)]
    return [r.ex.make(f"Balance sheet ({RS} Cr)", _table(r, spec),
                      "Screener.in balance sheet and schedules, model estimates. Loans & advances held flat "
                      "(non-operating); CWIP capitalised over the configured years; cash is the cash-flow output.")]


def cash_flow(r: R) -> list:
    f, fc = r.res.fin, r.res.forecast
    L, F = f.lines, fc.lines
    c = lambda k, fk=None: _cat(r, L[k], F[fk or k])
    npf, dep, cfo = c("net_profit"), c("depreciation"), c("cfo")
    wc = _cat(r, [o - p - d for o, p, d in zip(L["cfo"], L["net_profit"], L["depreciation"])],
              [-x for x in F["delta_nwc"]])
    capex = c("capex")
    cfi = c("cfi")
    other_inv = _cat(r, [i + k for i, k in zip(L["cfi"], L["capex"])], [0.0] * len(F["cfi"]))
    divs = _cat(r, L["dividends_paid"], F["dividends"])
    cff = c("cff")
    other_fin = _cat(r, [a + d for a, d in zip(L["cff"], L["dividends_paid"])], [0.0] * len(F["cff"]))
    neg = lambda s: [-x if x is not None else None for x in s]
    fcf = [a - b for a, b in zip(cfo, capex)]
    spec = [("Net profit", npf, fmt_num, False), ("Depreciation", dep, fmt_num, False),
            ("Working capital and other", wc, fmt_num, False), ("Cash from operations", cfo, fmt_num, True),
            ("Capex", neg(capex), fmt_num, False), ("Other investing", other_inv, fmt_num, False),
            ("Cash from investing", cfi, fmt_num, True), ("Dividends paid", neg(divs), fmt_num, False),
            ("Other financing", other_fin, fmt_num, False), ("Cash from financing", cff, fmt_num, True),
            ("Net cash flow", c("net_cash_flow"), fmt_num, True), ("Free cash flow (CFO - capex)", fcf, fmt_num, False),
            ("FCFF used in the DCF", [None] * min(N_ACT, f.n) + list(F["fcff"]), fmt_num, False)]
    return [r.ex.make(f"Cash flow statement ({RS} Cr)", _table(r, spec),
                      "Screener.in cash flow (reported working capital and other = CFO - net profit - depreciation), "
                      "model estimates. Forecast debt and investments held flat, so other investing and financing are 0.")]


def ratios(r: R) -> list:
    res = r.res
    f, fc = res.fin, res.forecast
    L, F = f.lines, fc.lines
    m = key_metrics(res)
    k0 = max(0, f.n - N_ACT)
    cut = lambda s: list(s[k0:])
    c = lambda k, fk=None: _cat(r, L[k], F[fk or k])
    rt = {k: cut(v) for k, v in annual(res).items()}
    ebit = _cat(r, [e + o - d for e, o, d in zip(L["ebitda"], L["other_income"], L["depreciation"])], F["ebit"])
    cover = [e / i if i and i > 0 else None for e, i in zip(ebit, c("interest"))]
    bvps = [q / f.shares_cr for q in c("total_equity")]
    pct, pg = fmt_pct, (lambda x: fmt_pct(x, 1, True))
    xf = lambda x: f"{x:.1f}x"
    days = lambda x: f"{x:.0f}"
    two = lambda x: f"{x:.2f}"
    spec = [("Growth", None, None, False),
            ("Revenue", cut(m["sales_growth"]), pg, False), ("EPS", cut(m["eps_growth"]), pg, False),
            ("Margins", None, None, False),
            ("Gross margin", rt["gross_margin"], pct, False), ("EBITDA margin", cut(m["ebitda_margin"]), pct, False),
            ("Net margin", cut(m["net_margin"]), pct, False),
            ("Returns (average balances)", None, None, False),
            ("ROE", rt["roe"], pct, False), ("ROCE (pre-tax)", rt["roce"], pct, False),
            ("Working capital (days)", None, None, False),
            ("Receivable days", rt["dso"], days, False), ("Inventory days", rt["dio"], days, False),
            ("Payable days", rt["dpo"], days, False), ("Cash conversion cycle", rt["ccc"], days, False),
            ("Balance sheet", None, None, False),
            ("Net debt / equity (x)", cut(m["net_debt_equity"]), two, False), ("Interest cover (x)", cover, xf, False),
            (f"Book value per share ({RS})", bvps, lambda x: fmt_num(x, 1), False),
            ("Valuation at the current price", None, None, False),
            ("P/E", cut(m["pe"]), xf, False), ("EV/EBITDA", cut(m["ev_ebitda"]), xf, False), ("P/B", cut(m["pb"]), xf, False),
            ("Dividend yield", cut(m["div_yield"]), pct, False)]
    return [r.ex.make("Key ratios", _table(r, spec),
                      "Screener.in, model estimates. Inventory days on " + cost_of_revenue_label(f, short=True).lower() +
                      "; payable days on " +
                      (cost_of_revenue_label(f, short=True).lower() if res.drivers.dpo_basis == "cogs" else "total expenses") +
                      "; interest cover = EBIT incl. other income / interest; multiples at today's price and EV.")]


def financial_statements(r: R, num: int) -> list:
    out = section_title(r.theme, "Financial statements", num)
    out += income_statement(r) + balance_sheet(r) + cash_flow(r) + ratios(r)
    return out
