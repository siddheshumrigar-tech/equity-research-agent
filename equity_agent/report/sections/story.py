"""Story in charts: six exhibits on one page (2 x 3 grid), all from reported data and the model."""
from __future__ import annotations

from ..common import R, pair
from ..layout import section_title
from ..metrics import key_metrics
from ..styles import HALF_W, RS
from .quarterly import last_quarters, short_quarter

CH_H = 190


def story_in_charts(r: R, num: int) -> list:
    res = r.res
    f, fc, v = res.fin, res.forecast, res.valuation
    m = key_metrics(res)
    labels, n = m["labels"], m["n_actual"]
    ch = r.charts
    out = section_title(r.theme, "Story in charts", num)

    # 1. revenue and growth
    p1 = ch.bars_line("story_rev", labels, m["sales"], m["sales_growth"], n, HALF_W, CH_H, f"Revenue ({RS} Cr)",
                      "Revenue growth (YoY)")
    e1 = r.ex.make("Revenue and year-on-year growth", r.img(p1, HALF_W, CH_H),
                   "Screener.in, model estimates. Dark = reported, light = forecast.")

    # 2. quarterly revenue and EBITDA margin
    q = last_quarters(f, 12)
    if q:
        labs = [short_quarter(x) for x in q["labels"]]
        marg = [e / s if s and e is not None else None for e, s in zip(q["ebitda"], q["sales"])]
        p2 = ch.bars_line("story_q", labs, [s or 0 for s in q["sales"]], marg, len(labs), HALF_W, CH_H,
                          f"Quarterly revenue ({RS} Cr)", "EBITDA margin")
    else:
        p2 = None
    e2 = r.ex.make(f"Last {len(q['labels']) if q else 0} quarters: revenue and EBITDA margin", r.img(p2, HALF_W, CH_H),
                   "Screener.in quarterly results")

    # 3. returns vs cost of capital
    p3 = ch.lines("story_returns", labels, [("ROE", m["roe"]), ("ROCE (pre-tax)", m["roce"])], HALF_W, CH_H, split=n,
                  legend_cols=2, hlines=[(v.wacc, f"WACC {v.wacc:.1%}")])
    e3 = r.ex.make("Return on equity and capital vs WACC", r.img(p3, HALF_W, CH_H),
                   "Screener.in, model; averages of opening and closing balances")

    # 4. cash generation
    cfo = list(f.lines["cfo"]) + list(fc.lines["cfo"])
    capex = list(f.lines["capex"]) + list(fc.lines["capex"])
    fcf = [a - b for a, b in zip(cfo, capex)]
    p4 = ch.grouped_bars("story_cash", labels, [("Cash from operations", cfo), ("Free cash flow (CFO - capex)", fcf)],
                         HALF_W, CH_H, n_actual=n)
    e4 = r.ex.make(f"Cash generation ({RS} Cr)", r.img(p4, HALF_W, CH_H), "Screener.in cash flow, model estimates")

    # 5. valuation history
    if v.pe_history and v.pe_mean:
        p5 = ch.pe_band("story_pe", list(v.pe_history), list(v.pe_history.values()), v.pe_mean, f.pe, HALF_W, CH_H)
        e5 = r.ex.make("P/E at each fiscal year-end vs its mean, and today", r.img(p5, HALF_W, CH_H),
                       "yfinance month-end prices, Screener.in EPS; 'Now' = Screener trailing P/E")
    else:
        e5 = r.ex.make("P/E at each fiscal year-end vs its mean", r.img(None, HALF_W, CH_H), "yfinance, Screener.in")

    # 6. shareholding
    sh, sl = f.shareholding, f.shareholding_labels
    if sh and sl:
        k0 = max(0, len(sl) - 12)
        series = [(lab, [None if x is None else x / 100 for x in sh[lab][k0:]])
                  for lab in ("Promoters", "FIIs", "DIIs", "Public") if lab in sh]
        p6 = ch.lines("story_shp", [short_quarter(x) for x in sl[k0:]], series, HALF_W, CH_H, legend_cols=4)
    else:
        p6 = None
    e6 = r.ex.make("Shareholding pattern (% of equity)", r.img(p6, HALF_W, CH_H), "Screener.in shareholding (BSE/NSE filings)")

    out += [pair(e1, e2), pair(e3, e4), pair(e5, e6)]
    return out
