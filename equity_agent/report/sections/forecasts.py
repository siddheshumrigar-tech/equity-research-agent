"""Forecasts vs history: the model's five-year path set against the company's own record.

No consensus estimates are used (there is no free source); the comparison is with reported history.
"""
from __future__ import annotations

from typing import List, Optional

from reportlab.platypus import Spacer

from ..common import R, pair
from ..layout import section_title
from ..metrics import key_metrics
from ..styles import CONTENT_W, HALF_W, data_table, fmt_pct

CH_H = 150


def _cagr(a: Optional[float], b: Optional[float], n: int) -> Optional[float]:
    if a is None or b is None or a <= 0 or b <= 0 or n <= 0:
        return None
    return (b / a) ** (1 / n) - 1


def _avg(vals: List[Optional[float]]) -> Optional[float]:
    v = [x for x in vals if x is not None]
    return sum(v) / len(v) if v else None


def forecasts_vs_history(r: R, num: int) -> list:
    res = r.res
    f, fc = res.fin, res.forecast
    m = key_metrics(res)
    n, nf = f.n, len(fc.years)
    out = section_title(r.theme, "Forecasts vs history", num)
    h5, h3 = min(5, n - 1), min(3, n - 1)
    last = n - 1
    end = n + nf - 1

    def growth_row(label, series):
        return (label, _cagr(series[last - h5], series[last], h5), _cagr(series[last - h3], series[last], h3),
                _cagr(series[last], series[end], nf), True)

    def avg_row(label, series):
        return (label, _avg(series[last - h5 + 1:last + 1]), _avg(series[last - h3 + 1:last + 1]),
                _avg(series[n:end + 1]), False)

    sales = m["sales"]
    capex = list(f.lines["capex"]) + list(fc.lines["capex"])
    cfo = list(f.lines["cfo"]) + list(fc.lines["cfo"])
    payout = [x / p if p and p > 0 else None for x, p in
              zip(list(f.lines["dividends_paid"]) + list(fc.lines["dividends"]), m["net_profit"])]
    spec = [growth_row("Revenue, CAGR", sales), growth_row("EBITDA, CAGR", m["ebitda"]),
            growth_row("Net profit, CAGR", m["net_profit"]), growth_row("EPS, CAGR", m["eps"]),
            avg_row("EBITDA margin, average", m["ebitda_margin"]), avg_row("Net margin, average", m["net_margin"]),
            avg_row("ROE, average", m["roe"]), avg_row("ROCE (pre-tax), average", m["roce"]),
            avg_row("Capex / revenue, average", [c / s if s else None for c, s in zip(capex, sales)]),
            avg_row("CFO / EBITDA, average", [c / e if e and e > 0 else None for c, e in zip(cfo, m["ebitda"])]),
            avg_row("Dividends / net profit, average", payout)]
    y = f.year_labels
    heads = [f"Last {h5} years ({y[last - h5][:4]}-{y[last][:4]})", f"Last {h3} years ({y[last - h3][:4]}-{y[last][:4]})",
             f"Forecast ({y[last][:4]}-{fc.year_labels[-1][:4]})", "Forecast vs 3-yr (pp)"]
    rows = [[r.P("", "thl")] + [r.P(h, "th") for h in heads]]
    for lab, a5, a3, fv, _ in spec:
        gap = (fv - a3) * 100 if fv is not None and a3 is not None else None
        gap_s = "n/a" if gap is None else ("0.0" if abs(gap) < 0.05 else f"{gap:+.1f}")
        rows.append([r.P(lab, "tdl"), r.P(fmt_pct(a5), "td"), r.P(fmt_pct(a3), "td"), r.P(fmt_pct(fv), "tdb"),
                     r.P(gap_s, "td")])
    w = CONTENT_W
    t = data_table(r.theme, rows, [w * 0.3, w * 0.17, w * 0.17, w * 0.18, w * 0.18], estimate_cols=(3,))
    out.append(r.ex.make("Model forecast against the company's own record", t,
                         "Screener.in (reported), model estimates (base case). CAGRs from the last reported year; "
                         "averages of yearly ratios. No consensus estimates are used."))
    labels = m["labels"]
    p1 = r.charts.bars("fvh_growth", labels[1:], [x or 0 for x in m["sales_growth"][1:]], HALF_W, CH_H, n_actual=n - 1,
                       pct=True)
    p2 = r.charts.lines("fvh_margin", labels, [("EBITDA margin", m["ebitda_margin"]), ("ROE", m["roe"])], HALF_W, CH_H,
                        split=n, legend_cols=2)
    out.append(pair(r.ex.make("Revenue growth, reported | forecast", r.img(p1, HALF_W, CH_H),
                              "Screener.in, model estimates. Dark = reported, light = forecast."),
                    r.ex.make("EBITDA margin and ROE, reported | forecast", r.img(p2, HALF_W, CH_H))))
    out.append(Spacer(1, 2))
    return out
