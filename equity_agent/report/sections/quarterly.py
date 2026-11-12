"""Quarterly results: last 8 quarters with year-on-year and quarter-on-quarter change, and year to date vs FY1E."""
from __future__ import annotations

import re
from typing import Dict, List, Optional

from reportlab.platypus import Spacer

from ..common import R, pair
from ..layout import section_title
from ..styles import CONTENT_W, HALF_W, RS, data_table, fmt_num, fmt_pct

KEYS = ("sales", "expenses", "ebitda", "other_income", "interest", "depreciation", "pbt", "net_profit", "eps")
_LAB = re.compile(r"([A-Za-z]{3})\w*\s+(\d{4})")


def short_quarter(label: str) -> str:
    """'Sep 2026' -> 'Sep 26'."""
    m = _LAB.match(label.strip())
    return f"{m.group(1)} {m.group(2)[2:]}" if m else label


def _prior_year(label: str) -> Optional[str]:
    m = _LAB.match(label.strip())
    return f"{m.group(1)} {int(m.group(2)) - 1}" if m else None


def last_quarters(fin, n: int = 8) -> Optional[Dict[str, List]]:
    """The last ``n`` quarters with a revenue figure, plus YoY / QoQ growth looked up across the full history."""
    q, labs = fin.quarters, fin.quarter_labels
    if not q or not labs or "sales" not in q:
        return None
    full = {lab: i for i, lab in enumerate(labs)}
    keep = [i for i, lab in enumerate(labs) if i < len(q["sales"]) and q["sales"][i] is not None][-n:]
    if not keep:
        return None
    get = lambda key, i: (q[key][i] if key in q and i is not None and i < len(q[key]) else None)
    out: Dict[str, List] = {"labels": [labs[i] for i in keep]}
    for key in KEYS:
        out[key] = [get(key, i) for i in keep]

    def chg(key, i, j):
        a, b = get(key, j), get(key, i)
        return b / a - 1 if a and a > 0 and b is not None else None

    for key in ("sales", "ebitda", "net_profit"):
        out[f"{key}_yoy"] = [chg(key, i, full.get(_prior_year(labs[i]))) for i in keep]
        out[f"{key}_qoq"] = [chg(key, i, i - 1 if i > 0 else None) for i in keep]
    return out


def quarterly_results(r: R, num: int) -> list:
    res = r.res
    f, fc = res.fin, res.forecast
    out = section_title(r.theme, "Quarterly results", num)
    q = last_quarters(f, 8)
    if q is None:
        out.append(r.P("Screener.in shows no quarterly results for this company.", "small"))
        return out
    margin = lambda a, b: [x / s if s and x is not None else None for x, s in zip(q[a], q[b])]
    pg = lambda x: fmt_pct(x, 1, True)
    spec = [("Revenue", q["sales"], fmt_num), ("  YoY", q["sales_yoy"], pg), ("  QoQ", q["sales_qoq"], pg),
            ("Operating expenses", q["expenses"], fmt_num), ("EBITDA", q["ebitda"], fmt_num),
            ("  EBITDA margin", margin("ebitda", "sales"), fmt_pct), ("  YoY", q["ebitda_yoy"], pg),
            ("Other income", q["other_income"], fmt_num), ("Depreciation", q["depreciation"], fmt_num),
            ("Interest", q["interest"], fmt_num), ("Profit before tax", q["pbt"], fmt_num),
            ("Net profit", q["net_profit"], fmt_num), ("  net margin", margin("net_profit", "sales"), fmt_pct),
            ("  YoY", q["net_profit_yoy"], pg), ("  QoQ", q["net_profit_qoq"], pg),
            (f"EPS ({RS})", q["eps"], lambda x: fmt_num(x, 2))]
    heads = [short_quarter(x) for x in q["labels"]]
    rows = [[r.P(f"{RS} Cr", "thl")] + [r.P(h, "th") for h in heads]]
    tot = []
    for k, (lab, vals, fmt) in enumerate(spec, 1):
        bold = lab in ("EBITDA", "Net profit")
        if bold:
            tot.append(k)
        rows.append([r.P(f"<b>{lab}</b>" if bold else lab, "tdl")] +
                    [r.P(fmt(x) if x is not None else "n/a", "tdb" if bold else "td") for x in vals])
    lw = CONTENT_W * 0.2
    t = data_table(r.theme, rows, [lw] + [(CONTENT_W - lw) / len(heads)] * len(heads), total_rows=tot)
    out.append(r.ex.make(f"Last {len(heads)} quarters ({RS} Cr)", t,
                         "Screener.in quarterly results (" + ("consolidated" if f.consolidated else "standalone") +
                         "). YoY = same quarter a year earlier."))

    q12 = last_quarters(f, 12)
    labs = [short_quarter(x) for x in q12["labels"]]
    p = r.charts.lines("q_growth", labs, [("Revenue YoY", q12["sales_yoy"]), ("EBITDA YoY", q12["ebitda_yoy"]),
                                          ("Net profit YoY", q12["net_profit_yoy"])], HALF_W, 150)
    left = r.ex.make("Year-on-year growth by quarter", r.img(p, HALF_W, 150), "Screener.in quarterly results")

    # year to date vs the first forecast year
    fy1 = fc.year_labels[0]
    rows = [[r.P("", "thl"), r.P("Year to date", "th"), r.P(f"{fy1} forecast", "th"), r.P("Gap (pp)", "th")]]
    fc_growth = {"sales": fc.lines["sales"][0] / f.lines["sales"][-1] - 1 if f.lines["sales"][-1] else None,
                 "net_profit": (fc.lines["net_profit"][0] / f.lines["net_profit"][-1] - 1
                                if f.lines["net_profit"][-1] > 0 else None)}
    nq = 0
    for key, lab in (("sales", "Revenue growth"), ("net_profit", "Net profit growth")):
        n_q, g = f.ytd_growth(key)
        nq = max(nq, n_q)
        gap = (g - fc_growth[key]) * 100 if g is not None and fc_growth[key] is not None else None
        rows.append([r.P(lab, "tdl"), r.P(fmt_pct(g, 1, True), "td"), r.P(fmt_pct(fc_growth[key], 1, True), "td"),
                     r.P(f"{gap:+.1f}" if gap is not None else "n/a", "td")])
    widths = [HALF_W * 0.43, HALF_W * 0.19, HALF_W * 0.2, HALF_W * 0.18]
    parts = [data_table(r.theme, rows, widths)]
    note = (f"Year to date = the {nq} reported quarter{'s' if nq != 1 else ''} of {fy1[:-1]} vs the same quarters a "
            "year earlier. A gap above 5 pp is flagged in the model checks." if nq else
            f"No quarter of {fy1[:-1]} is reported yet.")
    parts += [Spacer(1, 2), r.P(note, "small")]
    ttm = f.ttm or {}
    if ttm.get("sales"):
        rows = [[r.P(f"{RS} Cr", "thl"), r.P("Trailing 12m", "th"), r.P(f.year_labels[-1], "th"), r.P("Change", "th")]]
        for key, lab in (("sales", "Revenue"), ("ebitda", "EBITDA"), ("net_profit", "Net profit")):
            last, tv = f.lines[key][-1], ttm.get(key)
            rows.append([r.P(lab, "tdl"), r.P(fmt_num(tv), "td"), r.P(fmt_num(last), "td"),
                         r.P(fmt_pct(tv / last - 1, 1, True) if tv is not None and last else "n/a", "td")])
        parts += [Spacer(1, 6), data_table(r.theme, rows, widths)]
    right = r.ex.make(f"Year to date vs the {fy1} forecast, and trailing twelve months", parts,
                      "Screener.in quarterly results and TTM column, model estimates")
    out.append(pair(left, right))
    return out
