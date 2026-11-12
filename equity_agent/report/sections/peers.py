"""Peer comparison table and a ratios-vs-peers page (one row per ratio: values, rank, comment, mini chart)."""
from __future__ import annotations

import statistics
from typing import List, Optional

from reportlab.platypus import Spacer

from ..common import R
from ..layout import section_title
from ..styles import CONTENT_W, MUTED, RS, c, data_table, fmt_num, fmt_pct, fmt_x
from ...peers import EV_EBITDA_RANGE, PE_RANGE, subject_snapshot

# key, label, unit ("pct" = fraction, "pctpt" = already in %, "x"), kind ("quality" = higher is better,
# "multiple" = higher is more expensive)
RATIOS = [
    ("sales_growth_3y", "Revenue growth, 3-yr CAGR", "pct", "quality"),
    ("ebitda_margin", "EBITDA margin", "pct", "quality"),
    ("roe", "Return on equity", "pctpt", "quality"),
    ("roce", "Return on capital employed", "pctpt", "quality"),
    ("pe", "P/E (trailing)", "x", "multiple"),
    ("ev_ebitda", "EV/EBITDA (trailing)", "x", "multiple"),
]
_RANGE = {"pe": PE_RANGE, "ev_ebitda": EV_EBITDA_RANGE}


def _plain(name: str) -> str:
    """Company name without the legal suffix (for charts)."""
    for tail in (" Ltd", " Limited", " India", " (India)"):
        if name.endswith(tail):
            name = name[: -len(tail)]
    return name


def _short(name: str) -> str:
    """Company name for a Paragraph: suffix removed and XML-escaped ('L&T' would otherwise render as 'L&T;')."""
    from xml.sax.saxutils import escape

    return escape(_plain(name))


def _fmt(unit: str, v: Optional[float]) -> str:
    if v is None:
        return "n/a"
    return fmt_x(v) if unit == "x" else (fmt_pct(v) if unit == "pct" else f"{v:.1f}%")


def _valid(key: str, v: Optional[float]) -> bool:
    if v is None:
        return False
    lo, hi = _RANGE.get(key, (-1e9, 1e9))
    return lo < v < hi


def ratio_comment(key: str, unit: str, kind: str, subj: Optional[float], peer_vals: List[float]) -> str:
    """One computed sentence comparing the subject with its peers."""
    if subj is None or not _valid(key, subj):
        return "Not meaningful for the company (missing or outside the range used for peers)."
    if len(peer_vals) < 2:
        return "Too few peers with this ratio to compare."
    med = statistics.median(peer_vals)
    allv = sorted(peer_vals + [subj], reverse=True)
    rank = allv.index(subj) + 1
    n = len(allv)
    if kind == "multiple":
        if not med:
            return f"Peer median is {_fmt(unit, med)}."
        prem = subj / med - 1
        side = "premium" if prem >= 0 else "discount"
        where = ("the most expensive" if rank == 1 else "the cheapest" if rank == n
                 else f"the {_ordinal(rank)} most expensive")
        return f"A {abs(prem):.0%} {side} to the peer median of {_fmt(unit, med)}; {where} of {n}."
    gap = (subj - med) * (100 if unit == "pct" else 1)
    side = "above" if gap >= 0 else "below"
    where = "the highest" if rank == 1 else "the lowest" if rank == n else f"the {_ordinal(rank)} highest"
    return f"{abs(gap):.1f} pp {side} the peer median of {_fmt(unit, med)}; {where} of {n}."


def _ordinal(k: int) -> str:
    suffix = "th" if 10 <= k % 100 <= 20 else {1: "st", 2: "nd", 3: "rd"}.get(k % 10, "th")
    return f"{k}{suffix}"


def _group(r: R):
    ps = r.res.peers
    subj = subject_snapshot(r.res.fin)
    peers = sorted(ps.peers, key=lambda p: -(p.market_cap or 0))
    return subj, peers


def peer_section(r: R, num: int) -> list:
    res = r.res
    out = section_title(r.theme, "Peer comparison", num)
    ps = res.peers
    if ps is None or not ps.peers:
        out.append(r.P("No peer data is available for this company; the EV/EBITDA method is left out of the target.",
                       "small"))
        return out
    subj, peers = _group(r)
    cols = [("Company", None), (f"Mcap ({RS} Cr)", "market_cap"), (f"Revenue ({RS} Cr)", "sales"),
            ("Rev. CAGR 3y", "sales_growth_3y"), ("EBITDA margin", "ebitda_margin"), ("Net margin", "net_margin"),
            ("ROE", "roe"), ("ROCE", "roce"), ("P/E", "pe"), ("EV/ EBITDA", "ev_ebitda"), ("P/B", "pb"),
            ("Div. yield", "div_yield")]
    unit = {"market_cap": "num", "sales": "num", "sales_growth_3y": "pct", "ebitda_margin": "pct", "net_margin": "pct",
            "roe": "pctpt", "roce": "pctpt", "pe": "x", "ev_ebitda": "x", "pb": "x", "div_yield": "pctpt"}

    def cell(key, v):
        if key is None:
            return v
        if v is None or (key in _RANGE and not _valid(key, v)):
            return "n/a"
        u = unit[key]
        return fmt_num(v) if u == "num" else _fmt(u, v)

    rows = [[r.P(c, "thl" if k is None else "th") for c, k in cols]]
    for p in [subj] + peers:
        bold = p is subj
        rows.append([r.P(f"<b>{_short(p.name)}</b>" if bold else _short(p.name), "tdl")] +
                    [r.P(cell(k, getattr(p, k)), "tdb" if bold else "td") for _, k in cols[1:]])
    med_row = [r.P("<b>Peer median</b>", "tdl")]
    for _, k in cols[1:]:
        vals = [getattr(p, k) for p in peers if getattr(p, k) is not None and (k not in _RANGE or _valid(k, getattr(p, k)))]
        med_row.append(r.P(cell(k, statistics.median(vals)) if len(vals) >= 2 else "n/a", "tdb"))
    rows.append(med_row)
    w = CONTENT_W
    lw = w * 0.17
    t = data_table(r.theme, rows, [lw] + [(w - lw) / (len(cols) - 1)] * (len(cols) - 1), section_rows=(1,),
                   total_rows=(len(rows) - 1,))
    basis = sorted({p.basis for p in [subj] + peers if p.basis})
    out.append(r.ex.make("Peer comparison on one set of definitions", t,
                         f"{ps.source}, fetched {ps.fetched_at}. Margins and EV/EBITDA on "
                         f"{' / '.join(basis) or 'latest'} figures; EV = market cap + borrowings - cash - investments. "
                         "P/E above 200x and EV/EBITDA above 100x are shown as n/a and left out of medians."))
    out += ratios_vs_peers(r, subj, peers)
    out += peer_history_section(r, subj, peers)
    return out


HIST_LABELS = {"sales_growth": "Revenue growth", "ebitda_margin": "EBITDA margin", "net_margin": "Net margin",
               "roe": "ROE", "roce": "ROCE (pre-tax)"}


def years_above(subj_vals: List[Optional[float]], med_vals: List[Optional[float]]):
    """(years the company beat the peer median, years compared)."""
    pairs = [(a, b) for a, b in zip(subj_vals, med_vals) if a is not None and b is not None]
    return sum(1 for a, b in pairs if a > b), len(pairs)


def peer_history_section(r: R, subj, peers) -> list:
    """Multi-year ratios: the company against the peer median, year by year, on one set of definitions."""
    from reportlab.platypus import CondPageBreak

    ps = r.res.peers
    if not ps.has_history:
        return [r.P("Peer history is not available (peer pages not cached); the comparison above is a snapshot.",
                    "small")]
    years = ps.history_years()
    if len(years) < 2:
        return []
    rows = [[r.P("", "thl"), r.P("", "thl")] + [r.P(y, "th") for y in years] + [r.P("Years above median", "th")]]
    sections = []
    panels = []
    for key in HISTORY_METRICS_ORDER:
        comp = [subj.history.get(key, {}).get(y) for y in years]
        med = [ps.history_median(key, y) for y in years]
        others = [[p.history.get(key, {}).get(y) for y in years] for p in peers if p.history]
        above, n = years_above(comp, med)
        gap = [None if a is None or b is None else (a - b) for a, b in zip(comp, med)]
        start = len(rows)
        sections.append(start)
        rows.append([r.P(f"<b>{HIST_LABELS[key]}</b>", "tdl"), r.P(_short(subj.name), "tdl")]
                    + [r.P(fmt_pct(v), "tdb") for v in comp] + [r.P(f"{above} of {n}" if n else "n/a", "tdb")])
        rows.append([r.P("", "tdl"), r.P("Peer median", "tdl")] + [r.P(fmt_pct(v), "td") for v in med] + [r.P("", "td")])
        rows.append([r.P("", "tdl"), r.P("Gap (pp)", "tdl")]
                    + [r.P("n/a" if g is None else ("0.0" if abs(g) < 0.0005 else f"{g * 100:+.1f}"), "td")
                       for g in gap] + [r.P("", "td")])
        panels.append((HIST_LABELS[key], comp, med, others))
    w = CONTENT_W
    fixed = w * 0.17 + w * 0.15 + w * 0.12
    yw = (w - fixed) / len(years)
    t = data_table(r.theme, rows, [w * 0.17, w * 0.15] + [yw] * len(years) + [w * 0.12])
    from reportlab.platypus import TableStyle

    t.setStyle(TableStyle([("LINEABOVE", (0, k), (-1, k), 0.6, c(MUTED)) for k in sections[1:]]))
    odd = sorted({p.year_end for p in peers if p.year_end and p.year_end != subj.year_end})
    odd_names = [_short(p.name) for p in peers if p.year_end and p.year_end != subj.year_end]
    note = (f" Fiscal years are labelled by the calendar year they end in; {', '.join(odd_names)} "
            f"report{'s' if len(odd_names) == 1 else ''} to a different month ({', '.join(odd)}) than the company "
            f"({subj.year_end})." if odd_names else "")
    src = ("Screener.in annual P&amp;L and balance sheets of the company and each peer. ROE = net profit / average "
           "(equity capital + reserves); ROCE = (PBT + interest) / average (equity + borrowings); medians need at "
           "least two peers. Chart scales follow the company and the median; outlying peers may run off a panel."
           + note)
    ch_h = 128
    p = r.charts.peer_history("peer_hist", years, panels, CONTENT_W, ch_h, subject=_plain(subj.name))
    return [CondPageBreak(300), Spacer(1, 4),
            r.ex.make(f"{len(years)}-year ratios vs the peer median", t, src),
            r.ex.make("Ratio history: company, peer median and each peer", r.img(p, CONTENT_W, ch_h),
                      "Screener.in; same definitions as the table above")]


HISTORY_METRICS_ORDER = ("sales_growth", "ebitda_margin", "net_margin", "roe", "roce")


def ratios_vs_peers(r: R, subj, peers) -> list:
    names = [_plain(subj.name)] + [_plain(p.name) for p in peers]  # chart labels: plain text
    w = CONTENT_W
    widths = [w * 0.15, w * 0.085, w * 0.085, w * 0.13, w * 0.25, w * 0.30]
    rows = [[r.P("Ratio", "thl"), r.P("Company", "th"), r.P("Peer median", "th"), r.P("Peer range", "th"),
             r.P("Comment", "thl"), r.P("Company (coloured) vs peers; dashed line = peer median", "thl")]]
    ch_h = 14 + 9.5 * len(names)
    for key, label, unit, kind in RATIOS:
        s_val = getattr(subj, key)
        p_vals = [getattr(p, key) for p in peers]
        valid = [v for v in p_vals if _valid(key, v)]
        med = statistics.median(valid) if len(valid) >= 2 else None
        rng = f"{_fmt(unit, min(valid))} to {_fmt(unit, max(valid))}" if valid else "n/a"
        vals = [s_val if _valid(key, s_val) else None] + [v if _valid(key, v) else None for v in p_vals]
        p = r.charts.peer_bars(f"peer_{key}", names, vals, 0, med, widths[5] - 6, ch_h, unit=unit)
        rows.append([r.P(f"<b>{label}</b>", "tdl"), r.P(_fmt(unit, s_val) if _valid(key, s_val) else "n/a", "tdb"),
                     r.P(_fmt(unit, med), "td"), r.P(rng, "td"),
                     r.P(ratio_comment(key, unit, kind, s_val, valid), "small"), r.img(p, widths[5] - 6, ch_h)])
    t = data_table(r.theme, rows, widths)
    from reportlab.platypus import TableStyle

    t.setStyle(TableStyle([("VALIGN", (0, 1), (-1, -1), "MIDDLE")]))
    # The table may continue on the next page (header row repeats) instead of leaving a half-empty page: no
    # KeepTogether, and the title is not tied to the whole table (a tall table would move to the next page with it).
    parts = r.ex.make("Ratios vs peers", t,
                      "Screener.in (peers' headline ratios and P&amp;L; the company on the same definitions). Snapshot "
                      "at the latest reported period; comments are computed, not analyst views.", keep=False)
    parts[0].keepWithNext = False
    return [Spacer(1, 4)] + parts
