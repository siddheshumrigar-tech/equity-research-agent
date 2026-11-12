"""Sections written from the host agent's ``narrative.json`` (and sourced ``analyst_inputs.json`` items).

Every function returns ``[]`` when its content is missing, before creating any exhibit, so the
builder can skip the section without using a section number. Drafts (no narrative) show none of these.
"""
from __future__ import annotations

import re
from datetime import datetime
from typing import List, Optional

from reportlab.platypus import KeepTogether, Paragraph, Spacer, Table, TableStyle

from ..common import R, pair
from ..layout import section_title
from ..metrics import key_metrics
from ..styles import CONTENT_W, HALF_W, MUTED, RS, c, data_table, fmt_pct, tint
from ...analyst_inputs import source_text
from ...narrative import PESTLE, PORTER_FORCES, PORTER_LABELS, SWOT

RATING_COL = {"low": "15803D", "moderate": "B45309", "high": "B91C1C"}
SWOT_LABELS = {"strengths": "Strengths", "weaknesses": "Weaknesses", "opportunities": "Opportunities",
               "threats": "Threats"}
IMPACT_COL = {"positive": "15803D", "negative": "B91C1C", "neutral": MUTED}


def _has(r: R, key: str) -> bool:
    return r.nar is not None and r.nar.has(key)


def _cites(r: R, ids) -> str:
    ids = [int(i) for i in (ids or []) if str(i).isdigit()]
    return r.N(f" [{', '.join(str(i) for i in ids)}]") if ids else ""


def _paras(r: R, text, style: str = "body") -> list:
    """Narrative text split on blank lines into paragraphs."""
    return [Paragraph(r.N(p.strip()), r.st[style]) for p in re.split(r"\n\s*\n", str(text)) if p.strip()]


def _src_names(r: R, texts: List[str]) -> str:
    """'Narrative sources [1], [2]' for the source line of an exhibit, from the markers used in its text."""
    ids = sorted({int(n) for t in texts for m in re.finditer(r"\[(\d+(?:\s*,\s*\d+)*)\]", str(t or ""))
                  for n in re.split(r"\s*,\s*", m.group(1))})
    if not ids:
        return "Analyst narrative"
    titles = {s["id"]: s.get("title", "") for s in r.nar.sources}
    return "; ".join(f"[{i}] {titles.get(i, '')}".strip() for i in ids)


# --------------------------------------------------------------------------- #
# investment thesis
# --------------------------------------------------------------------------- #
THESIS_EXHIBITS = ("rev_ebitda", "margins", "fcf", "roe", "wc_days", "peers_pe", "price")


def _thesis_chart(r: R, key: str, w: float, h: float, name: str):
    """(png path, exhibit title, source) for a thesis exhibit keyword; (None, ..) when the data is missing."""
    res = r.res
    m = key_metrics(res)
    labels, n = m["labels"], m["n_actual"]
    ch, v, f = r.charts, res.valuation, res.fin
    if key == "rev_ebitda":
        return (ch.bars_line(name, labels, m["sales"], m["ebitda_margin"], n, w, h, f"Revenue ({RS} Cr)", "EBITDA margin"),
                "Revenue and EBITDA margin", "Screener.in, model estimates")
    if key == "margins":
        return (ch.lines(name, labels, [("EBITDA margin", m["ebitda_margin"]), ("Net margin", m["net_margin"])], w, h,
                         split=n, legend_cols=2), "EBITDA and net margin", "Screener.in, model estimates")
    if key == "fcf":
        fc = res.forecast
        cfo = list(f.lines["cfo"]) + list(fc.lines["cfo"])
        capex = list(f.lines["capex"]) + list(fc.lines["capex"])
        return (ch.grouped_bars(name, labels, [("Cash from operations", cfo),
                                               ("CFO - capex", [a - b for a, b in zip(cfo, capex)])], w, h, n_actual=n),
                f"Cash generation ({RS} Cr)", "Screener.in cash flow, model estimates")
    if key == "roe":
        return (ch.lines(name, labels, [("ROE", m["roe"]), ("ROCE (pre-tax)", m["roce"])], w, h, split=n, legend_cols=2,
                         hlines=[(v.wacc, f"WACC {v.wacc:.1%}")]), "Returns vs the cost of capital",
                "Screener.in, model")
    if key == "wc_days":
        L = f.lines
        dso = [x / s * 365 if s else None for x, s in zip(L["receivables"], L["sales"])]
        dpo = [x / b * 365 if b and b > 0 else None for x, b in zip(L["trade_payables"], L["cogs"])]
        dio = [x / b * 365 if b and b > 0 else None for x, b in zip(L["inventory"], L["cogs"])]
        return (ch.lines(name, f.year_labels, [("Receivable days", dso), ("Inventory days", dio), ("Payable days", dpo)],
                         w, h, pct=False, legend_cols=3), "Working-capital days (reported)", "Screener.in")
    if key == "peers_pe" and res.peers is not None:
        from ...peers import subject_snapshot

        rows = [subject_snapshot(f)] + list(res.peers.peers)
        names = [p.name.replace(" Ltd", "").replace(" Limited", "")[:26] for p in rows]
        return (ch.peer_bars(name, names, [p.pe for p in rows], 0, res.peers.median("pe"), w, h, unit="x"),
                "Trailing P/E vs peers", "Screener.in peer pages")
    if key == "price" and res.market and res.market.price_close:
        d, cl, ix = res.market.last_year()
        return (ch.price_vs_index(name, d, cl, ix, w, h, label=f.symbol), "12-month price vs the Nifty 50 (rebased)",
                "yfinance")
    return None, "", ""


def thesis(r: R, num: int) -> list:
    if not _has(r, "thesis_points"):
        return []
    out = section_title(r.theme, "Investment thesis", num)
    out.append(Paragraph(r.N(r.nar.get("headline")), r.st["h2"]))
    for k, tp in enumerate(r.nar.get("thesis_points")):
        title = Paragraph(f"{k + 1}. {r.N(tp['title'])}", r.st["h3"])
        body = _paras(r, tp["text"])
        key = str(tp.get("exhibit") or "").strip()
        p = None
        if key in THESIS_EXHIBITS:
            wc, hc = HALF_W * 0.92, 140
            p, ex_title, src = _thesis_chart(r, key, wc, hc, f"thesis_{k}")
        if p and len(str(tp["text"])) <= 900:
            chart = r.ex.make(ex_title, r.img(p, wc, hc), src)
            out.append(KeepTogether([title, pair(body, chart, w_left=0.52)]))
        elif p:  # long text: full width, chart below (a table cell cannot split across pages)
            out += [title] + body + [r.ex.make(ex_title, r.img(p, wc, hc), src)]
        else:
            out.append(KeepTogether([title] + body))
        out.append(Spacer(1, 3))
    return out


# --------------------------------------------------------------------------- #
# company overview, segments
# --------------------------------------------------------------------------- #
def _segment_sources(segs) -> str:
    """Unique source titles with their dates; URLs are listed on the sources page of analyst_inputs.json."""
    seen = {}
    for s in segs:
        src = s["source"]
        seen.setdefault(src.get("title", ""), src.get("date", ""))
    return "; ".join(f"{t} ({d})" for t, d in seen.items())


def company_overview(r: R, num: int) -> list:
    keys = [k for k in ("company_overview", "business_model", "management") if _has(r, k)]
    ai = r.res.analyst
    segs = list(ai.segments) if ai is not None else []
    if not keys and not segs:
        return []
    out = section_title(r.theme, "Company overview", num)
    heads = {"company_overview": "What the company does", "business_model": "Business model",
             "management": "Management and governance"}
    text = []
    for k in keys:
        text.append(Paragraph(heads[k], r.st["h3"]))
        text += _paras(r, r.nar.get(k))
    if not segs:
        out += text
        return out
    out += text
    side = []
    for key, lab in (("revenue_share", "Revenue"), ("ebit_share", "EBIT")):
        vals = [s.get(key) for s in segs]
        if not any(vals):
            continue
        period = segs[0].get("period", "")
        p = r.charts.donut(f"seg_{key}", [s["name"] for s in segs], [x or 0 for x in vals], HALF_W, 112, lab)
        if p:
            side.append(r.ex.make(f"Segment share of {lab if lab == 'EBIT' else lab.lower()}"
                                  f"{' (' + period + ')' if period else ''} - analyst input",
                                  r.img(p, HALF_W, 112), _segment_sources(segs)))
    if len(side) == 2:
        out.append(pair(side[0], side[1]))
    else:
        out += side
    return out


# --------------------------------------------------------------------------- #
# industry: text, market share, Porter, SWOT
# --------------------------------------------------------------------------- #
def _porter_table(r: R, porter: dict):
    rows = [[r.P("Force", "thl"), r.P("Threat", "thl"), r.P("Evidence", "thl")]]
    styles = []
    for i, k in enumerate([k for k in PORTER_FORCES if k in porter], 1):
        v = porter[k]
        col = RATING_COL[v["rating"]]
        rows.append([r.P(PORTER_LABELS[k], "tdlb"), r.P(f"<font color='#{col}'><b>{v['rating'].title()}</b></font>", "tdl"),
                     r.P(r.N(v["evidence"]), "tdl")])
        styles.append(("BACKGROUND", (1, i), (1, i), c(tint(col, 0.88))))
    t = data_table(r.theme, rows, [CONTENT_W * 0.24, CONTENT_W * 0.1, CONTENT_W * 0.66])
    t.setStyle(TableStyle(styles + [("VALIGN", (0, 0), (-1, -1), "TOP")]))
    return t


def _swot_grid(r: R, swot: dict):
    cols = {"strengths": "15803D", "weaknesses": "B45309", "opportunities": r.theme.accent, "threats": "B91C1C"}
    w = (CONTENT_W - 8) / 2

    def cell(k):
        items = swot.get(k) or []
        parts = [Paragraph(f"<font color='#{cols[k]}'><b>{SWOT_LABELS[k]}</b></font>", r.st["h3"])]
        parts += [Paragraph(r.N(x), r.st["bul"], bulletText="•") for x in items] or [r.P("Not provided.", "small")]
        return parts

    t = Table([[cell("strengths"), cell("weaknesses")], [cell("opportunities"), cell("threats")]], colWidths=[w, w])
    st = [("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
          ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 5)]
    for (rr, cc), k in {(0, 0): "strengths", (1, 0): "weaknesses", (0, 1): "opportunities", (1, 1): "threats"}.items():
        st += [("BACKGROUND", (rr, cc), (rr, cc), c(tint(cols[k], 0.93))),
               ("LINEABOVE", (rr, cc), (rr, cc), 1.4, c(cols[k]))]
    t.setStyle(TableStyle(st))
    return t


def _market_share(r: R):
    ms = [x for x in (r.nar.get("market_share") or []) if isinstance(x, dict) and x.get("share") is not None]
    if not ms:
        return None
    f = r.res.fin
    first = f.name.split()[0].lower()
    hi = next((i for i, x in enumerate(ms) if str(x.get("company", "")).lower().startswith(first)
               or str(x.get("company", "")).upper() == f.symbol), -1)
    rows = [[r.P("Company", "thl"), r.P("Share", "th"), r.P("Measure", "thl")]]
    for x in ms:
        rows.append([r.P(r.N(str(x["company"])), "tdl"), r.P(fmt_pct(x["share"], 1), "td"),
                     r.P(r.N(str(x.get("metric", ""))) + _cites(r, x.get("source_ids")), "small")])
    t = data_table(r.theme, rows, [HALF_W * 0.38, HALF_W * 0.16, HALF_W * 0.46])
    h = 30 + 15 * len(ms)
    p = r.charts.peer_bars("mkt_share", [str(x["company"])[:26] for x in ms], [x["share"] for x in ms], hi, None,
                           HALF_W, h, unit="pct")
    ids = sorted({int(i) for x in ms for i in (x.get("source_ids") or []) if str(i).isdigit()})
    titles = {s["id"]: s.get("title", "") for s in r.nar.sources}
    src = "; ".join(f"[{i}] {titles.get(i, '')}" for i in ids) or "Analyst narrative"
    return pair(r.ex.make("Market share", t, src), r.ex.make("Market share by company", r.img(p, HALF_W, h), src))


def industry(r: R, num: int) -> list:
    if not any(_has(r, k) for k in ("industry", "market_share", "porter", "swot")):
        return []
    nar = r.nar
    out = section_title(r.theme, "Industry and competitive position", num)
    if _has(r, "industry"):
        out += _paras(r, nar.get("industry"))
    ms = _market_share(r)
    if ms is not None:
        out.append(ms)
    porter = {k: v for k, v in (nar.get("porter") or {}).items() if k in PORTER_FORCES}
    if porter:
        n_hi = sum(1 for v in porter.values() if v["rating"] == "high")
        out.append(r.ex.make("Porter's five forces: threat to profitability", _porter_table(r, porter),
                             _src_names(r, [v["evidence"] for v in porter.values()]) +
                             f". {n_hi} of {len(porter)} forces rated high."))
    swot = {k: v for k, v in (nar.get("swot") or {}).items() if k in SWOT}
    if swot:
        out.append(r.ex.make("SWOT analysis", _swot_grid(r, swot),
                             _src_names(r, [x for v in swot.values() for x in v])))
    return out


# --------------------------------------------------------------------------- #
# PESTLE and ESG
# --------------------------------------------------------------------------- #
def pestle_esg(r: R, num: int) -> list:
    pe = {k: v for k, v in ((r.nar.get("pestle") if r.nar else None) or {}).items() if k in PESTLE and str(v).strip()}
    if not pe and not _has(r, "esg"):
        return []
    out = section_title(r.theme, "PESTLE and ESG", num)
    if pe:
        rows = [[r.P("Factor", "thl"), r.P("What it means for the company", "thl")]]
        for k in PESTLE:
            if k in pe:
                rows.append([r.P(k.title(), "tdlb"), r.P(r.N(pe[k]), "tdl")])
        t = data_table(r.theme, rows, [CONTENT_W * 0.16, CONTENT_W * 0.84])
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
        out.append(r.ex.make("PESTLE analysis", t, _src_names(r, list(pe.values()))))
    if _has(r, "esg"):
        out.append(Paragraph("Environmental, social and governance", r.st["h3"]))
        out += _paras(r, r.nar.get("esg"))
    return out


# --------------------------------------------------------------------------- #
# risks and catalysts
# --------------------------------------------------------------------------- #
def catalyst_date(s: str) -> Optional[datetime]:
    """YYYY-MM-DD, YYYY-MM, FYxx (31 Mar) or 'Qn FYxx' (Indian fiscal quarters: Q1 ends 30 Jun of FYxx-1)."""
    s = str(s).strip()
    for fmt in ("%Y-%m-%d", "%Y-%m"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            pass
    m = re.match(r"^(?:Q([1-4]) )?FY(\d{2})$", s)
    if not m:
        return None
    fy = 2000 + int(m.group(2))
    if not m.group(1):
        return datetime(fy, 3, 31)
    q = int(m.group(1))
    return {1: datetime(fy - 1, 6, 30), 2: datetime(fy - 1, 9, 30), 3: datetime(fy - 1, 12, 31),
            4: datetime(fy, 3, 31)}[q]


def risks_catalysts(r: R, num: int) -> list:
    risks = [x for x in ((r.nar.get("risks") if r.nar else None) or []) if isinstance(x, dict)]
    ai = r.res.analyst
    cats = list(ai.catalysts) if ai is not None else []
    if r.nar is None or (not risks and not cats):
        return []
    v = r.res.valuation
    out = section_title(r.theme, "Risks and catalysts", num)
    if risks:
        rows = [[r.P("Risk", "thl"), r.P("What happens", "thl"), r.P("Impact on target", "th"),
                 r.P(f"Target if it happens ({RS})", "th"), r.P("Mitigant", "thl")]]
        for x in risks:
            pct = x.get("impact_on_tp_pct")
            pct = float(pct) if isinstance(pct, (int, float)) else None
            rows.append([r.P(r.N(x["risk"]), "tdlb"), r.P(r.N(x["impact"]), "tdl"),
                         r.P(fmt_pct(pct, 0, True) if pct is not None else "n/a", "td"),
                         r.P(f"{v.target_price * (1 + pct):,.0f}" if pct is not None else "n/a", "td"),
                         r.P(r.N(x.get("mitigant", "")) or "-", "tdl")])
        w = CONTENT_W
        t = data_table(r.theme, rows, [w * 0.17, w * 0.35, w * 0.11, w * 0.12, w * 0.25])
        t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP")]))
        out.append(r.ex.make("Key risks: what happens, effect on the target, mitigants", t,
                             _src_names(r, [x.get("impact", "") for x in risks] + [x.get("mitigant", "") for x in risks]) +
                             f". Target if it happens = {RS}{v.target_price:,.0f} x (1 + impact), the analyst's estimate."))
        sized = [(str(x["risk"])[:34], float(x["impact_on_tp_pct"])) for x in risks
                 if isinstance(x.get("impact_on_tp_pct"), (int, float))]
        if len(sized) >= 2:
            h = 30 + 16 * len(sized)
            p = r.charts.impact_bars("risk_impact", [a for a, _ in sized], [b for _, b in sized], CONTENT_W * 0.6, h)
            out.append(r.ex.make("Estimated impact of each risk on the 12-month target", r.img(p, CONTENT_W * 0.6, h),
                                 "Analyst estimates in the narrative"))
    if cats:
        rows = [[r.P("Date", "thl"), r.P("Event", "thl"), r.P("Expected impact", "thl"), r.P("Source", "thl")]]
        items = []
        for x in sorted(cats, key=lambda x: catalyst_date(x.get("date")) or datetime.max):
            imp = str(x.get("impact", "neutral")).lower()
            col = IMPACT_COL.get(imp, MUTED)
            rows.append([r.P(str(x["date"]), "tdl"), r.P(r.N(x["event"]), "tdl"),
                         r.P(f"<font color='#{col}'>{imp.title()}</font>", "tdl"),
                         r.P(f"{x['source'].get('title', '')} ({x['source'].get('date', '')})", "small")])
            d = catalyst_date(x.get("date"))
            if d is not None:
                lab = str(x["event"])
                items.append((d, lab if len(lab) <= 40 else lab[:38].rstrip() + "...", imp))
        w = CONTENT_W
        parts = []
        if len(items) >= 2:
            p = r.charts.timeline("catalysts", items, CONTENT_W, 92)
            if p:
                parts.append(r.img(p, CONTENT_W, 92))
        parts.append(data_table(r.theme, rows, [w * 0.12, w * 0.38, w * 0.12, w * 0.38]))
        out.append(r.ex.make("Catalyst timeline (analyst input)", parts, "analyst_inputs.json, each item sourced"))
    return out


# --------------------------------------------------------------------------- #
# corporate actions
# --------------------------------------------------------------------------- #
def corporate_actions(r: R, num: int) -> list:
    acts = [x for x in ((r.nar.get("corporate_actions") if r.nar else None) or []) if isinstance(x, dict)
            and str(x.get("event", "")).strip()]
    if not acts:
        return []
    out = section_title(r.theme, "Corporate actions, M&amp;A and capital allocation", num)
    rows = [[r.P("Date", "thl"), r.P("Event", "thl")]]
    for x in sorted(acts, key=lambda x: str(x.get("date", ""))):
        rows.append([r.P(str(x.get("date", "")), "tdl"), r.P(r.N(x["event"]), "tdl")])
    t = data_table(r.theme, rows, [CONTENT_W * 0.14, CONTENT_W * 0.86])
    out.append(r.ex.make("Corporate actions", t, _src_names(r, [x["event"] for x in acts])))
    return out


__all__ = ["thesis", "company_overview", "industry", "pestle_esg", "risks_catalysts", "corporate_actions",
           "catalyst_date", "THESIS_EXHIBITS"]
