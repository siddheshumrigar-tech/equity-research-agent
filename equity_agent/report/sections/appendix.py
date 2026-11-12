"""Assumptions, analyst inputs, data notes, sources and disclosures."""
from __future__ import annotations

from reportlab.platypus import Paragraph, Spacer

from ..common import R
from ..layout import section_title
from ..styles import CONTENT_W, data_table, fmt_pct
from ...drivers import DRIVER_ORDER


def assumptions(r: R, num: int) -> list:
    res = r.res
    drv, fc = res.drivers, res.forecast
    out = section_title(r.theme, "Forecast assumptions", num)
    rows = [[r.P("Driver", "thl"), r.P("History", "th"), r.P("Base", "th"), r.P("Bull", "th"), r.P("Bear", "th"),
             r.P(fc.year_labels[0], "th"), r.P(fc.year_labels[-1], "th"), r.P("Basis", "thl")]]
    flagged = []
    for key in DRIVER_ORDER:
        d = drv.items[key]
        f = (lambda v: fmt_pct(v)) if d.unit == "pct" else (lambda v: f"{v:.0f}")
        y1 = drv.value(key, "base", 0) if key != "ronic_share" else d.base
        yn = drv.value(key, "base", len(fc.years) - 1) if key != "ronic_share" else d.base
        ov = key in drv.overrides
        lab = d.label + (" *" if ov else "")
        if ov:
            flagged.append(key)
        rows.append([r.P(lab, "tdl"), r.P(f(d.historical) if d.historical is not None else "n/a", "td"), r.P(f(d.base), "td"),
                     r.P(f(d.bull), "td"), r.P(f(d.bear), "td"), r.P(f(y1), "tdb" if ov else "td"),
                     r.P(f(yn), "tdb" if ov else "td"), r.P(d.basis + (" (clamped)" if d.clamped else ""), "small")])
    w = CONTENT_W
    out.append(r.ex.make("Forecast drivers (base case per year; * = analyst input)",
                         data_table(r.theme, rows, [w * 0.25, w * 0.08, w * 0.07, w * 0.07, w * 0.07, w * 0.07, w * 0.07, w * 0.32]),
                         "Screener.in history (3-year averages), config/market.json, analyst_inputs.json"))
    ai = res.analyst
    if ai is not None and ai.entries():
        rows = [[r.P("Analyst input", "thl"), r.P("Value", "th"), r.P("Model value", "th"), r.P("Rationale and source", "thl")]]
        for e in ai.entries():
            val = e["value"]
            vs = fmt_pct(val) if isinstance(val, float) and abs(val) < 5 else (f"{val:,.0f}" if isinstance(val, float) else str(val))
            mv = fmt_pct(e["model"]) if isinstance(e["model"], float) and abs(e["model"]) < 5 else (
                f"{e['model']:,.1f}" if isinstance(e["model"], float) else "")
            rows.append([r.P(e["item"], "tdl"), r.P(vs, "td"), r.P(mv, "td"),
                         r.P(f"{e['rationale']} - {e['source']}", "small")])
        out.append(r.ex.make("Analyst inputs applied (each with its source)",
                             data_table(r.theme, rows, [w * 0.24, w * 0.12, w * 0.12, w * 0.52]), "analyst_inputs.json"))
    return out


def data_notes(r: R, num: int) -> list:
    res = r.res
    out = section_title(r.theme, "Data notes, sources and methodology", num)
    st = r.st
    f = res.fin
    view = "consolidated" if f.consolidated else "standalone"
    out.append(Paragraph("Data sources", st["h2"]))
    lines = [f"Fundamentals: Screener.in {view} statements and schedules ({f.source_url}), fetched {f.fetched_at}.",
             f"Prices, beta, technicals: yfinance ({f.symbol}.NS vs Nifty 50), fetched {res.market.fetched_at or 'n/a'}.",
             f"Market inputs: config/market.json as of {res.cfg.get('as_of', 'n/a')}.",
             (f"Peers: {res.peers.source}, fetched {res.peers.fetched_at}." if res.peers else "Peers: not available."),
             f"Brand colour: {res.palette.source}."]
    from reportlab.lib.enums import TA_LEFT
    from reportlab.lib.styles import ParagraphStyle

    left = ParagraphStyle("bul_left", parent=st["bul"], alignment=TA_LEFT)  # long URLs: no justified gaps
    for s in lines:
        out.append(Paragraph(s, left, bulletText="-"))
    nar = res.narrative
    if nar is not None and nar.sources:
        out.append(Paragraph("Narrative sources (cited as [n])", st["h2"]))
        for s in nar.sources:
            ref = s.get("url") or s.get("document") or ""
            out.append(Paragraph(f"[{s['id']}] {s.get('title', '')} ({s.get('type', '')}, {s.get('date', '')}). {ref}",
                                 st["small"]))
    out.append(Paragraph("Methodology", st["h2"]))
    meth = [
        "Forecast: 3-statement model; drivers are 3-year historical averages unless a sourced analyst input replaces "
        "them; working-capital days glide from the last reported year; loans & advances held flat; CWIP capitalised "
        "evenly; minority share of profit at its 3-year average.",
        "DCF: FCFF for 5 explicit years, 5 fade years (growth to g, ROIC to RONIC, reinvestment = NOPAT x g / ROIC) and a "
        "value-driver terminal value; mid-year discounting; rolled forward 12 months at cost of equity minus dividend yield.",
        "Relative: mean fiscal-year-end P/E (actual traded prices, that year's share count) and peer median EV/EBITDA, "
        "each on next-twelve-months earnings.",
        "Rating on expected total return (upside + dividend yield): " + ", ".join(
            f"{lab} >= {t:+.0%}" for t, lab in res.cfg["rating_thresholds"]) + f", else {res.cfg.get('rating_floor', 'SELL')}.",
    ]
    for s in meth:
        out.append(Paragraph(s, st["bul"], bulletText="-"))
    if res.warnings:
        out.append(Paragraph("Data notes and warnings", st["h2"]))
        for wtxt in res.warnings:
            out.append(Paragraph(wtxt, st["small"]))
    return out


def _rating_rows(r: R) -> list:
    """Rating bands from config, highest first: [(label, 'above +15%'), ..., (floor, 'below -5%')]."""
    cfg = r.res.cfg
    bands = sorted(cfg["rating_thresholds"], key=lambda x: -x[0])
    rows, upper = [], None
    for t, lab in bands:
        rows.append((lab, f"{t:+.0%} or more" if upper is None else f"{t:+.0%} to {upper:+.0%}"))
        upper = t
    rows.append((cfg.get("rating_floor", "SELL"), f"below {upper:+.0%}"))
    return rows


def disclosures(r: R) -> list:
    from datetime import date

    from reportlab.platypus import KeepTogether

    st = r.st
    res = r.res
    v, f = res.valuation, res.fin
    out = [Spacer(1, 6), Paragraph("Disclosures", st["h2"])]
    rows = [[r.P("Rating", "thl"), r.P("12-month expected total return (price upside + dividend yield)", "thl")]]
    for lab, band in _rating_rows(r):
        rows.append([r.P(f"<b>{lab}</b>" + (" (this report)" if lab == v.rating else ""), "tdl"), r.P(band, "tdl")])
    table = data_table(r.theme, rows, [CONTENT_W * 0.25, CONTENT_W * 0.75])
    out.append(KeepTogether([Paragraph("Rating definitions", st["h3"]), table]))
    nar = res.narrative
    prep = (f"Prepared on {date.fromisoformat(v.valuation_date):%d %B %Y} by equity-research-agent, an open-source model, "
            f"from public data (Screener.in, yfinance). The price used is {f.price:,.2f} (Screener.in, fetched "
            f"{f.fetched_at}).")
    if nar is not None:
        who = nar.get("prepared_by") or "the host AI agent"
        prep += (f" Qualitative sections, flag responses and analyst inputs were written by {who}, an AI agent, from the "
                 "numbered sources listed above. No human analyst reviewed this report.")
    else:
        prep += " This is a draft: the qualitative sections have not been written and no analyst has reviewed it."
    items = [
        ("Status", "Model-generated research for education. It is not investment advice and is not a research report "
                   "under the SEBI (Research Analysts) Regulations, 2014."),
        ("Preparation", prep),
        ("Interests", "The software and its authors hold no position in the securities discussed, provide no services to "
                      "and receive no compensation from the companies covered."),
        ("Limits", "Forecasts are model estimates driven by historical averages unless an analyst input with a source "
                   "replaces them, and will differ from outcomes. The target and rating are mechanical results of the "
                   "methods in the methodology section; flags in the valuation section show where the methods disagree. "
                   "Verify figures against company filings before acting."),
    ]
    for head, txt in items:
        out.append(Paragraph(f"<b>{head}.</b> {txt}", st["small"]))
        out.append(Spacer(1, 2))
    return out
