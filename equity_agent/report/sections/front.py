"""Page 1: headline, thesis, key financials (main column) and the data sidebar."""
from __future__ import annotations

from typing import List

from reportlab.platypus import KeepInFrame, Paragraph, Spacer, Table, TableStyle

from ..common import R
from ..metrics import key_metrics, window
from ..styles import CONTENT_W, INK, MUTED, RATING_COLOURS, RS, SHADE, c, data_table, fmt_num, fmt_pct, fmt_x, tint

MAIN_SHARE = 0.655
GAP = 12


def auto_headline(r: R) -> str:
    v = r.res.valuation
    return (f"{v.rating}: 12-month target {RS}{v.target_price:,.0f} implies {v.expected_total_return:+.0%} "
            "total return")


def auto_points(r: R) -> List[str]:
    """Number-only summary used before the host agent writes the narrative (labelled as automatic)."""
    res = r.res
    f, v, d = res.fin, res.valuation, res.drivers
    L = f.lines
    n = f.n
    pts = []
    if n > 3 and L["sales"][-4] > 0:
        c3 = (L["sales"][-1] / L["sales"][-4]) ** (1 / 3) - 1
        pts.append(f"<b>Growth.</b> Revenue {RS}{L['sales'][-1]:,.0f} Cr in FY{f.years[-1] % 100:02d}, {c3:+.1%} a "
                   f"year over three years; the base case carries {d.value('sales_growth'):.1%} growth, the 3-year "
                   "average, into the forecast.")
    m_now = L["ebitda"][-1] / L["sales"][-1]
    pts.append(f"<b>Profitability.</b> EBITDA margin {m_now:.1%}; current ROIC {v.roic:.0%} against a WACC of "
               f"{v.wacc:.1%}, fading to a terminal return on new capital of {v.ronic:.0%}.")
    methods = []
    if v.method_values.get("dcf") is not None:
        methods.append(f"DCF {RS}{v.dcf_value_12m:,.0f}")
    if v.pe_value is not None:
        methods.append(f"P/E {RS}{v.pe_value:,.0f} ({v.pe_mean:.0f}x mean)")
    if v.ev_ebitda_value is not None:
        methods.append(f"EV/EBITDA {RS}{v.ev_ebitda_value:,.0f} ({v.peer_ev_ebitda:.0f}x peers)")
    pts.append(f"<b>Valuation.</b> " + ", ".join(methods) + f"; target {RS}{v.target_price:,.0f} vs CMP "
               f"{RS}{f.price:,.0f}.")
    if v.reverse_dcf_growth is not None:
        pts.append(f"<b>What the price implies.</b> Revenue growth of {v.reverse_dcf_growth:.1%} a year for five years "
                   f"(base case {d.value('sales_growth'):.1%}), with margins and returns unchanged.")
    else:
        pts.append("<b>What the price implies.</b> No five-year growth rate in the model reaches the current price; "
                   "the market prices returns or a horizon beyond the model.")
    hi = [fl for fl in res.flags if fl.severity == "high"]
    if hi:
        names = {"TV_SHARE_HIGH": "terminal value dominates the DCF",
                 "DCF_RELATIVE_DIVERGENCE": "DCF and multiples disagree by more than 2x",
                 "EXTREME_UPSIDE": "target is more than 40% from the price",
                 "NEGATIVE_FCFF": "free cash flow turns negative"}
        pts.append("<b>Model flags.</b> " + "; ".join(names.get(fl.code, fl.code) for fl in hi)
                   + ". These need the analyst's response before publication.")
    return pts


def _key_risks(r: R) -> list:
    """Up to three narrative risks with the analyst's estimated effect on the target (page 1)."""
    sized = lambda x: isinstance(x.get("impact_on_tp_pct"), (int, float))
    risks = [x for x in (r.nar.get("risks") or []) if isinstance(x, dict) and str(x.get("risk", "")).strip()
             and not (sized(x) and x["impact_on_tp_pct"] > 0)]  # upside cases are not risks to the call
    if not risks:
        return []
    tp = r.res.valuation.target_price
    risks = sorted(risks, key=lambda x: -abs(x["impact_on_tp_pct"]) if sized(x) else 0)[:3]
    out = [Spacer(1, 2), Paragraph("Key risks to the call", r.st["h3"])]
    for x in risks:
        size = (f" ({x['impact_on_tp_pct']:+.0%} on the target, {RS}{tp * (1 + x['impact_on_tp_pct']):,.0f})"
                if sized(x) else "")
        out.append(Paragraph(f"<b>{r.N(x['risk'])}</b>{size}: {r.N(x['impact'])}", r.st["bul"], bulletText="■"))
    return out


def _key_financials(r: R, width: float):
    m = key_metrics(r.res)
    idx = window(m, 3, 3)
    n = m["n_actual"]
    heads = [m["labels"][i] for i in idx]
    rows = [[r.P(f"YE {r.res.fin.period_labels[-1][:3]} ({RS} Cr)", "thl")] + [r.P(h, "th") for h in heads]]
    spec = [("Revenue", "sales", fmt_num), ("  growth", "sales_growth", lambda x: fmt_pct(x, 1, True)),
            ("EBITDA", "ebitda", fmt_num), ("  margin", "ebitda_margin", fmt_pct), ("Net profit", "net_profit", fmt_num),
            (f"EPS ({RS})", "eps", lambda x: fmt_num(x, 2)), ("  growth", "eps_growth", lambda x: fmt_pct(x, 1, True)),
            ("P/E (x)", "pe", fmt_x), ("EV/EBITDA (x)", "ev_ebitda", fmt_x), ("P/B (x)", "pb", fmt_x),
            ("ROE", "roe", fmt_pct), ("Dividend yield", "div_yield", fmt_pct)]
    for lab, key, fmt in spec:
        rows.append([r.P(lab, "tdl")] + [r.P(fmt(m[key][i]) if m[key][i] is not None else "n/a", "td") for i in idx])
    lw = width * 0.22
    cw = (width - lw) / len(heads)
    est = [k + 1 for k, i in enumerate(idx) if i >= n]
    t = data_table(r.theme, rows, [lw] + [cw] * len(heads), estimate_cols=est, bold_rows=(1, 5))
    t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 1.5), ("RIGHTPADDING", (0, 0), (-1, -1), 1.5)]))
    return t


def _valuation_strip(r: R, width: float):
    """Compact method -> target table for page 1."""
    from ...valuation import METHOD_LABELS

    v = r.res.valuation
    rows = [[r.P("Method", "thl"), r.P(f"Value ({RS})", "th"), r.P("Weight", "th")]]
    if v.target_method == "sotp":
        rows.append([r.P(METHOD_LABELS["sotp"], "tdl"), r.P(f"{v.sotp_value:,.0f}", "td"), r.P("100%", "td")])
    else:
        for m in ("dcf", "pe", "ev_ebitda"):
            val = v.method_values.get(m)
            rows.append([r.P(METHOD_LABELS[m], "tdl"), r.P(f"{val:,.0f}" if val is not None else "n/a", "td"),
                         r.P(fmt_pct(v.weights.get(m, 0), 0) if val is not None else "-", "td")])
    rows.append([r.P("<b>12-month target</b>", "tdl"), r.P(f"<b>{v.target_price:,.0f}</b>", "tdb"), r.P("", "td")])
    return data_table(r.theme, rows, [width * 0.56, width * 0.19, width * 0.25], total_rows=(len(rows) - 1,))


def _kv_table(r: R, rows, width: float, bold_last: bool = False):
    data = [[r.P(a, "side"), r.P(b, "side_r")] for a, b in rows]
    t = Table(data, colWidths=[width * 0.56, width * 0.44])
    st = [("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
          ("TOPPADDING", (0, 0), (-1, -1), 1.1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.1),
          ("LINEBELOW", (0, 0), (-1, -1), 0.25, c("E5E7EB"))]
    t.setStyle(TableStyle(st))
    return t


def _rating_box(r: R, width: float):
    v, f = r.res.valuation, r.res.fin
    col = RATING_COLOURS.get(v.rating, INK)
    st = r.st
    inner = [
        [Paragraph(f"<font color='#{col}'>{v.rating}</font>", st["rating"]), ""],
        [r.P("CMP", "side"), r.P(f"{RS}{f.price:,.2f}", "side_rb")],
        [r.P("12-month target", "side"), r.P(f"<b>{RS}{v.target_price:,.0f}</b>", "side_rb")],
        [r.P("Upside / (downside)", "side"), r.P(fmt_pct(v.upside, 1, True), "side_r")],
        [r.P("Dividend yield", "side"), r.P(fmt_pct(v.dividend_yield, 1), "side_r")],
        [r.P("<b>Expected total return</b>", "side"), r.P(f"<b>{fmt_pct(v.expected_total_return, 1, True)}</b>", "side_r")],
    ]
    t = Table(inner, colWidths=[width * 0.58 - 8, width * 0.42 - 8])
    t.setStyle(TableStyle([("SPAN", (0, 0), (1, 0)), ("BACKGROUND", (0, 0), (-1, -1), c(r.theme.light)),
                           ("LINEABOVE", (0, 0), (-1, 0), 2.2, c(col)), ("LEFTPADDING", (0, 0), (-1, -1), 8),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 8), ("TOPPADDING", (0, 0), (-1, -1), 1.3),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 1.3), ("BOTTOMPADDING", (0, -1), (-1, -1), 6),
                           ("TOPPADDING", (0, 0), (-1, 0), 5)]))
    return t


def _sidebar(r: R, width: float):
    res = r.res
    f, v, t = res.fin, res.valuation, res.technicals
    st = r.st
    out = [_rating_box(r, width), Spacer(1, 2)]
    hi = t.high_52w if t else f.high_52
    lo = t.low_52w if t else f.low_52
    stock = [("NSE code", f.symbol)] + ([("BSE code", f.bse_code)] if f.bse_code else []) + [
             (f"Market cap ({RS} Cr)", f"{f.market_cap:,.0f}"), ("Shares (Cr)", f"{f.shares_cr:,.2f}"),
             ("52-week high / low", f"{fmt_num(hi)} / {fmt_num(lo)}"),
             ("Beta (Blume adj.)", f"{v.beta:.2f}"), ("Trailing P/E", fmt_x(f.pe)), ("WACC / g", f"{v.wacc:.1%} / {v.terminal_growth:.1%}")]
    out += [Paragraph("Stock data", st["side_h"]), _kv_table(r, stock, width)]
    if f.shareholding and f.shareholding_labels:
        q = f.shareholding_labels
        k0 = max(0, len(q) - 5)
        rows = [[r.P("%", "label"), r.P(q[-1], "label"), r.P(q[k0], "label")]]
        for lab in ("Promoters", "FIIs", "DIIs", "Government", "Public"):
            vals = f.shareholding.get(lab)
            if vals and vals[-1] is not None:
                rows.append([r.P(lab, "side"), r.P(f"{vals[-1]:.1f}", "side_r"),
                             r.P(f"{vals[k0]:.1f}" if vals[k0] is not None else "n/a", "side_r")])
        tb = Table(rows, colWidths=[width * 0.44, width * 0.28, width * 0.28])
        tb.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                                ("LINEBELOW", (0, 0), (-1, 0), 0.5, c(r.theme.accent)),
                                ("LINEBELOW", (0, 1), (-1, -1), 0.25, c("E5E7EB")), ("ALIGN", (1, 0), (-1, 0), "RIGHT")]))
        out += [Paragraph("Shareholding (%)", st["side_h"]), tb]
    if t is not None:
        rows = [[r.P("%", "label"), r.P("1M", "label"), r.P("3M", "label"), r.P("12M", "label")],
                [r.P("Absolute", "side")] + [r.P(fmt_pct(t.returns.get(k), 1, True), "side_r") for k in ("1M", "3M", "12M")],
                [r.P("vs Nifty 50", "side")] + [r.P(fmt_pct(t.relative.get(k), 1, True), "side_r") for k in ("1M", "3M", "12M")]]
        tb = Table(rows, colWidths=[width * 0.34] + [width * 0.22] * 3)
        tb.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                ("TOPPADDING", (0, 0), (-1, -1), 1), ("BOTTOMPADDING", (0, 0), (-1, -1), 1),
                                ("LINEBELOW", (0, 0), (-1, 0), 0.5, c(r.theme.accent)), ("ALIGN", (1, 0), (-1, 0), "RIGHT"),
                                ("LINEBELOW", (0, 1), (-1, -1), 0.25, c("E5E7EB"))]))
        out += [Paragraph("Price performance", st["side_h"]), tb]
        d, cl, ix = res.market.last_year()
        p = r.charts.price_vs_index("front_price", d, cl, ix, width, 92, label=f.symbol)
        if p:
            out += [Spacer(1, 3), r.img(p, width, 92), r.P("12-month price, rebased to 100. Source: yfinance.", "source")]
    sc = v.scenarios
    if sc:
        p = r.charts.risk_reward("front_rr", sc["bear"].target, sc["base"].target, sc["bull"].target, f.price, width, 52)
        out += [Paragraph("Risk-reward (12-month value)", st["side_h"]), r.img(p, width, 52)]
    return out


def front_page(r: R) -> list:
    res = r.res
    f = res.fin
    st = r.st
    wm = (CONTENT_W - GAP) * MAIN_SHARE
    ws = CONTENT_W - GAP - wm
    nar = r.nar
    main = [Paragraph(f"{f.name}", st["h2"]),
            Paragraph(r.N(nar.get("headline")) if nar else auto_headline(r), st["headline"])]
    if nar and nar.get("thesis_points"):
        for tp in nar.get("thesis_points")[:5]:
            main.append(Paragraph(f"<b>{r.N(tp['title'])}.</b> {r.N(tp.get('summary') or tp['text'])}", st["bul"],
                                  bulletText="■"))
        main += _key_risks(r)
    else:
        main.append(Paragraph("<i>Automatic summary from the model; the analyst narrative has not been written yet.</i>",
                              st["small"]))
        for pt in auto_points(r):
            main.append(Paragraph(pt, st["bul"], bulletText="■"))
    main += [Spacer(1, 4), Paragraph("Key financials and valuation", st["h3"]), _key_financials(r, wm),
             Paragraph("Source: Screener.in (A = reported), model estimates (E). Multiples at the current price and "
                       "enterprise value. EPS for reported years as reported.", st["source"])]
    fc = res.forecast
    labels = f.year_labels[-6:] + fc.year_labels
    sales = f.lines["sales"][-6:] + fc.lines["sales"]
    ebitda = f.lines["ebitda"][-6:] + fc.lines["ebitda"]
    na = min(6, f.n)
    ch_w = (wm - 10) * 0.52
    p = r.charts.bars_line("front_rev", labels, sales, [e / s if s else 0 for e, s in zip(ebitda, sales)], na, ch_w, 118,
                           f"Revenue ({RS} Cr)", "EBITDA margin")
    left_col = [Paragraph("Revenue and EBITDA margin", st["h3"]), r.img(p, ch_w, 118),
                Paragraph("Dark = reported, light = forecast. Source: Screener.in, model.", st["source"])]
    right_col = [Paragraph("Valuation: methods to target", st["h3"]), _valuation_strip(r, wm - 10 - ch_w),
                 Paragraph(f"DCF rolled forward 12 months; multiples on next-twelve-months earnings. "
                           f"Bull / bear: {RS}{res.valuation.scenarios['bull'].target:,.0f} / "
                           f"{RS}{res.valuation.scenarios['bear'].target:,.0f}." if res.valuation.scenarios else "",
                           st["source"])]
    from ..common import pair
    main += [Spacer(1, 6), pair(left_col, right_col, w_left=0.52, gap=10, width=wm)]
    height = r.front_h - 4
    left = KeepInFrame(wm, height, main, mode="shrink")
    right = KeepInFrame(ws, height, _sidebar(r, ws), mode="shrink")
    t = Table([[left, "", right]], colWidths=[wm, GAP, ws])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    return [t]
