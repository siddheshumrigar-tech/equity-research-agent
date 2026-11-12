"""Financial analysis sections built from reported data and the model (no narrative needed)."""
from __future__ import annotations

from reportlab.platypus import Paragraph, Spacer

from ..common import R, pair
from ..layout import section_title
from ..metrics import key_metrics
from ..styles import CONTENT_W, HALF_W, RS, data_table, fmt_num, fmt_pct, fmt_x

CH_H = 150


def performance(r: R, num: int) -> list:
    """Charts only; the full income statement is in the Financial statements section."""
    res = r.res
    f = res.fin
    out = section_title(r.theme, "Financial performance", num)
    fc = res.forecast
    labels = f.year_labels + fc.year_labels
    s_all = f.lines["sales"] + fc.lines["sales"]
    e_all = f.lines["ebitda"] + fc.lines["ebitda"]
    p1 = r.charts.bars_line("perf_rev", labels, s_all, [e / x if x else 0 for e, x in zip(e_all, s_all)], f.n,
                            HALF_W, CH_H, f"Revenue ({RS} Cr)", "EBITDA margin")
    gp = f.lines["gross_profit"] + fc.lines["gross_profit"]
    npf = f.lines["net_profit"] + fc.lines["net_profit"]
    p2 = r.charts.lines("perf_margins", labels, [("Gross margin", [g / x if x else None for g, x in zip(gp, s_all)]),
                                                 ("EBITDA margin", [e / x if x else None for e, x in zip(e_all, s_all)]),
                                                 ("Net margin", [p / x if x else None for p, x in zip(npf, s_all)])],
                        HALF_W, CH_H, split=f.n)
    out.append(pair(r.ex.make("Revenue and EBITDA margin (dark = reported, light = forecast)",
                              r.img(p1, HALF_W, CH_H)),
                    r.ex.make("Margin trends, reported | forecast", r.img(p2, HALF_W, CH_H))))
    return out


def costs_and_working_capital(r: R, num: int) -> list:
    res = r.res
    f = res.fin
    L = f.lines
    k3 = max(0, f.n - 4)
    rows = [[r.P("Expense item", "thl"), r.P("Type", "th"), r.P(f"FY{f.years[-1] % 100:02d} {RS} Cr", "th"),
             r.P(f"FY{f.years[-1] % 100:02d} % rev.", "th"), r.P(f"FY{f.years[k3] % 100:02d} % rev.", "th"),
             r.P("Chg. (pp)", "th")]]
    for it in f.expense_items:
        now = it.values[-1] / L["sales"][-1] if L["sales"][-1] else None
        then = it.values[k3] / L["sales"][k3] if L["sales"][k3] else None
        rows.append([r.P(it.name, "tdl"), r.P(it.kind, "td"), r.P(fmt_num(it.values[-1]), "td"), r.P(fmt_pct(now), "td"),
                     r.P(fmt_pct(then), "td"),
                     r.P(f"{(now - then) * 100:+.1f}" if now is not None and then is not None else "n/a", "td")])
    w = HALF_W
    t = data_table(r.theme, rows, [w * 0.36, w * 0.11, w * 0.15, w * 0.13, w * 0.13, w * 0.12])
    labels_a = f.year_labels
    series = [(it.name, [v / x if x else 0 for v, x in zip(it.values, L["sales"])]) for it in f.expense_items]
    p = r.charts.stacked("cost_mix", labels_a, series, HALF_W, 170)
    out = section_title(r.theme, "Cost structure and working capital", num)
    out.append(pair(r.ex.make("Cost structure", [t, Spacer(1, 2), r.P("COGS = material and manufacturing items; "
                                                                    "operating expenses = all other items.", "small")],
                              "Screener.in expense schedule"),
                    r.ex.make("Expense mix (% of revenue)", r.img(p, HALF_W, 170), "Screener.in")))
    dso = [x / s * 365 if s else None for x, s in zip(L["receivables"], L["sales"])]
    dio = [x / c_ * 365 if c_ > 0 else None for x, c_ in zip(L["inventory"], L["cogs"])]
    base = L["cogs"] if res.drivers.dpo_basis == "cogs" else L["expenses"]
    dpo = [x / b * 365 if b > 0 else None for x, b in zip(L["trade_payables"], base)]
    ccc = [None if a is None or c_ is None else a + (b or 0) - c_ for a, b, c_ in zip(dso, dio, dpo)]
    p2 = r.charts.lines("wc_days", labels_a, [("Receivable days", dso), ("Inventory days", dio),
                                              ("Payable days", dpo), ("Cash cycle", ccc)], HALF_W, CH_H, pct=False,
                        legend_cols=2)
    cfo, eb = L["cfo"], L["ebitda"]
    conv = [a / b if b else None for a, b in zip(cfo, eb)]
    capex_int = [a / b if b else None for a, b in zip(L["capex"], L["sales"])]
    p3 = r.charts.lines("cash_conv", labels_a, [("CFO / EBITDA", conv), ("Capex / revenue", capex_int)], HALF_W, CH_H,
                        legend_cols=2)
    out.append(pair(r.ex.make("Working-capital days (reported balances)", r.img(p2, HALF_W, CH_H),
                              "Screener.in; DPO on " + ("COGS" if res.drivers.dpo_basis == "cogs" else "total expenses")),
                    r.ex.make("Cash conversion and capex intensity", r.img(p3, HALF_W, CH_H), "Screener.in")))
    return out


def balance_sheet_cash_flow(r: R, num: int) -> list:
    """Free cash flow chart; the full balance sheet and cash flow are in the Financial statements section."""
    res = r.res
    f = res.fin
    out = section_title(r.theme, "Free cash flow", num)
    m = key_metrics(res)
    labels = m["labels"]
    p = r.charts.bars("fcf", labels, m["fcf"], CONTENT_W, 120, n_actual=f.n)
    out.append(r.ex.make(f"Free cash flow ({RS} Cr): reported CFO - capex | forecast FCFF", r.img(p, CONTENT_W, 120)))
    return out


def returns(r: R, num: int) -> list:
    res = r.res
    f = res.fin
    L = f.lines
    idx = list(range(max(0, f.n - 6), f.n))
    lab = [f.year_labels[i] for i in idx]
    npm = [L["net_profit"][i] / L["sales"][i] if L["sales"][i] else None for i in idx]
    at = [L["sales"][i] / L["total_assets"][i] if L["total_assets"][i] else None for i in idx]
    em = [L["total_assets"][i] / L["total_equity"][i] if L["total_equity"][i] else None for i in idx]
    roe = [L["net_profit"][i] / L["total_equity"][i] if L["total_equity"][i] else None for i in idx]
    p = r.charts.small_multiples("dupont", lab, [("Net margin", npm, "pct"), ("Asset turnover (x)", at, "x"),
                                                 ("Equity multiplier (x)", em, "x"), ("ROE", roe, "pct")], CONTENT_W, 120)
    rows = [[r.P("Ratio", "thl")] + [r.P(h, "th") for h in lab]]
    roce = f.reported_ratios.get("ROCE %")
    spec = [("Net margin", npm, fmt_pct), ("Asset turnover (x)", at, lambda x: f"{x:.2f}"),
            ("Equity multiplier (x)", em, lambda x: f"{x:.2f}"), ("ROE (closing equity)", roe, fmt_pct),
            ("ROCE % (Screener)", [roce[i] / 100 if roce and roce[i] is not None else None for i in idx], fmt_pct),
            ("Debt / equity (x)", [L["borrowings"][i] / L["total_equity"][i] if L["total_equity"][i] else None for i in idx],
             lambda x: f"{x:.2f}")]
    for name, vals, fmt in spec:
        rows.append([r.P(name, "tdl")] + [r.P(fmt(v) if v is not None else "n/a", "td") for v in vals])
    lw = CONTENT_W * 0.28
    t = data_table(r.theme, rows, [lw] + [(CONTENT_W - lw) / len(lab)] * len(lab))
    out = section_title(r.theme, "Return ratios (DuPont)", num)
    out.append(r.P("ROE = net margin x asset turnover x equity multiplier, from reported closing balances.", "small"))
    out += [r.ex.make("DuPont decomposition", r.img(p, CONTENT_W, 120), "Screener.in"),
            r.ex.make("Return and leverage ratios", t, "Screener.in")]
    return out
