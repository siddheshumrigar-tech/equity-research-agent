"""Financial analysis sections built from reported data and the model (no narrative needed)."""
from __future__ import annotations

from reportlab.platypus import Paragraph, Spacer

from ...ratios import annual, on_average
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
    npf = f.lines["net_profit"] + fc.lines["net_profit"]
    p2 = r.charts.lines("perf_margins", labels, [("Gross margin", annual(res)["gross_margin"]),
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
    if f.cost_basis == "services":
        note = ("Service cost structure: COGS (cost of revenue) = material, manufacturing and employee cost; "
                "operating expenses = all other items.")
    else:
        note = "COGS = material and manufacturing items; operating expenses = all other items."
    out.append(pair(r.ex.make("Cost structure", [t, Spacer(1, 2), r.P(note, "small")], "Screener.in expense schedule"),
                    r.ex.make("Expense mix (% of revenue)", r.img(p, HALF_W, 170), "Screener.in")))
    rt = {k: v[:f.n] for k, v in annual(res).items()}
    dso, dio, dpo, ccc = rt["dso"], rt["dio"], rt["dpo"], rt["ccc"]
    p2 = r.charts.lines("wc_days", labels_a, [("Receivable days", dso), ("Inventory days", dio),
                                              ("Payable days", dpo), ("Cash cycle", ccc)], HALF_W, CH_H, pct=False,
                        legend_cols=2)
    cfo, eb = L["cfo"], L["ebitda"]
    conv = [a / b if b else None for a, b in zip(cfo, eb)]
    capex_int = [a / b if b else None for a, b in zip(L["capex"], L["sales"])]
    p3 = r.charts.lines("cash_conv", labels_a, [("CFO / EBITDA", conv), ("Capex / revenue", capex_int)], HALF_W, CH_H,
                        legend_cols=2)
    out.append(pair(r.ex.make("Working-capital days (reported balances)", r.img(p2, HALF_W, CH_H),
                              "Screener.in; DIO on COGS; DPO on " +
                              ("COGS" if res.drivers.dpo_basis == "cogs" else "total expenses")),
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
    # Average balances throughout, so the ROE here is the ROE of the key-ratios table and the Excel model.
    rt = annual(res)
    turnover = on_average(L["sales"], L["total_assets"])
    avg_ta = [None] + [(a + b) / 2 for a, b in zip(L["total_assets"], L["total_assets"][1:])]
    avg_eq = [None] + [(a + b) / 2 for a, b in zip(L["total_equity"], L["total_equity"][1:])]
    npm = [L["net_profit"][i] / L["sales"][i] if L["sales"][i] else None for i in idx]
    at = [turnover[i] for i in idx]
    em = [avg_ta[i] / avg_eq[i] if avg_ta[i] is not None and avg_eq[i] and avg_eq[i] > 0 else None for i in idx]
    roe = [rt["roe"][i] for i in idx]
    p = r.charts.small_multiples("dupont", lab, [("Net margin", npm, "pct"), ("Asset turnover (x)", at, "x"),
                                                 ("Equity multiplier (x)", em, "x"), ("ROE", roe, "pct")], CONTENT_W, 120)
    rows = [[r.P("Ratio", "thl")] + [r.P(h, "th") for h in lab]]
    spec = [("Net margin", npm, fmt_pct), ("Asset turnover (x)", at, lambda x: f"{x:.2f}"),
            ("Equity multiplier (x)", em, lambda x: f"{x:.2f}"), ("ROE", roe, fmt_pct),
            ("ROCE (pre-tax)", [rt["roce"][i] for i in idx], fmt_pct),
            ("Debt / equity (x)", [L["borrowings"][i] / L["total_equity"][i] if L["total_equity"][i] else None for i in idx],
             lambda x: f"{x:.2f}")]
    for name, vals, fmt in spec:
        rows.append([r.P(name, "tdl")] + [r.P(fmt(v) if v is not None else "n/a", "td") for v in vals])
    lw = CONTENT_W * 0.28
    t = data_table(r.theme, rows, [lw] + [(CONTENT_W - lw) / len(lab)] * len(lab))
    out = section_title(r.theme, "Return ratios (DuPont)", num)
    out.append(r.P("ROE = net margin x asset turnover x equity multiplier, on average balances (prior and current "
                   "year-end). ROCE = (PBT + interest) / average (equity + borrowings). Debt / equity at year-end.",
                   "small"))
    out += [r.ex.make("DuPont decomposition", r.img(p, CONTENT_W, 120), "Screener.in"),
            r.ex.make("Return and leverage ratios", t, "Screener.in")]
    return out
