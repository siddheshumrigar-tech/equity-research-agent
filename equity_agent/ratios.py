"""One definition per ratio for the subject company, used by the PDF report, the Excel model and the agent files.

Series cover every reported year followed by every forecast year.

* Gross margin = (revenue - cost of revenue) / revenue. Cost of revenue follows ``fin.cost_basis``.
* Returns use average balances (prior and current year-end), so the first reported year has none:
  ROE = net profit / average equity (equity capital + reserves); ROA = net profit / average total assets;
  ROCE (pre-tax) = (PBT + interest) / average (equity + borrowings). The peer tables use the same ROE and ROCE.
* Working-capital days use year-end balances: DSO = receivables / revenue x 365; DIO = inventory / cost of
  revenue x 365; DPO = trade payables / payables base x 365, where the base is cost of revenue, or total expenses
  when ``drivers.dpo_basis`` is "expenses"; cash conversion cycle = DSO + DIO - DPO.

The Excel model writes these as live formulas. ``tests/test_parity.py`` checks that the recalculated workbook
equals these series in every year.
"""
from __future__ import annotations

from typing import Dict, List, Optional

Series = List[Optional[float]]


def _div(a: Optional[float], b: Optional[float]) -> Optional[float]:
    return a / b if a is not None and b else None


def _pos_div(a: Optional[float], b: Optional[float]) -> Optional[float]:
    return a / b if a is not None and b is not None and b > 0 else None


def on_average(num: Series, bal: Series) -> Series:
    """num[i] / average(bal[i - 1], bal[i]); None in the first year or when the average is not positive."""
    out: Series = [None]
    for i in range(1, len(num)):
        a, b = bal[i - 1], bal[i]
        out.append(_pos_div(num[i], (a + b) / 2) if a is not None and b is not None else None)
    return out


def days(bal: Series, base: Series) -> Series:
    """bal / base x 365; None when the base is not positive."""
    return [None if (r := _pos_div(x, b)) is None else r * 365 for x, b in zip(bal, base)]


def annual(res) -> Dict[str, Series]:
    """Ratio series for reported + forecast years of a pipeline result (``fin``, ``forecast``, ``drivers``)."""
    f, fc = res.fin, res.forecast
    L, F = f.lines, fc.lines
    cat = lambda k: list(L[k]) + list(F[k])
    sales, cogs, gp = cat("sales"), cat("cogs"), cat("gross_profit")
    npf, pbt, interest = cat("net_profit"), cat("pbt"), cat("interest")
    equity, assets, debt = cat("total_equity"), cat("total_assets"), cat("borrowings")
    expenses = list(L["expenses"]) + [c + o for c, o in zip(F["cogs"], F["opex"])]
    pay_base = cogs if res.drivers.dpo_basis == "cogs" else expenses
    out: Dict[str, Series] = {"labels": f.year_labels + fc.year_labels}
    out["gross_margin"] = [_div(g, s) for g, s in zip(gp, sales)]
    out["roe"] = on_average(npf, equity)
    out["roa"] = on_average(npf, assets)
    out["roce"] = on_average([p + (i or 0) for p, i in zip(pbt, interest)], [q + (d or 0) for q, d in zip(equity, debt)])
    out["dso"] = days(cat("receivables"), sales)
    out["dio"] = days(cat("inventory"), cogs)
    out["dpo"] = days(cat("trade_payables"), pay_base)
    out["ccc"] = [None if a is None or p is None else a + (i or 0.0) - p
                  for a, i, p in zip(out["dso"], out["dio"], out["dpo"])]
    return out


def latest_actual(res, key: str) -> Optional[float]:
    """The ratio for the last reported year."""
    return annual(res)[key][res.fin.n - 1]
