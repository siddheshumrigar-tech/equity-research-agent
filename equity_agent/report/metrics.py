"""Per-year metrics across reported and forecast years, on one set of definitions.

Valuation multiples are at the current price and current enterprise value
(EV = market cap + borrowings - cash - investments), as broker key-financials tables show them.
"""
from __future__ import annotations

from typing import Dict, List, Optional

from ..ratios import annual


def key_metrics(res) -> Dict[str, List[Optional[float]]]:
    f, fc, v = res.fin, res.forecast, res.valuation
    L, F = f.lines, fc.lines
    shares = f.shares_cr
    n = f.n
    cat = lambda k, fk=None: list(L[k]) + list(F[fk or k])
    sales, ebitda, npf = cat("sales"), cat("ebitda"), cat("net_profit")
    equity, debt = cat("total_equity"), cat("borrowings")
    payout = list(L["dividend_payout"]) + [res.drivers.value("dividend_payout", fc.scenario, t) for t in range(len(fc.years))]
    eps = [(f.eps[i] if i < n and f.eps[i] is not None else npf[i] / shares) for i in range(len(npf))]
    out: Dict[str, List[Optional[float]]] = {"labels": f.year_labels + fc.year_labels}
    out["sales"], out["ebitda"], out["net_profit"], out["eps"] = sales, ebitda, npf, eps
    out["sales_growth"] = [None] + [b / a - 1 if a else None for a, b in zip(sales, sales[1:])]
    out["eps_growth"] = [None] + [b / a - 1 if a and a > 0 else None for a, b in zip(eps, eps[1:])]
    out["ebitda_margin"] = [e / s if s else None for e, s in zip(ebitda, sales)]
    out["net_margin"] = [p / s if s else None for p, s in zip(npf, sales)]
    out["pe"] = [f.price / e if e and e > 0 else None for e in eps]
    out["ev_ebitda"] = [v.current_ev / e if e and e > 0 else None for e in ebitda]
    out["pb"] = [f.price / (q / shares) if q and q > 0 else None for q in equity]
    rt = annual(res)
    out["roe"], out["roce"] = rt["roe"], rt["roce"]
    out["dps"] = [max(p, 0) * po / shares for p, po in zip(npf, payout)]
    out["div_yield"] = [d / f.price for d in out["dps"]]
    cash, inv = cat("cash"), cat("investments")
    out["net_debt_equity"] = [(d - c_ - i) / q if q else None for d, c_, i, q in zip(debt, cash, inv, equity)]
    out["fcf"] = list(L["fcf"]) + list(F["fcff"])
    out["n_actual"] = n
    return out


def window(m: dict, n_act_before: int = 3, n_fc: int = 3):
    """Indices for 'last n actual + first n forecast' columns."""
    n = m["n_actual"]
    return list(range(max(0, n - n_act_before), n)) + list(range(n, min(len(m["labels"]), n + n_fc)))
