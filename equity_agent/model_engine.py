"""Articulated 3-statement forecast (Python mirror of the Excel formulas).

The Excel workbook writes the same logic as formulas; tests check that both agree.
Balance sheet balances by construction: cash is the output of the cash-flow
statement and there is no plug line.

Treatment notes:
* Loans & advances are non-operating and held flat (they are not working capital).
* Capital work in progress is capitalised into the block evenly over ``cwip_years``.
* Minority interest & associates take the historical share of profit after tax.
"""
from __future__ import annotations

import sys
from dataclasses import dataclass
from typing import Dict, List, Optional

from .drivers import Drivers, build_drivers
from .financials import CompanyFinancials

FORECAST_YEARS = 5
CWIP_YEARS = 2


@dataclass
class Forecast:
    scenario: str
    years: List[int]
    lines: Dict[str, List[float]]

    @property
    def year_labels(self) -> List[str]:
        return [f"FY{y % 100:02d}E" for y in self.years]


def base_year_balances(fin: CompanyFinancials) -> Dict[str, float]:
    """Last reported balances, in the grouping the forecast uses."""
    L = fin.lines
    b = {k: L[k][-1] for k in (
        "sales", "cogs", "fixed_assets", "cwip", "investments", "receivables", "inventory", "cash",
        "equity_capital", "reserves", "borrowings", "lt_borrowings", "st_borrowings", "lease_liabilities",
        "other_borrowings", "trade_payables", "minority_interest", "loans_advances")}
    b["other_ca"] = L["other_asset_items"][-1]
    b["other_liab"] = L["other_liability_items"][-1] + L["advance_from_customers"][-1]
    b["gross_block"] = fin.gross_block[-1] if fin.gross_block else None
    b["acc_dep"] = fin.acc_dep[-1] if fin.acc_dep else None
    b["nwc"] = operating_nwc(b)
    return b


def operating_nwc(b: Dict[str, float]) -> float:
    return b["receivables"] + b["inventory"] + b["other_ca"] - b["trade_payables"] - b["other_liab"]


def invested_capital(b: Dict[str, float]) -> float:
    """Operating invested capital = net block + CWIP + operating working capital."""
    return b["fixed_assets"] + b["cwip"] + b["nwc"]


def run_forecast(fin: CompanyFinancials, drivers: Drivers, scenario: str = "base", n_years: int = FORECAST_YEARS,
                 cwip_years: Optional[int] = None) -> Forecast:
    cwip_years = max(1, int(cwip_years or CWIP_YEARS))
    dv = lambda k, t: drivers.value(k, scenario, t)
    b = base_year_balances(fin)
    cwip0 = b["cwip"]
    years = [fin.years[-1] + i for i in range(1, n_years + 1)]
    keys = ["sales", "cogs", "gross_profit", "opex", "ebitda", "other_income", "depreciation", "ebit", "interest",
            "pbt", "tax", "minority_other", "net_profit", "dividends", "capex", "cwip_transfer", "gross_block",
            "acc_dep", "fixed_assets", "cwip", "investments", "receivables", "inventory", "trade_payables", "other_ca",
            "loans_advances", "other_liab", "minority_interest", "borrowings", "lt_borrowings", "st_borrowings",
            "lease_liabilities", "other_borrowings", "equity_capital", "reserves", "total_equity", "cash",
            "total_assets", "total_liabilities", "balance_check", "nwc", "delta_nwc", "cfo", "cfi", "cff",
            "net_cash_flow", "nopat", "fcff", "ebit_operating", "invested_capital", "dso", "dio", "dpo"]
    out: Dict[str, List[float]] = {k: [] for k in keys}
    prev = dict(b)
    for t in range(n_years):
        cur: Dict[str, float] = {}
        cur["sales"] = prev["sales"] * (1 + dv("sales_growth", t))
        cur["cogs"] = cur["sales"] * dv("cogs_pct", t)
        cur["gross_profit"] = cur["sales"] - cur["cogs"]
        cur["ebitda"] = cur["sales"] * dv("ebitda_margin", t)
        cur["opex"] = cur["gross_profit"] - cur["ebitda"]
        cur["other_income"] = cur["sales"] * dv("other_income_pct", t)
        if drivers.dep_basis == "gross":
            cur["depreciation"] = prev["gross_block"] * dv("dep_rate", t)
        else:
            cur["depreciation"] = prev["fixed_assets"] * dv("dep_rate", t)
        cur["ebit"] = cur["ebitda"] + cur["other_income"] - cur["depreciation"]
        cur["interest"] = prev["borrowings"] * dv("interest_rate", t)
        cur["pbt"] = cur["ebit"] - cur["interest"]
        cur["tax"] = max(cur["pbt"], 0.0) * dv("tax_rate", t)
        cur["minority_other"] = (cur["pbt"] - cur["tax"]) * dv("minority_pct", t)
        cur["net_profit"] = cur["pbt"] - cur["tax"] - cur["minority_other"]
        cur["dividends"] = max(cur["net_profit"], 0.0) * dv("dividend_payout", t)
        cur["capex"] = cur["sales"] * dv("capex_pct", t)
        cur["cwip_transfer"] = min(prev["cwip"], cwip0 / cwip_years) if cwip0 > 0 else 0.0
        cur["cwip"] = prev["cwip"] - cur["cwip_transfer"]
        added = cur["capex"] + cur["cwip_transfer"]
        cur["gross_block"] = (prev["gross_block"] + added) if prev["gross_block"] is not None else None
        cur["acc_dep"] = (prev["acc_dep"] + cur["depreciation"]) if prev["acc_dep"] is not None else None
        cur["fixed_assets"] = prev["fixed_assets"] + added - cur["depreciation"]
        for k in ("investments", "minority_interest", "borrowings", "lt_borrowings", "st_borrowings",
                  "lease_liabilities", "other_borrowings", "equity_capital", "loans_advances"):
            cur[k] = prev[k]
        cur["dso"], cur["dio"], cur["dpo"] = dv("dso", t), dv("dio", t), dv("dpo", t)
        cur["receivables"] = cur["sales"] * cur["dso"] / 365
        cur["inventory"] = cur["cogs"] * cur["dio"] / 365
        pay_base = cur["cogs"] if drivers.dpo_basis == "cogs" else cur["cogs"] + cur["opex"]
        cur["trade_payables"] = pay_base * cur["dpo"] / 365
        cur["other_ca"] = cur["sales"] * dv("other_ca_pct", t)
        cur["other_liab"] = cur["sales"] * dv("other_liab_pct", t)
        cur["nwc"] = operating_nwc(cur)
        cur["delta_nwc"] = cur["nwc"] - prev["nwc"]
        cur["reserves"] = prev["reserves"] + cur["net_profit"] - cur["dividends"]
        cur["total_equity"] = cur["equity_capital"] + cur["reserves"]
        cur["cfo"] = cur["net_profit"] + cur["depreciation"] - cur["delta_nwc"]
        cur["cfi"] = -cur["capex"]
        cur["cff"] = -cur["dividends"]
        cur["net_cash_flow"] = cur["cfo"] + cur["cfi"] + cur["cff"]
        cur["cash"] = prev["cash"] + cur["net_cash_flow"]
        cur["total_assets"] = (cur["fixed_assets"] + cur["cwip"] + cur["investments"] + cur["receivables"]
                               + cur["inventory"] + cur["cash"] + cur["other_ca"] + cur["loans_advances"])
        cur["total_liabilities"] = (cur["equity_capital"] + cur["reserves"] + cur["borrowings"] + cur["trade_payables"]
                                    + cur["other_liab"] + cur["minority_interest"])
        cur["balance_check"] = cur["total_assets"] - cur["total_liabilities"]
        cur["ebit_operating"] = cur["ebitda"] - cur["depreciation"]
        cur["nopat"] = cur["ebit_operating"] * (1 - dv("tax_rate", t))
        cur["fcff"] = cur["nopat"] + cur["depreciation"] - cur["capex"] - cur["delta_nwc"]
        cur["invested_capital"] = invested_capital(cur)
        for k in keys:
            out[k].append(cur[k])
        prev = cur
    return Forecast(scenario=scenario, years=years, lines=out)


def actual_balance_check(fin: CompanyFinancials) -> List[float]:
    L = fin.lines
    return [a - l for a, l in zip(L["total_assets"], L["total_liabilities"])]


def _main(argv):
    import argparse

    from .financials import load_financials
    from .screener_client import ScreenerClient

    ap = argparse.ArgumentParser(description="Print the 5-year forecast and balance checks.")
    ap.add_argument("symbol")
    ap.add_argument("--scenario", default="base", choices=["base", "bull", "bear"])
    ap.add_argument("--cache-dir", default="cache")
    ap.add_argument("--offline", action="store_true")
    a = ap.parse_args(argv)
    fin = load_financials(a.symbol, ScreenerClient(cache_dir=a.cache_dir, offline=a.offline))
    drv = build_drivers(fin)
    fc = run_forecast(fin, drv, a.scenario)
    print(f"{fin.name} - {a.scenario} scenario")
    for dr in drv:
        print(f"  {dr.label:<40} {dr.value(a.scenario):.4f}  ({dr.basis})")
    print("  " + " ".join(f"{lab:>10}" for lab in ["line"] + fc.year_labels))
    for k in ("sales", "ebitda", "net_profit", "cash", "total_assets", "total_liabilities", "balance_check", "fcff"):
        print(f"  {k:>10} " + " ".join(f"{v:>10,.1f}" for v in fc.lines[k]))
    return 0


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
