"""Forecast drivers derived from the company's own reported history.

Each driver is the 3-year average (or 3-year CAGR for sales) of the actual ratio,
clamped to a sensible range. Bull / Bear move sales growth by +/- 3 pp, the
EBITDA margin by +/- 1.5 pp and capex intensity by x0.9 / x1.1 (config
``scenarios``). Every value is written to the Drivers sheet where the user can
edit it.

Per-year values: working-capital days glide from the last reported year to the
driver value over ``wc_glide_years`` (so year 1 has no artificial jump), and a
sourced analyst override (``analyst_inputs.json``) replaces the base value for a
given forecast year while keeping the scenario offset.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Optional

from .config import with_defaults
from .financials import SERVICE_COGS_SHARE, CompanyFinancials

SCENARIOS = ("base", "bull", "bear")
GROWTH_OFFSET = 0.03
MARGIN_OFFSET = 0.015
WINDOW = 3
GLIDE_KEYS = ("dso", "dio", "dpo")


@dataclass
class Driver:
    key: str
    label: str
    unit: str  # "pct" or "days"
    historical: Optional[float]
    base: float
    bull: float
    bear: float
    basis: str
    clamped: bool = False
    raw: Optional[float] = None  # value before clamping

    def value(self, scenario: str) -> float:
        return {"base": self.base, "bull": self.bull, "bear": self.bear}[scenario]


@dataclass
class Override:
    """A sourced analyst value for one forecast year (index 0 = first forecast year)."""
    value: float
    source: str


@dataclass
class Drivers:
    items: Dict[str, Driver]
    dep_basis: str  # "gross" (on opening gross block) or "net" (on opening net block)
    window_years: List[int]
    notes: List[str] = field(default_factory=list)
    dpo_basis: str = "cogs"  # "cogs" or "expenses" (COGS + operating expenses)
    glide_from: Dict[str, float] = field(default_factory=dict)  # last reported days, per GLIDE_KEYS
    glide_years: int = 2
    overrides: Dict[str, Dict[int, Override]] = field(default_factory=dict)

    def value(self, key: str, scenario: str = "base", t: Optional[int] = None) -> float:
        """Driver value; with ``t`` (forecast year index) the per-year value incl. glide and overrides."""
        item = self.items[key]
        v = item.value(scenario)
        if t is None:
            return v
        ov = self.overrides.get(key, {}).get(t)
        if ov is not None:
            return ov.value + (v - item.base)
        if key in self.glide_from:
            g0 = self.glide_from[key]
            w = min(1.0, (t + 1) / max(1, self.glide_years))
            return g0 + (v - g0) * w
        return v

    def is_overridden(self, key: str, t: int) -> bool:
        return t in self.overrides.get(key, {})

    def __iter__(self):
        return iter(self.items.values())


def _clamp(v: float, lo: float, hi: float):
    c = min(max(v, lo), hi)
    return c, abs(c - v) > 1e-12


def _avg(vals: List[float]) -> Optional[float]:
    vals = [v for v in vals if v is not None]
    return sum(vals) / len(vals) if vals else None


def payables_basis(fin: CompanyFinancials) -> str:
    """"cogs" when COGS averages at least SERVICE_COGS_SHARE of total expenses over the driver window, else "expenses"."""
    L = fin.lines
    n = fin.n
    idx = range(n - min(WINDOW, n - 1), n)
    share = _avg([L["cogs"][k] / L["expenses"][k] for k in idx if L["expenses"][k] > 0]) or 0.0
    return "cogs" if share >= SERVICE_COGS_SHARE else "expenses"


def base_year_days(fin: CompanyFinancials, dpo_basis: str) -> Dict[str, float]:
    """Last reported DSO / DIO / DPO, computed exactly like the Working Capital sheet."""
    L = fin.lines
    s, c, e = L["sales"][-1], L["cogs"][-1], L["expenses"][-1]
    pay_base = c if dpo_basis == "cogs" else e
    return {
        "dso": L["receivables"][-1] / s * 365 if s else 0.0,
        "dio": L["inventory"][-1] / c * 365 if c > 0 else 0.0,
        "dpo": L["trade_payables"][-1] / pay_base * 365 if pay_base > 0 else 0.0,
    }


def build_drivers(fin: CompanyFinancials, cfg: Optional[dict] = None) -> Drivers:
    cfg = with_defaults(cfg)
    sc = cfg["scenarios"]
    g_off, m_off = sc.get("growth_offset", GROWTH_OFFSET), sc.get("margin_offset", MARGIN_OFFSET)
    capex_mult = sc.get("capex_mult", {"bull": 1.0, "bear": 1.0})
    L = fin.lines
    n = fin.n
    w = min(WINDOW, n - 1)
    idx = list(range(n - w, n))
    notes: List[str] = []
    items: Dict[str, Driver] = {}

    def add(key, label, unit, hist, lo, hi, basis, offset=0.0, default=None):
        raw = hist if hist is not None else default
        if raw is None:
            raw = 0.0
        base, clamped = _clamp(raw, lo, hi)
        if clamped:
            notes.append(f"{label}: historical {raw:.4f} clamped to {base:.4f}.")
        if hist is None and default is not None:
            notes.append(f"{label}: no usable history, default {default:.4f} used.")
        items[key] = Driver(key, label, unit, hist, base, base + offset, base - offset, basis, clamped, raw)

    s = L["sales"]
    start = s[n - 1 - w]
    cagr = (s[-1] / start) ** (1.0 / w) - 1 if start > 0 and s[-1] > 0 else None
    add("sales_growth", "Revenue growth (YoY)", "pct", cagr, -0.10, 0.30,
        f"{w}-yr revenue CAGR FY{fin.years[n-1-w] % 100:02d}-FY{fin.years[-1] % 100:02d}", g_off, default=0.0)
    cost_items = ("material + manufacturing + employee cost; service cost structure" if fin.cost_basis == "services"
                  else "material + manufacturing costs")
    add("cogs_pct", "COGS (cost of revenue) % of revenue", "pct", _avg([L["cogs"][k] / s[k] for k in idx if s[k]]),
        0.0, 0.95, f"{w}-yr average ({cost_items})")
    add("ebitda_margin", "EBITDA margin", "pct", _avg([L["ebitda"][k] / s[k] for k in idx if s[k]]), -0.20, 0.60,
        f"{w}-yr average operating profit margin", m_off)
    # keep OPEX >= 0 in every scenario
    gm = 1 - items["cogs_pct"].base
    d = items["ebitda_margin"]
    if d.bull > gm:
        notes.append("EBITDA margin capped at gross margin so operating expenses stay positive.")
        d.bull = gm
        d.base = min(d.base, gm)
    add("other_income_pct", "Other income % of revenue", "pct",
        _avg([L["other_income_normal"][k] / s[k] for k in idx if s[k]]), 0.0, 0.15,
        f"{w}-yr average, excluding exceptional items")
    rates = []
    for k in idx:
        avg_b = (L["borrowings"][k] + L["borrowings"][k - 1]) / 2 if k > 0 else L["borrowings"][k]
        if avg_b > 1:
            rates.append(L["interest"][k] / avg_b)
    add("interest_rate", "Interest rate on borrowings", "pct", _avg(rates), 0.04, 0.15,
        f"{w}-yr average interest / average borrowings (incl. leases)", default=0.09)
    if fin.gross_block is not None:
        dep_basis = "gross"
        dr = [L["depreciation"][k] / fin.gross_block[k - 1] for k in idx if k > 0 and fin.gross_block[k - 1] > 0]
        label = "Depreciation % of opening gross block"
    else:
        dep_basis = "net"
        dr = [L["depreciation"][k] / L["fixed_assets"][k - 1] for k in idx if k > 0 and L["fixed_assets"][k - 1] > 0]
        label = "Depreciation % of opening net block"
    add("dep_rate", label, "pct", _avg(dr), 0.01, 0.50, f"{w}-yr average")
    taxes = [L["tax_rate"][k] for k in idx if L["pbt"][k] > 0]
    add("tax_rate", "Effective tax rate", "pct", _avg(taxes), 0.0, 0.40,
        f"{w}-yr average of reported tax %", default=0.2517)
    mins = [L["minority_other"][k] / (L["pbt"][k] - L["tax"][k]) for k in idx if L["pbt"][k] - L["tax"][k] > 0]
    add("minority_pct", "Minority interest & associates % of profit after tax", "pct", _avg(mins), -0.50, 0.50,
        f"{w}-yr average of (PBT - tax - reported net profit) / (PBT - tax)", default=0.0)
    capex_hist = _avg([L["capex"][k] / s[k] for k in idx if s[k]])
    add("capex_pct", "Capex % of revenue", "pct", capex_hist, 0.0, 0.50,
        f"{w}-yr average fixed assets purchased / revenue; bull x{capex_mult.get('bull', 1)}, "
        f"bear x{capex_mult.get('bear', 1)}")
    cx = items["capex_pct"]
    cx.bull, cx.bear = cx.base * capex_mult.get("bull", 1.0), cx.base * capex_mult.get("bear", 1.0)
    add("dso", "Receivable days (DSO)", "days", _avg([L["receivables"][k] / s[k] * 365 for k in idx if s[k]]), 0, 365,
        f"{w}-yr average receivables / revenue x 365")
    c = L["cogs"]
    add("dio", "Inventory days (DIO)", "days", _avg([L["inventory"][k] / c[k] * 365 for k in idx if c[k] > 0]), 0, 365,
        f"{w}-yr average inventory / COGS x 365", default=0.0)
    # When COGS is still a small part of costs (no employee line to add for a service business), payables / COGS
    # explodes, so payable days are measured on total expenses instead.
    dpo_basis = payables_basis(fin)
    base_cost = c if dpo_basis == "cogs" else L["expenses"]
    add("dpo", "Payable days (DPO)" + ("" if dpo_basis == "cogs" else " on total expenses"), "days",
        _avg([L["trade_payables"][k] / base_cost[k] * 365 for k in idx if base_cost[k] > 0]), 0, 365,
        f"{w}-yr average trade payables / "
        f"{'COGS' if dpo_basis == 'cogs' else f'total expenses (COGS < {SERVICE_COGS_SHARE:.0%} of costs)'} x 365",
        default=0.0)
    add("dividend_payout", "Dividend payout", "pct", _avg([L["dividend_payout"][k] for k in idx]), 0.0, 1.0,
        f"{w}-yr average reported payout")
    oca = L["other_asset_items"][-1] / s[-1]
    add("other_ca_pct", "Other operating assets % of revenue", "pct", oca, 0.0, 1.0,
        "Last reported year (other asset items; loans & advances are held flat as non-operating)")
    ol = (L["other_liability_items"][-1] + L["advance_from_customers"][-1]) / s[-1]
    add("other_liab_pct", "Other liabilities % of revenue", "pct", ol, 0.0, 3.0,
        "Last reported year (other liability items + customer advances)")
    share = cfg["ronic_premium_share"]
    items["ronic_share"] = Driver("ronic_share", "Terminal RONIC: share of ROIC premium over WACC kept", "pct", None,
                                  share["base"], share["bull"], share["bear"],
                                  "Terminal return on new capital = WACC + share x (current ROIC - WACC)")
    glide_from = base_year_days(fin, dpo_basis)
    return Drivers(items=items, dep_basis=dep_basis, window_years=[fin.years[k] for k in idx], notes=notes,
                   dpo_basis=dpo_basis, glide_from=glide_from, glide_years=int(cfg.get("wc_glide_years", 2)))


DRIVER_ORDER = [
    "sales_growth", "cogs_pct", "ebitda_margin", "other_income_pct", "interest_rate", "dep_rate", "tax_rate",
    "minority_pct", "capex_pct", "dso", "dio", "dpo", "dividend_payout", "other_ca_pct", "other_liab_pct", "ronic_share",
]
# drivers with a per-forecast-year value (everything except valuation-only inputs)
YEAR_DRIVERS = [k for k in DRIVER_ORDER if k != "ronic_share"]
