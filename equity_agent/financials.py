"""Turn raw Screener tables + schedules into one validated ``CompanyFinancials``.

Every number here is a reported Screener.in figure (₹ Cr) or an exact arithmetic
combination of reported figures. Nothing is estimated. Where a line item is not
reported, the run either stops (``DataValidationError``) or records a warning;
it never invents a value.
"""
from __future__ import annotations

import json
import re
import sys
from dataclasses import asdict, dataclass, field
from typing import Dict, List, Optional

from .errors import DataValidationError
from .screener_client import RawCompany, ScheduleItem, ScreenerClient, Table, parse_number

COGS_KEYWORDS = ("material", "raw", "purchase", "stock", "inventor", "power", "fuel", "manufactur", "mfr", "excise")
# Cost of revenue. For a goods business it is the material and manufacturing items of Screener's expense schedule.
# When those items average less than SERVICE_COGS_SHARE of total expenses over the last COST_BASIS_WINDOW years
# (IT, engineering and other service businesses), the staff who deliver the service are the main cost of revenue,
# so employee cost is counted in cost of revenue as well. The rule uses each company's own cost mix.
SERVICE_COGS_SHARE = 0.30
COST_BASIS_WINDOW = 3
EMPLOYEE_KEYWORDS = ("employee",)
_MONTHS = {m: i for i, m in enumerate(["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"], 1)}


@dataclass
class ExpenseItem:
    name: str
    kind: str  # "COGS" or "OPEX"
    values: List[float]


@dataclass
class CompanyFinancials:
    symbol: str
    name: str
    consolidated: bool
    company_id: str
    website: Optional[str]
    bse_code: Optional[str]
    logo_url: Optional[str]
    sector: str
    industry: str
    source_url: str
    fetched_at: str
    # market data (Screener top ratios)
    price: float
    market_cap: float
    high_52: Optional[float]
    low_52: Optional[float]
    pe: Optional[float]
    book_value: Optional[float]
    dividend_yield: Optional[float]
    roce: Optional[float]
    roe: Optional[float]
    face_value: Optional[float]
    # annual data
    years: List[int]
    period_labels: List[str]
    lines: Dict[str, List[float]]
    expense_items: List[ExpenseItem]
    eps: List[Optional[float]]
    gross_block: Optional[List[float]]
    acc_dep: Optional[List[float]]
    reported_ratios: Dict[str, List[Optional[float]]]
    ttm: Dict[str, Optional[float]]
    quarters: Dict[str, List[Optional[float]]]
    quarter_labels: List[str]
    warnings: List[str] = field(default_factory=list)
    shareholding: Dict[str, List[Optional[float]]] = field(default_factory=dict)  # % by holder class, by quarter
    shareholding_labels: List[str] = field(default_factory=list)
    cost_basis: str = "goods"  # "goods" (material + manufacturing) or "services" (+ employee cost)

    # -- helpers ------------------------------------------------------------ #
    def line(self, key: str) -> List[float]:
        return self.lines[key]

    def latest(self, key: str) -> float:
        return self.lines[key][-1]

    @property
    def n(self) -> int:
        return len(self.years)

    @property
    def year_labels(self) -> List[str]:
        return [f"FY{y % 100:02d}A" for y in self.years]

    @property
    def shares_cr(self) -> float:
        """Shares outstanding in crore = market cap (₹ Cr) / price (₹)."""
        return self.market_cap / self.price

    def to_dict(self) -> dict:
        return asdict(self)

    def quarter_fy(self, label: str) -> Optional[int]:
        """Fiscal year a quarter label ('Sep 2026') belongs to, using the annual year-end month."""
        m = re.match(r"([A-Za-z]{3})\w*\s+(\d{4})", label.strip())
        if not m:
            return None
        mon, yr = _MONTHS.get(m.group(1)[:3].title()), int(m.group(2))
        end = _MONTHS.get(self.period_labels[-1][:3].title(), 3)
        if mon is None:
            return None
        return yr if mon <= end else yr + 1

    def ytd_growth(self, key: str = "sales"):
        """Year-to-date growth of the first forecast year from reported quarters.

        Returns (quarters_reported, growth) comparing the quarters of FY(last reported + 1) with the same
        quarters a year earlier, or (0, None) when none are reported.
        """
        q = self.quarters.get(key)
        if not q or not self.quarter_labels:
            return 0, None
        fy1 = self.years[-1] + 1
        idx = {lab: i for i, lab in enumerate(self.quarter_labels)}
        cur = prev = 0.0
        n = 0
        for lab, i in idx.items():
            if self.quarter_fy(lab) != fy1 or q[i] is None:
                continue
            m = re.match(r"([A-Za-z]{3})\w*\s+(\d{4})", lab.strip())
            prior = f"{m.group(1)} {int(m.group(2)) - 1}"
            j = idx.get(prior)
            if j is None or q[j] is None:
                continue
            cur += q[i]
            prev += q[j]
            n += 1
        if n == 0 or prev <= 0:
            return 0, None
        return n, cur / prev - 1


# --------------------------------------------------------------------------- #
# helpers
# --------------------------------------------------------------------------- #
def _period_year(label: str) -> Optional[int]:
    m = re.match(r"([A-Za-z]{3})\w*\s+(\d{4})", label.strip())
    if not m:
        return None
    return int(m.group(2))


def _tol(sales: float) -> float:
    return max(2.0, abs(sales) * 0.005)


def _row(table: Table, label: str, idx: List[int], required: bool = True, warn: Optional[list] = None) -> List[float]:
    if label not in table.rows:
        if required:
            raise DataValidationError(f"Screener table is missing the '{label}' row.")
        if warn is not None:
            warn.append(f"'{label}' row not reported by Screener; treated as 0.")
        return [0.0] * len(idx)
    vals = table.rows[label]
    return [float(vals[i]) if i is not None and i < len(vals) and vals[i] is not None else 0.0 for i in idx]


def _sched_series(items: Dict[str, ScheduleItem], name: str, labels: List[str], sales: Optional[List[float]] = None) -> Optional[List[float]]:
    """Return ₹ Cr values of a schedule item aligned to ``labels`` (None if item absent)."""
    key = None
    for k in items:
        if k.lower() == name.lower():
            key = k
            break
    if key is None:
        return None
    it = items[key]
    out = []
    for i, lab in enumerate(labels):
        v = None
        if it.amounts is not None:
            v = it.amounts.get(lab)
        if v is None and not it.is_percent:
            v = it.values.get(lab)
        if v is None and it.is_percent and sales is not None and it.values.get(lab) is not None:
            v = it.values[lab] / 100.0 * sales[i]
        out.append(float(v) if v is not None else 0.0)
    return out


def _top(raw: RawCompany, key: str) -> Optional[str]:
    for k, v in raw.top_ratios.items():
        if k.lower() == key.lower():
            return v
    return None


def classify_expense(name: str) -> str:
    low = name.lower()
    return "COGS" if any(k in low for k in COGS_KEYWORDS) else "OPEX"


def goods_cost_share(items: List["ExpenseItem"], expenses: List[float], window: int = COST_BASIS_WINDOW) -> Optional[float]:
    """Average share of material and manufacturing items in total expenses over the last ``window`` years."""
    n = len(expenses)
    shares = [sum(it.values[k] for it in items if classify_expense(it.name) == "COGS") / expenses[k]
              for k in range(max(0, n - window), n) if expenses[k] > 0]
    return sum(shares) / len(shares) if shares else None


def apply_cost_basis(items: List["ExpenseItem"], expenses: List[float]) -> str:
    """Set each item's kind for cost of revenue and return the basis ("goods" or "services")."""
    share = goods_cost_share(items, expenses)
    staff = {id(it) for it in items if any(k in it.name.lower() for k in EMPLOYEE_KEYWORDS)}
    services = share is not None and share < SERVICE_COGS_SHARE and bool(staff)
    for it in items:
        it.kind = "COGS" if classify_expense(it.name) == "COGS" or (services and id(it) in staff) else "OPEX"
    return "services" if services else "goods"


def cost_of_revenue_label(fin: "CompanyFinancials", short: bool = False) -> str:
    if fin.cost_basis == "services":
        return "Cost of revenue" if short else "Cost of revenue (material + manufacturing + employee cost)"
    return "Cost of goods sold" if short else "Cost of goods sold (material + manufacturing)"


def opex_label(fin: "CompanyFinancials") -> str:
    return "Operating expenses (other)" if fin.cost_basis == "services" else "Operating expenses (employee + other)"


# --------------------------------------------------------------------------- #
# build
# --------------------------------------------------------------------------- #
def build_financials(raw: RawCompany) -> CompanyFinancials:
    warnings: List[str] = []
    pl = raw.tables.get("profit-loss")
    bs = raw.tables.get("balance-sheet")
    cf = raw.tables.get("cash-flow")
    if pl is None or bs is None:
        raise DataValidationError(f"{raw.symbol}: Screener page lacks Profit & Loss or Balance Sheet tables.")

    # ---- choose annual columns present in P&L and BS -------------------- #
    pl_annual = [(i, h) for i, h in enumerate(pl.headers) if h.upper() != "TTM" and _period_year(h)]
    bs_index = {h: i for i, h in enumerate(bs.headers)}
    cf_index = {h: i for i, h in enumerate(cf.headers)} if cf is not None else {}
    chosen = []
    sales_row = pl.rows.get("Sales")
    if sales_row is None:
        raise DataValidationError(f"{raw.symbol}: no 'Sales' row on Screener P&L.")
    for i, h in pl_annual:
        if h not in bs_index:
            continue
        if sales_row[i] is None:
            continue
        chosen.append((i, h))
    # drop duplicate fiscal years (year-end change): keep the later column
    by_year: Dict[int, tuple] = {}
    for i, h in chosen:
        y = _period_year(h)
        if y in by_year:
            warnings.append(f"Two annual columns map to FY{y} ('{by_year[y][1]}' and '{h}'); kept '{h}'.")
        by_year[y] = (i, h)
    chosen = [by_year[y] for y in sorted(by_year)]
    if len(chosen) < 3:
        raise DataValidationError(f"{raw.symbol}: need at least 3 annual years of data, Screener has {len(chosen)}.")
    labels = [h for _, h in chosen]
    years = [_period_year(h) for h in labels]
    pl_idx = [i for i, _ in chosen]
    bs_idx = [bs_index[h] for h in labels]
    cf_idx = [cf_index.get(h) for h in labels]
    if cf is None or any(i is None for i in cf_idx):
        missing = [labels[k] for k, i in enumerate(cf_idx) if i is None] if cf is not None else labels
        warnings.append(f"Cash flow not reported for {', '.join(missing)}; those cash-flow cells are 0.")

    # ---- P&L --------------------------------------------------------------- #
    L: Dict[str, List[float]] = {}
    L["sales"] = _row(pl, "Sales", pl_idx)
    L["expenses"] = _row(pl, "Expenses", pl_idx)
    L["ebitda"] = _row(pl, "Operating Profit", pl_idx)
    L["other_income"] = _row(pl, "Other Income", pl_idx, required=False, warn=warnings)
    L["interest"] = _row(pl, "Interest", pl_idx, required=False, warn=warnings)
    L["depreciation"] = _row(pl, "Depreciation", pl_idx, required=False, warn=warnings)
    L["pbt"] = _row(pl, "Profit before tax", pl_idx)
    tax_pct = _row(pl, "Tax %", pl_idx, required=False, warn=warnings)
    L["tax_rate"] = [t / 100.0 for t in tax_pct]
    L["net_profit"] = _row(pl, "Net Profit", pl_idx)
    L["dividend_payout"] = [v / 100.0 for v in _row(pl, "Dividend Payout %", pl_idx, required=False)]
    L["tax"] = [round(p * t, 2) for p, t in zip(L["pbt"], L["tax_rate"])]
    # Deduction between (PBT - tax) and reported net profit: minority share, associates, tax-% rounding.
    L["minority_other"] = [round((p - t) - npf, 2) for npf, p, t in zip(L["net_profit"], L["pbt"], L["tax"])]
    eps_row = pl.rows.get("EPS in Rs")
    eps = [eps_row[i] if eps_row and i < len(eps_row) else None for i in pl_idx]

    oi_items = raw.schedule("profit-loss", pl.schedules.get("Other Income", "Other Income"))
    oi_normal = _sched_series(oi_items, "Other income normal", labels)
    exceptional = _sched_series(oi_items, "Exceptional items", labels)
    L["other_income_normal"] = oi_normal if oi_normal is not None else list(L["other_income"])
    L["exceptional_items"] = exceptional if exceptional is not None else [0.0] * len(labels)

    # ---- expense breakdown ------------------------------------------------- #
    exp_items_raw = raw.schedule("profit-loss", pl.schedules.get("Expenses", "Expenses"))
    expense_items: List[ExpenseItem] = []
    for name in exp_items_raw:
        vals = _sched_series(exp_items_raw, name, labels, L["sales"])
        if vals is None or all(abs(v) < 1e-9 for v in vals):
            continue
        expense_items.append(ExpenseItem(name=name, kind=classify_expense(name), values=vals))
    if not expense_items:
        warnings.append("Screener has no expense breakdown; COGS shown as 0 and all expenses as operating expenses.")
        expense_items = [ExpenseItem("Total expenses", "OPEX", list(L["expenses"]))]
    cost_basis = apply_cost_basis(expense_items, L["expenses"])
    L["cogs"] = [round(sum(it.values[k] for it in expense_items if it.kind == "COGS"), 2) for k in range(len(labels))]
    L["opex"] = [round(e - c, 2) for e, c in zip(L["expenses"], L["cogs"])]
    L["gross_profit"] = [round(s - c, 2) for s, c in zip(L["sales"], L["cogs"])]
    L["employee_cost"] = _sched_series(exp_items_raw, "Employee Cost", labels, L["sales"]) or [0.0] * len(labels)

    # ---- balance sheet ---------------------------------------------------- #
    L["equity_capital"] = _row(bs, "Equity Capital", bs_idx)
    L["reserves"] = _row(bs, "Reserves", bs_idx)
    L["borrowings"] = _row(bs, "Borrowings", bs_idx, required=False, warn=warnings)
    L["other_liabilities"] = _row(bs, "Other Liabilities", bs_idx)
    L["total_liabilities"] = _row(bs, "Total Liabilities", bs_idx)
    L["fixed_assets"] = _row(bs, "Fixed Assets", bs_idx)
    L["cwip"] = _row(bs, "CWIP", bs_idx, required=False)
    L["investments"] = _row(bs, "Investments", bs_idx, required=False)
    L["other_assets"] = _row(bs, "Other Assets", bs_idx)
    L["total_assets"] = _row(bs, "Total Assets", bs_idx)
    L["total_equity"] = [a + b for a, b in zip(L["equity_capital"], L["reserves"])]

    borr = raw.schedule("balance-sheet", bs.schedules.get("Borrowings", "Borrowings"))
    if borr:
        L["lt_borrowings"] = _sched_series(borr, "Long term Borrowings", labels) or [0.0] * len(labels)
        L["st_borrowings"] = _sched_series(borr, "Short term Borrowings", labels) or [0.0] * len(labels)
        L["lease_liabilities"] = _sched_series(borr, "Lease Liabilities", labels) or [0.0] * len(labels)
        L["other_borrowings"] = [round(b - lt - st - le, 2) for b, lt, st, le in
                                 zip(L["borrowings"], L["lt_borrowings"], L["st_borrowings"], L["lease_liabilities"])]
    else:
        if any(L["borrowings"]):
            warnings.append("No borrowings breakdown on Screener; all borrowings shown as long-term.")
        L["lt_borrowings"] = list(L["borrowings"])
        L["st_borrowings"] = [0.0] * len(labels)
        L["lease_liabilities"] = [0.0] * len(labels)
        L["other_borrowings"] = [0.0] * len(labels)

    ol = raw.schedule("balance-sheet", bs.schedules.get("Other Liabilities", "Other Liabilities"))
    if not ol:
        raise DataValidationError(f"{raw.symbol}: Screener has no 'Other Liabilities' breakdown (trade payables needed).")
    L["trade_payables"] = _sched_series(ol, "Trade Payables", labels) or [0.0] * len(labels)
    L["advance_from_customers"] = _sched_series(ol, "Advance from Customers", labels) or [0.0] * len(labels)
    L["minority_interest"] = _sched_series(ol, "Non controlling int", labels) or [0.0] * len(labels)
    reported_oli = _sched_series(ol, "Other liability items", labels)
    # Residual against the reported *total*, so the components add up to Total Liabilities exactly
    # (Screener rounds parent rows and totals separately; the gap is at most a few crore).
    L["other_liability_items"] = [round(tl - ec - rs - b - p - a - m, 2) for tl, ec, rs, b, p, a, m in
                                  zip(L["total_liabilities"], L["equity_capital"], L["reserves"], L["borrowings"],
                                      L["trade_payables"], L["advance_from_customers"], L["minority_interest"])]

    oa = raw.schedule("balance-sheet", bs.schedules.get("Other Assets", "Other Assets"))
    if not oa:
        raise DataValidationError(f"{raw.symbol}: Screener has no 'Other Assets' breakdown (receivables, inventory, cash needed).")
    L["inventory"] = _sched_series(oa, "Inventories", labels) or [0.0] * len(labels)
    L["receivables"] = _sched_series(oa, "Trade receivables", labels) or [0.0] * len(labels)
    L["cash"] = _sched_series(oa, "Cash Equivalents", labels) or [0.0] * len(labels)
    L["loans_advances"] = _sched_series(oa, "Loans n Advances", labels) or [0.0] * len(labels)
    reported_oai = _sched_series(oa, "Other asset items", labels)
    L["other_asset_items"] = [round(ta - fx - cw - iv - i - r - c - la, 2) for ta, fx, cw, iv, i, r, c, la in
                              zip(L["total_assets"], L["fixed_assets"], L["cwip"], L["investments"], L["inventory"],
                                  L["receivables"], L["cash"], L["loans_advances"])]

    fa = raw.schedule("balance-sheet", bs.schedules.get("Fixed Assets", "Fixed Assets"))
    gross_block = _sched_series(fa, "Gross Block", labels) if fa else None
    acc_dep = _sched_series(fa, "Accumulated Depreciation", labels) if fa else None
    if gross_block is None or acc_dep is None or any(g <= 0 for g in gross_block):
        warnings.append("Gross block / accumulated depreciation not fully reported; PP&E schedule uses net block only.")
        gross_block = acc_dep = None

    # ---- cash flow --------------------------------------------------------- #
    def cf_row(label: str, required: bool = True) -> List[float]:
        if cf is None or label not in cf.rows:
            if required:
                warnings.append(f"Cash-flow row '{label}' not reported; treated as 0.")
            return [0.0] * len(labels)
        vals = cf.rows[label]
        return [float(vals[i]) if i is not None and i < len(vals) and vals[i] is not None else 0.0 for i in cf_idx]

    L["cfo"] = cf_row("Cash from Operating Activity")
    L["cfi"] = cf_row("Cash from Investing Activity")
    L["cff"] = cf_row("Cash from Financing Activity")
    L["net_cash_flow"] = cf_row("Net Cash Flow")
    cfi_items = raw.schedule("cash-flow", cf.schedules.get("Cash from Investing Activity", "")) if cf is not None else {}
    capex = _sched_series(cfi_items, "Fixed assets purchased", [labels[k] for k in range(len(labels))]) if cfi_items else None
    fcf_rep = cf_row("Free Cash Flow", required=False)
    if capex is not None:
        L["capex"] = [abs(v) for v in capex]
    else:
        warnings.append("Capex ('Fixed assets purchased') not reported; derived as CFO - Free Cash Flow.")
        L["capex"] = [max(0.0, round(o - f, 2)) for o, f in zip(L["cfo"], fcf_rep)]
    L["fcf"] = fcf_rep if any(fcf_rep) else [round(o - c, 2) for o, c in zip(L["cfo"], L["capex"])]
    cff_items = raw.schedule("cash-flow", cf.schedules.get("Cash from Financing Activity", "")) if cf is not None else {}
    div = _sched_series(cff_items, "Dividends paid", labels) if cff_items else None
    L["dividends_paid"] = [abs(v) for v in div] if div is not None else [0.0] * len(labels)

    # ---- reported ratios (cross-check only) -------------------------------- #
    ratios = raw.tables.get("ratios")
    rep: Dict[str, List[Optional[float]]] = {}
    if ratios is not None:
        r_index = {h: i for i, h in enumerate(ratios.headers)}
        for lab in ("Debtor Days", "Inventory Days", "Days Payable", "Cash Conversion Cycle", "Working Capital Days", "ROCE %"):
            if lab in ratios.rows:
                rep[lab] = [ratios.rows[lab][r_index[h]] if h in r_index else None for h in labels]

    # ---- TTM & quarters ----------------------------------------------------- #
    ttm: Dict[str, Optional[float]] = {}
    if pl.headers and pl.headers[-1].upper() == "TTM":
        t = len(pl.headers) - 1
        for key, lab in (("sales", "Sales"), ("expenses", "Expenses"), ("ebitda", "Operating Profit"),
                         ("net_profit", "Net Profit"), ("eps", "EPS in Rs")):
            ttm[key] = pl.rows[lab][t] if lab in pl.rows else None
    q = raw.tables.get("quarters")
    quarters: Dict[str, List[Optional[float]]] = {}
    q_labels: List[str] = []
    if q is not None:
        q_labels = list(q.headers)
        for key, lab in (("sales", "Sales"), ("expenses", "Expenses"), ("ebitda", "Operating Profit"),
                         ("other_income", "Other Income"), ("interest", "Interest"), ("depreciation", "Depreciation"),
                         ("pbt", "Profit before tax"), ("net_profit", "Net Profit"), ("eps", "EPS in Rs")):
            if lab in q.rows:
                quarters[key] = list(q.rows[lab])

    # ---- shareholding ------------------------------------------------------- #
    shp: Dict[str, List[Optional[float]]] = {}
    shp_labels: List[str] = []
    sh = raw.tables.get("shareholding")
    if sh is not None:
        shp_labels = list(sh.headers)
        for lab in ("Promoters", "FIIs", "DIIs", "Government", "Public", "Others"):
            if lab in sh.rows:
                shp[lab] = list(sh.rows[lab])

    # ---- market data -------------------------------------------------------- #
    price = parse_number(_top(raw, "Current Price"))
    mcap = parse_number(_top(raw, "Market Cap"))
    if not price or not mcap or price <= 0 or mcap <= 0:
        raise DataValidationError(f"{raw.symbol}: Screener did not report a current price and market cap.")
    hl = _top(raw, "High / Low") or ""
    parts = [parse_number(p) for p in hl.split("/")] if "/" in hl else [None, None]

    path = raw.sector_path
    sector = path[1] if len(path) > 1 else (path[0] if path else "Unclassified")
    industry = path[-1] if path else "Unclassified"

    fin = CompanyFinancials(
        symbol=raw.symbol, name=raw.name, consolidated=raw.consolidated, company_id=raw.company_id,
        website=raw.website, bse_code=raw.bse_code, logo_url=raw.logo_url, sector=sector, industry=industry,
        source_url=raw.source_url, fetched_at=raw.fetched_at,
        price=price, market_cap=mcap, high_52=parts[0], low_52=parts[1] if len(parts) > 1 else None,
        pe=parse_number(_top(raw, "Stock P/E")), book_value=parse_number(_top(raw, "Book Value")),
        dividend_yield=parse_number(_top(raw, "Dividend Yield")), roce=parse_number(_top(raw, "ROCE")),
        roe=parse_number(_top(raw, "ROE")), face_value=parse_number(_top(raw, "Face Value")),
        years=years, period_labels=labels, lines=L, expense_items=expense_items, eps=eps,
        gross_block=gross_block, acc_dep=acc_dep, reported_ratios=rep, ttm=ttm,
        quarters=quarters, quarter_labels=q_labels, warnings=warnings, shareholding=shp, shareholding_labels=shp_labels,
        cost_basis=cost_basis,
    )
    validate(fin, reported_oli=reported_oli, reported_oai=reported_oai)
    return fin


# --------------------------------------------------------------------------- #
# validation
# --------------------------------------------------------------------------- #
def validate(fin: CompanyFinancials, reported_oli=None, reported_oai=None) -> List[str]:
    """Check accounting identities. Small rounding gaps become warnings, large gaps stop the run."""
    L = fin.lines
    hard: List[str] = []

    def check(name: str, lhs: List[float], rhs: List[float], hard_pct: float = 0.05):
        for k, y in enumerate(fin.years):
            diff = lhs[k] - rhs[k]
            tol = _tol(L["sales"][k])
            if abs(diff) > tol:
                msg = f"FY{y}: {name} differs by {diff:,.1f} Cr"
                scale = max(abs(rhs[k]), abs(L["sales"][k]), 1.0)
                if abs(diff) > hard_pct * scale:
                    hard.append(msg)
                else:
                    fin.warnings.append(msg + " (rounding/reporting gap, kept as reported)")

    if fin.latest("sales") <= 0:
        hard.append("Latest annual sales are zero or negative.")
    check("Sales - Expenses vs Operating Profit", [s - e for s, e in zip(L["sales"], L["expenses"])], L["ebitda"])
    check("Operating Profit + Other Income - Interest - Depreciation vs PBT",
          [o + oi - i - d for o, oi, i, d in zip(L["ebitda"], L["other_income"], L["interest"], L["depreciation"])], L["pbt"])
    check("Total Assets vs Total Liabilities", L["total_assets"], L["total_liabilities"], hard_pct=0.01)
    check("Asset components vs Total Assets",
          [a + b + c + d for a, b, c, d in zip(L["fixed_assets"], L["cwip"], L["investments"], L["other_assets"])], L["total_assets"])
    check("Liability components vs Total Liabilities",
          [a + b + c + d for a, b, c, d in zip(L["equity_capital"], L["reserves"], L["borrowings"], L["other_liabilities"])],
          L["total_liabilities"])
    if any(L["net_cash_flow"]):
        check("CFO + CFI + CFF vs Net Cash Flow", [a + b + c for a, b, c in zip(L["cfo"], L["cfi"], L["cff"])], L["net_cash_flow"], hard_pct=1.0)
    item_sum = [sum(it.values[k] for it in fin.expense_items) for k in range(fin.n)]
    check("Expense breakdown vs total Expenses", item_sum, L["expenses"], hard_pct=1.0)
    if reported_oai is not None:
        check("Other Assets breakdown vs reported 'Other asset items'", L["other_asset_items"], reported_oai, hard_pct=1.0)
    if reported_oli is not None:
        check("Other Liabilities breakdown vs reported 'Other liability items'", L["other_liability_items"], reported_oli, hard_pct=1.0)
    if fin.face_value and fin.latest("equity_capital") > 0:
        implied = fin.latest("equity_capital") / fin.face_value
        if abs(implied / fin.shares_cr - 1) > 0.05:
            fin.warnings.append(
                f"Share count check: market cap / price = {fin.shares_cr:,.2f} Cr shares vs equity capital / face value "
                f"= {implied:,.2f} Cr; valuation uses market cap / price.")
    if hard:
        raise DataValidationError(f"{fin.symbol}: Screener data failed validation: " + "; ".join(hard))
    return fin.warnings


def load_financials(symbol: str, client: Optional[ScreenerClient] = None) -> CompanyFinancials:
    client = client or ScreenerClient()
    return build_financials(client.fetch_company(symbol))


def _main(argv: List[str]) -> int:
    import argparse
    from pathlib import Path

    ap = argparse.ArgumentParser(description="Write standardized financials JSON + validation report.")
    ap.add_argument("symbol")
    ap.add_argument("--cache-dir", default="cache")
    ap.add_argument("--offline", action="store_true")
    ap.add_argument("--out", default="output")
    args = ap.parse_args(argv)
    fin = load_financials(args.symbol, ScreenerClient(cache_dir=args.cache_dir, offline=args.offline))
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)
    path = out / f"{fin.symbol}_financials.json"
    path.write_text(json.dumps(fin.to_dict(), indent=1, ensure_ascii=False), encoding="utf-8")
    print(f"{fin.name}: {fin.n} years {fin.years[0]}-{fin.years[-1]} -> {path}")
    print("Validation warnings:" if fin.warnings else "Validation: all checks passed.")
    for w in fin.warnings:
        print("  -", w)
    return 0


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
