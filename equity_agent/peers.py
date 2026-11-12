"""Peer comparison from Screener.in.

Peer list: Screener's own peer table (``/api/company/{warehouse_id}/peers/``, the
request the company page makes), or a sourced list from ``analyst_inputs.json``.
For each peer the public company page is read (and its "Other Assets" schedule for
cash) through the same cached, rate-limited client, so every metric uses one set of
definitions:

* EV = market cap + borrowings - cash - investments (minority interest not deducted
  for peers; the same definition is shown for the subject company)
* EV/EBITDA and margins on trailing-twelve-month figures when Screener shows a TTM
  column, else the latest fiscal year
* P/E, ROE, ROCE, dividend yield from Screener's headline ratios; P/B = price / book value

The parsed snapshot is cached as ``<cache>/<SYMBOL>/peers/peers.json``.
"""
from __future__ import annotations

import json
import re
import statistics
import time
from dataclasses import asdict, dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

from bs4 import BeautifulSoup

from .errors import EquityAgentError
from .screener_client import BASE_URL, ScreenerClient, is_bank_layout, parse_company_page, parse_number, parse_schedule

METRICS = ("pe", "ev_ebitda", "pb", "roe", "roce", "sales_growth_3y", "ebitda_margin", "net_margin", "div_yield")
# Multi-year history, computed from each company's Screener annual P&L and balance sheet on one set of definitions
# (fractions): revenue growth; EBITDA (operating profit) and net margins; ROE = net profit / average equity
# (equity capital + reserves); ROCE = (PBT + interest) / average (equity + borrowings).
HISTORY_METRICS = ("sales_growth", "ebitda_margin", "net_margin", "roe", "roce")
HISTORY_YEARS = 5
EV_EBITDA_RANGE = (0.0, 100.0)  # values outside are treated as outliers for the median
PE_RANGE = (0.0, 200.0)


@dataclass
class Peer:
    symbol: str
    name: str
    consolidated: bool
    price: Optional[float] = None
    market_cap: Optional[float] = None
    pe: Optional[float] = None
    pb: Optional[float] = None
    ev: Optional[float] = None
    ev_ebitda: Optional[float] = None
    roe: Optional[float] = None
    roce: Optional[float] = None
    sales: Optional[float] = None
    sales_growth_3y: Optional[float] = None
    ebitda_margin: Optional[float] = None
    net_margin: Optional[float] = None
    div_yield: Optional[float] = None
    basis: str = ""
    notes: List[str] = field(default_factory=list)
    # {metric: {"FY26": value}}, fiscal years labelled by the calendar year they end in; year_end = month ("Mar")
    history: Dict[str, Dict[str, Optional[float]]] = field(default_factory=dict)
    year_end: str = ""


@dataclass
class PeerSet:
    subject: str
    peers: List[Peer]
    source: str
    fetched_at: str
    warnings: List[str] = field(default_factory=list)

    def values(self, metric: str) -> List[float]:
        lo, hi = {"ev_ebitda": EV_EBITDA_RANGE, "pe": PE_RANGE}.get(metric, (-1e9, 1e9))
        out = []
        for p in self.peers:
            v = getattr(p, metric)
            if v is not None and lo < v < hi:
                out.append(v)
        return out

    def median(self, metric: str) -> Optional[float]:
        vals = self.values(metric)
        return statistics.median(vals) if len(vals) >= 2 else None

    def count(self, metric: str) -> int:
        return len(self.values(metric))

    @property
    def has_history(self) -> bool:
        return any(p.history for p in self.peers)

    def history_years(self, n: int = HISTORY_YEARS) -> List[str]:
        """The latest ``n`` fiscal-year labels that at least two peers report."""
        count: Dict[str, int] = {}
        for p in self.peers:
            for y, v in p.history.get("ebitda_margin", {}).items():
                if v is not None:
                    count[y] = count.get(y, 0) + 1
        years = sorted((y for y, k in count.items() if k >= 2), key=lambda s: int(s[2:]))
        return years[-n:]

    def history_median(self, metric: str, year: str) -> Optional[float]:
        vals = [p.history.get(metric, {}).get(year) for p in self.peers]
        vals = [v for v in vals if v is not None]
        return statistics.median(vals) if len(vals) >= 2 else None

    def to_dict(self) -> dict:
        return {"subject": self.subject, "source": self.source, "fetched_at": self.fetched_at,
                "warnings": list(self.warnings), "peers": [asdict(p) for p in self.peers],
                "median": {m: self.median(m) for m in METRICS}}

    @classmethod
    def from_dict(cls, d: dict) -> "PeerSet":
        return cls(subject=d["subject"], peers=[Peer(**p) for p in d.get("peers", [])], source=d.get("source", ""),
                   fetched_at=d.get("fetched_at", ""), warnings=list(d.get("warnings", [])))


def parse_peer_table(html: str) -> List[dict]:
    """Rows of Screener's peer table: [{'symbol', 'name', 'consolidated'}] in Screener's order."""
    soup = BeautifulSoup(html, "html.parser")
    out = []
    for tr in soup.find_all("tr", attrs={"data-row-company-id": True}):
        a = tr.find("a", href=True)
        if a is None:
            continue
        m = re.match(r"^/company/([^/]+)/(consolidated/)?", a["href"])
        if not m:
            continue
        out.append({"symbol": m.group(1).upper(), "name": a.get_text(" ", strip=True), "consolidated": bool(m.group(2))})
    return out


def _cagr(a: Optional[float], b: Optional[float], n: int) -> Optional[float]:
    if a is None or b is None or a <= 0 or b <= 0 or n <= 0:
        return None
    return (b / a) ** (1 / n) - 1


def _last(vals: List[Optional[float]]) -> Optional[float]:
    for v in reversed(vals):
        if v is not None:
            return v
    return None


_PERIOD = re.compile(r"([A-Za-z]{3})\w*\s+(\d{4})")


def ratio_history(labels: List[str], sales, ebitda, net_profit, pbt, interest, equity, borrowings,
                  n: int = HISTORY_YEARS) -> Dict[str, Dict[str, Optional[float]]]:
    """Yearly ratios on the HISTORY_METRICS definitions for the last ``n`` periods (inputs aligned with ``labels``).

    Average balances need the prior year, so ``n + 1`` periods of data give ``n`` years of ratios."""
    div = lambda a, b: a / b if a is not None and b not in (None, 0) else None
    avg = lambda a, b: (a + b) / 2 if a is not None and b is not None else None
    out: Dict[str, Dict[str, Optional[float]]] = {m: {} for m in HISTORY_METRICS}
    start = max(0, len(labels) - n)
    for k in range(start, len(labels)):
        m = _PERIOD.match(labels[k].strip())
        if not m:
            continue
        y = f"FY{int(m.group(2)) % 100:02d}"
        prev = k - 1 if k > 0 else None
        s = sales[k]
        out["sales_growth"][y] = (div(s, sales[prev]) - 1 if prev is not None and sales[prev] and sales[prev] > 0
                                  and s is not None else None)
        out["ebitda_margin"][y] = div(ebitda[k], s) if s and s > 0 else None
        out["net_margin"][y] = div(net_profit[k], s) if s and s > 0 else None
        eq = avg(equity[prev], equity[k]) if prev is not None else None
        out["roe"][y] = div(net_profit[k], eq) if eq and eq > 0 else None
        cap = (avg(equity[prev] + (borrowings[prev] or 0), equity[k] + (borrowings[k] or 0))
               if prev is not None and equity[prev] is not None and equity[k] is not None else None)
        ebit = pbt[k] + (interest[k] or 0) if pbt[k] is not None else None
        out["roce"][y] = div(ebit, cap) if cap and cap > 0 else None
    return out


def history_from_tables(pl, bs, n: int = HISTORY_YEARS):
    """(history, year_end month) from a parsed Screener P&L and balance sheet; ({}, "") when unusable."""
    if pl is None or bs is None or "Sales" not in pl.rows:
        return {}, ""
    bs_idx = {h: i for i, h in enumerate(bs.headers)}
    cols = [(i, h) for i, h in enumerate(pl.headers) if h.upper() != "TTM" and _PERIOD.match(h) and h in bs_idx]
    if len(cols) < 2:
        return {}, ""
    labels = [h for _, h in cols]

    def pl_row(name):
        row = pl.rows.get(name, [])
        return [row[i] if i < len(row) else None for i, _ in cols]

    def bs_row(name):
        row = bs.rows.get(name, [])
        return [row[bs_idx[h]] if bs_idx[h] < len(row) else None for h in labels]

    eq_cap, res = bs_row("Equity Capital"), bs_row("Reserves")
    equity = [a + b if a is not None and b is not None else None for a, b in zip(eq_cap, res)]
    hist = ratio_history(labels, pl_row("Sales"), pl_row("Operating Profit"), pl_row("Net Profit"),
                         pl_row("Profit before tax"), pl_row("Interest"), equity, bs_row("Borrowings"), n)
    return hist, _PERIOD.match(labels[-1]).group(1).title()


def subject_history(fin, n: int = HISTORY_YEARS):
    """The subject company's history on the same definitions as its peers."""
    L = fin.lines
    return ratio_history(fin.period_labels, L["sales"], L["ebitda"], L["net_profit"], L["pbt"], L["interest"],
                         L["total_equity"], L["borrowings"], n)


def snapshot_from_page(parsed: dict, symbol: str, consolidated: bool, cash: Optional[float]) -> Peer:
    top = {k.lower(): v for k, v in parsed["top_ratios"].items()}
    p = Peer(symbol=symbol, name=parsed["name"], consolidated=consolidated)
    p.price = parse_number(top.get("current price"))
    p.market_cap = parse_number(top.get("market cap"))
    p.pe = parse_number(top.get("stock p/e"))
    bv = parse_number(top.get("book value"))
    p.pb = p.price / bv if p.price and bv and bv > 0 else None
    p.roe = parse_number(top.get("roe"))
    p.roce = parse_number(top.get("roce"))
    p.div_yield = parse_number(top.get("dividend yield"))
    pl = parsed["tables"].get("profit-loss")
    bs = parsed["tables"].get("balance-sheet")
    if pl is None or "Sales" not in pl.rows:
        p.notes.append("No Sales row on Screener P&L.")
        return p
    heads = pl.headers
    ttm = len(heads) > 0 and heads[-1].upper() == "TTM"
    col = len(heads) - 1
    annual = [i for i, h in enumerate(heads) if h.upper() != "TTM"]
    sales = pl.rows["Sales"]
    op = pl.rows.get("Operating Profit", [None] * len(heads))
    npf = pl.rows.get("Net Profit", [None] * len(heads))
    s_now, e_now, n_now = sales[col], op[col], npf[col]
    p.basis = "TTM" if ttm else heads[col]
    p.sales = s_now
    if len(annual) >= 4:
        p.sales_growth_3y = _cagr(sales[annual[-4]], sales[annual[-1]], 3)
    if s_now:
        p.ebitda_margin = e_now / s_now if e_now is not None else None
        p.net_margin = n_now / s_now if n_now is not None else None
    debt = _last(bs.rows.get("Borrowings", [])) if bs is not None else None
    inv = _last(bs.rows.get("Investments", [])) if bs is not None else None
    if cash is None:
        p.notes.append("Cash not reported; EV excludes cash.")
    if p.market_cap is not None:
        p.ev = p.market_cap + (debt or 0.0) - (cash or 0.0) - (inv or 0.0)
        if e_now and e_now > 0:
            p.ev_ebitda = p.ev / e_now
    p.history, p.year_end = history_from_tables(pl, bs)
    return p


def subject_snapshot(fin) -> Peer:
    """The subject company on the same definitions as its peers."""
    L = fin.lines
    t = fin.ttm or {}
    s = t.get("sales") or L["sales"][-1]
    e = t.get("ebitda") or L["ebitda"][-1]
    n = t.get("net_profit") or L["net_profit"][-1]
    p = Peer(symbol=fin.symbol, name=fin.name, consolidated=fin.consolidated, price=fin.price, market_cap=fin.market_cap,
             pe=fin.pe, roe=fin.roe, roce=fin.roce, div_yield=fin.dividend_yield, sales=s,
             basis="TTM" if t.get("sales") else fin.year_labels[-1])
    p.pb = fin.price / fin.book_value if fin.book_value else None
    p.ev = fin.market_cap + L["borrowings"][-1] - L["cash"][-1] - L["investments"][-1]
    p.ev_ebitda = p.ev / e if e and e > 0 else None
    p.ebitda_margin = e / s if s else None
    p.net_margin = n / s if s else None
    if fin.n >= 4:
        p.sales_growth_3y = _cagr(L["sales"][-4], L["sales"][-1], 3)
    p.history = subject_history(fin)
    p.year_end = fin.period_labels[-1][:3].title()
    return p


def backfill_history(ps: PeerSet, client: ScreenerClient) -> int:
    """Add multi-year history to peers cached before it existed, from their cached (or downloadable) pages.

    Returns the number of peers updated. Peers whose page is unavailable keep an empty history."""
    done = 0
    for p in ps.peers:
        if p.history:
            continue
        for cons in ([True, False] if p.consolidated else [False]):
            try:
                html = client.fetch_page(p.symbol, consolidated=cons)
            except EquityAgentError:
                html = None
            if not html:
                continue
            parsed = parse_company_page(html, p.symbol)
            hist, ye = history_from_tables(parsed["tables"].get("profit-loss"), parsed["tables"].get("balance-sheet"))
            if hist and any(hist["ebitda_margin"].values()):
                p.history, p.year_end = hist, ye
                done += 1
                break
    return done


def _cache_path(client: ScreenerClient, symbol: str) -> Path:
    return client.cache_dir / symbol.upper() / "peers" / "peers.json"


def fetch_peer_snapshot(client: ScreenerClient, sym: str, consolidated: bool) -> Optional[Peer]:
    html = client.fetch_page(sym, consolidated=consolidated)
    if html is None and consolidated:
        consolidated = False
        html = client.fetch_page(sym, consolidated=False)
    if html is None:
        return None
    parsed = parse_company_page(html, sym)
    if consolidated and not (parsed["tables"].get("profit-loss") and parsed["tables"]["profit-loss"].has_data("Sales")):
        html = client.fetch_page(sym, consolidated=False)
        if html is None:
            return None
        consolidated = False
        parsed = parse_company_page(html, sym)
    if is_bank_layout(parsed["tables"]):
        return None
    cash = None
    bs = parsed["tables"].get("balance-sheet")
    parent = bs.schedules.get("Other Assets") if bs is not None else None
    if parent:
        try:
            items = parse_schedule(client.fetch_schedule(sym, consolidated, parsed["company_id"], parent, "balance-sheet"))
            ce = items.get("Cash Equivalents")
            if ce is not None:
                last_head = bs.headers[-1]
                cash = (ce.amounts or {}).get(last_head) if ce.amounts else None
                if cash is None:
                    cash = ce.values.get(last_head)
        except EquityAgentError:
            cash = None
    return snapshot_from_page(parsed, sym, consolidated, cash)


def get_peers(client: ScreenerClient, raw, cfg: dict, symbols: Optional[List[str]] = None,
              reason: str = "") -> Optional[PeerSet]:
    """Peer snapshot for ``raw`` (a RawCompany). Returns None when nothing is available."""
    path = _cache_path(client, raw.symbol)
    wanted = [s.upper() for s in symbols] if symbols else None
    if client._fresh(path):
        ps = PeerSet.from_dict(json.loads(path.read_text(encoding="utf-8")))
        if wanted is None or [p.symbol for p in ps.peers] == wanted:
            if not ps.has_history and ps.peers:
                # snapshot cached before multi-year history existed: rebuild it from the peers' pages
                if backfill_history(ps, client) and not client.offline:
                    path.write_text(json.dumps(ps.to_dict(), indent=1, ensure_ascii=False), encoding="utf-8")
            return ps
    if client.offline:
        return None
    max_n = int(cfg.get("max_peers", 6))
    warnings: List[str] = []
    if wanted:
        rows = [{"symbol": s, "name": s, "consolidated": True} for s in wanted]
        source = f"Analyst-selected peers: {reason}" if reason else "Analyst-selected peers"
    else:
        if not raw.warehouse_id:
            return None
        resp = client._http_get(f"{BASE_URL}/api/company/{raw.warehouse_id}/peers/")
        if resp.status_code != 200:
            return None
        rows = [r for r in parse_peer_table(resp.text) if r["symbol"] != raw.symbol.upper()][:max_n]
        source = "Screener.in peer table (same industry, by market cap)"
    peers: List[Peer] = []
    for r in rows:
        try:
            p = fetch_peer_snapshot(client, r["symbol"], r["consolidated"])
        except EquityAgentError as exc:
            warnings.append(f"Peer {r['symbol']} skipped ({exc.__class__.__name__}).")
            continue
        if p is None:
            warnings.append(f"Peer {r['symbol']} skipped (no usable Screener page or a bank layout).")
            continue
        peers.append(p)
    ps = PeerSet(subject=raw.symbol, peers=peers, source=source, fetched_at=time.strftime("%Y-%m-%d %H:%M"),
                 warnings=warnings)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(ps.to_dict(), indent=1, ensure_ascii=False), encoding="utf-8")
    return ps
