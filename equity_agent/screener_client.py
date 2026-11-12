"""Fetch and parse Screener.in company data (public pages, no login).

Screener has no official API. This module reads:

* the public company page ``/company/{SYMBOL}/consolidated/`` (falls back to the
  standalone page ``/company/{SYMBOL}/`` when consolidated figures are empty), and
* the JSON "schedules" endpoint the page itself calls when a ``+`` row is expanded:
  ``/api/company/{id}/schedules/?parent=...&section=...``.

Both are undocumented, so every parse step fails loudly instead of guessing.
Responses are cached on disk (default 24 h) and requests are rate-limited to
about one per second. The search endpoint is disallowed by robots.txt and is
never used: callers pass the NSE/BSE symbol.
"""
from __future__ import annotations

import json
import re
import sys
import time
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

import requests
from bs4 import BeautifulSoup

from .errors import BankNotSupported, DataValidationError, NetworkError, SymbolNotFound

BASE_URL = "https://www.screener.in"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/124.0 Safari/537.36 equity-research-agent/5.0"
)
ANNUAL_SECTIONS = ("profit-loss", "balance-sheet", "cash-flow")
ALL_SECTIONS = ("quarters", "profit-loss", "balance-sheet", "cash-flow", "ratios", "shareholding")
_SCHEDULE_RE = re.compile(r"showSchedule\(\s*['\"](.+?)['\"]\s*,\s*['\"]([\w-]+)['\"]")
_NUM_RE = re.compile(r"-?\d+(?:\.\d+)?")


# --------------------------------------------------------------------------- #
# Number parsing
# --------------------------------------------------------------------------- #
def parse_number(text) -> Optional[float]:
    """Parse Screener number strings.

    Handles Indian digit grouping ("2,45,348"), signs ("-39", "+15%"),
    percentages ("17.57%" -> 17.57, *not* divided by 100), currency prefixes
    ("₹ 676") and blanks ("" -> None).
    """
    if text is None:
        return None
    if isinstance(text, (int, float)):
        return float(text)
    s = str(text).replace("\u20b9", "").replace("\xa0", " ").replace(",", "").strip()
    s = s.replace("Cr.", "").replace("Cr", "").replace("%", "").strip()
    if s in ("", "-", "--", "—"):
        return None
    m = _NUM_RE.search(s.replace("+", ""))
    if not m:
        return None
    return float(m.group(0))


def _clean_label(text: str) -> str:
    text = text.replace("\xa0", " ")
    text = re.sub(r"\s+", " ", text).strip()
    if text.endswith("+"):
        text = text[:-1].rstrip()
    if text.startswith("CFO/EBITDA"):
        text = "CFO/EBITDA"
    return text


# --------------------------------------------------------------------------- #
# Data containers
# --------------------------------------------------------------------------- #
@dataclass
class Table:
    """One Screener data table: ``rows[label][i]`` is the value under ``headers[i]``."""

    headers: List[str]
    rows: Dict[str, List[Optional[float]]]
    schedules: Dict[str, str] = field(default_factory=dict)  # row label -> schedule parent name

    def row(self, label: str) -> List[Optional[float]]:
        return self.rows[label]

    def has_data(self, label: str) -> bool:
        return label in self.rows and any(v is not None for v in self.rows[label])


@dataclass
class ScheduleItem:
    values: Dict[str, Optional[float]]
    amounts: Optional[Dict[str, Optional[float]]] = None  # ₹ Cr from tooltips when values are %
    is_percent: bool = False


@dataclass
class RawCompany:
    symbol: str
    consolidated: bool
    name: str
    company_id: str
    warehouse_id: Optional[str]
    website: Optional[str]
    bse_code: Optional[str]
    nse_symbol: Optional[str]
    logo_url: Optional[str]
    sector_path: List[str]
    top_ratios: Dict[str, str]
    tables: Dict[str, Table]
    schedules: Dict[str, Dict[str, Dict[str, ScheduleItem]]]  # section -> parent -> item -> data
    source_url: str
    fetched_at: str

    def schedule(self, section: str, parent: str) -> Dict[str, ScheduleItem]:
        return self.schedules.get(section, {}).get(parent, {})


# --------------------------------------------------------------------------- #
# HTML parsing
# --------------------------------------------------------------------------- #
def _parse_table(section) -> Optional[Table]:
    table = section.find("table", class_="data-table")
    if table is None or table.find("thead") is None or table.find("tbody") is None:
        return None
    headers = [re.sub(r"\s+", " ", th.get_text(" ", strip=True)) for th in table.find("thead").find_all("th")][1:]
    rows: Dict[str, List[Optional[float]]] = {}
    schedules: Dict[str, str] = {}
    for tr in table.find("tbody").find_all("tr", recursive=False):
        tds = tr.find_all("td", recursive=False)
        if not tds:
            continue
        label = _clean_label(tds[0].get_text(" ", strip=True))
        if not label:
            continue
        values = [parse_number(td.get_text(strip=True)) for td in tds[1:]]
        if len(values) < len(headers):
            values += [None] * (len(headers) - len(values))
        if label in rows:  # duplicate label: keep the first occurrence
            continue
        rows[label] = values[: len(headers)]
        btn = tds[0].find("button")
        if btn is not None:
            m = _SCHEDULE_RE.search(btn.get("onclick", "") or "")
            if m:
                schedules[label] = m.group(1)
    return Table(headers=headers, rows=rows, schedules=schedules)


def parse_company_page(html: str, symbol: str = "") -> dict:
    """Parse a Screener company page into a plain dict (no schedules)."""
    soup = BeautifulSoup(html, "html.parser")
    info = soup.find(attrs={"data-company-id": True})
    if info is None:
        raise DataValidationError(
            f"Screener page for {symbol or 'symbol'} has no company id; page layout may have changed."
        )
    company_id = info["data-company-id"]
    wh = soup.find(attrs={"data-warehouse-id": True})
    h1 = soup.find("h1")
    name = h1.get_text(" ", strip=True) if h1 else symbol

    website = bse_code = nse_symbol = None
    for a in soup.find_all("a", href=True):
        href = a["href"]
        txt = a.get_text(" ", strip=True)
        if "bseindia.com/stock-share-price" in href and bse_code is None:
            m = re.search(r"/(\d{6})/?$", href.rstrip("/") + "/")
            bse_code = m.group(1) if m else None
        elif "nseindia.com/get-quotes/equity" in href and nse_symbol is None:
            m = re.search(r"symbol=([^&]+)", href)
            nse_symbol = m.group(1) if m else None
        elif website is None and href.startswith("http") and "screener.in" not in href:
            parent_cls = " ".join(a.parent.get("class") or [])
            if "flex-row" in parent_cls and "bseindia" not in href and "nseindia" not in href:
                website = href
    logo_url = None
    for img in soup.find_all("img", src=True):
        if "company-logos" in img["src"]:
            logo_url = img["src"]
            break

    sector_path: List[str] = []
    peers = soup.find("section", id="peers")
    if peers is not None:
        for a in peers.find_all("a", href=True):
            if a["href"].startswith("/market/"):
                t = a.get_text(" ", strip=True)
                t = re.sub(r"\s+(Broad Sector|Sector|Broad Industry|Industry)$", "", t)
                sector_path.append(t)

    top: Dict[str, str] = {}
    top_ul = soup.find(id="top-ratios")
    if top_ul is not None:
        for li in top_ul.find_all("li"):
            n = li.find(class_="name")
            v = li.find(class_="value")
            if n and v:
                top[re.sub(r"\s+", " ", n.get_text(" ", strip=True))] = re.sub(r"\s+", " ", v.get_text(" ", strip=True))

    tables: Dict[str, Table] = {}
    for sec in ALL_SECTIONS:
        node = soup.find("section", id=sec)
        if node is not None:
            t = _parse_table(node)
            if t is not None:
                tables[sec] = t
    return {
        "name": name,
        "company_id": company_id,
        "warehouse_id": wh["data-warehouse-id"] if wh else None,
        "website": website,
        "bse_code": bse_code,
        "nse_symbol": nse_symbol,
        "logo_url": logo_url,
        "sector_path": sector_path,
        "top_ratios": top,
        "tables": tables,
    }


def parse_schedule(payload: dict) -> Dict[str, ScheduleItem]:
    """Parse the JSON returned by the schedules endpoint."""
    out: Dict[str, ScheduleItem] = {}
    if not isinstance(payload, dict):
        raise DataValidationError("Screener schedule response is not a JSON object.")
    for item, data in payload.items():
        if not isinstance(data, dict):
            continue
        is_pct = item.strip().endswith("%")
        name = item.strip()[:-1].strip() if is_pct else item.strip()
        values = {k: parse_number(v) for k, v in data.items() if k not in ("tooltips", "isExpandable", "setAttributes")}
        amounts = None
        tips = data.get("tooltips")
        if isinstance(tips, dict):
            amounts = {k: parse_number(v) for k, v in tips.items()}
        out[name] = ScheduleItem(values=values, amounts=amounts, is_percent=is_pct)
    return out


def is_bank_layout(tables: Dict[str, Table]) -> bool:
    pl = tables.get("profit-loss")
    bs = tables.get("balance-sheet")
    if pl is not None and ("Financing Profit" in pl.rows or ("Revenue" in pl.rows and "Sales" not in pl.rows)):
        return True
    if bs is not None and "Deposits" in bs.rows:
        return True
    return False


def _has_annual_data(tables: Dict[str, Table]) -> bool:
    pl = tables.get("profit-loss")
    return pl is not None and (pl.has_data("Sales") or pl.has_data("Revenue"))


# --------------------------------------------------------------------------- #
# HTTP client with cache
# --------------------------------------------------------------------------- #
def _slug(text: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_").lower()


class ScreenerClient:
    """HTTP client with a file cache.

    ``offline=True`` reads only from ``cache_dir`` (ignores TTL) and never touches
    the network; tests use it with ``tests/fixtures/screener``.
    """

    def __init__(
        self,
        cache_dir: str | Path = "cache",
        ttl_seconds: Optional[int] = 24 * 3600,
        offline: bool = False,
        refresh: bool = False,
        min_interval: float = 1.0,
        timeout: float = 30.0,
        retries: int = 3,
        session: Optional[requests.Session] = None,
    ):
        self.cache_dir = Path(cache_dir)
        self.ttl = ttl_seconds
        self.offline = offline
        self.refresh = refresh
        self.min_interval = min_interval
        self.timeout = timeout
        self.retries = retries
        self.session = session or requests.Session()
        self.session.headers.update({"User-Agent": USER_AGENT, "Accept-Language": "en-IN,en;q=0.9"})
        self._last_request = 0.0

    # -- cache helpers ------------------------------------------------------ #
    def _path(self, symbol: str, view: str, name: str) -> Path:
        return self.cache_dir / symbol.upper() / view / name

    def _fresh(self, path: Path) -> bool:
        if not path.exists():
            return False
        if self.offline:
            return True
        if self.refresh:
            return False
        if self.ttl is None:
            return True
        return (time.time() - path.stat().st_mtime) < self.ttl

    def _http_get(self, url: str, params: Optional[dict] = None) -> requests.Response:
        if self.offline:
            raise NetworkError(f"Offline mode: no cached response for {url} {params or ''}")
        last_exc: Optional[Exception] = None
        for attempt in range(1, self.retries + 1):
            wait = self.min_interval - (time.time() - self._last_request)
            if wait > 0:
                time.sleep(wait)
            try:
                resp = self.session.get(url, params=params, timeout=self.timeout)
                self._last_request = time.time()
            except requests.RequestException as exc:
                last_exc = exc
                self._last_request = time.time()
                time.sleep(min(2 ** attempt, 10))
                continue
            if resp.status_code == 429 or resp.status_code >= 500:
                last_exc = NetworkError(f"HTTP {resp.status_code} from {url}")
                time.sleep(min(2 ** attempt * 2, 30))
                continue
            return resp
        raise NetworkError(f"Could not reach Screener.in ({url}): {last_exc}")

    # -- fetchers ----------------------------------------------------------- #
    def fetch_page(self, symbol: str, consolidated: bool) -> Optional[str]:
        view = "consolidated" if consolidated else "standalone"
        path = self._path(symbol, view, "page.html")
        missing = self._path(symbol, view, "NOT_FOUND")
        if self._fresh(missing):
            return None
        if self._fresh(path):
            return path.read_text(encoding="utf-8")
        url = f"{BASE_URL}/company/{symbol.upper()}/" + ("consolidated/" if consolidated else "")
        resp = self._http_get(url)
        path.parent.mkdir(parents=True, exist_ok=True)
        if resp.status_code == 404:
            missing.write_text("404", encoding="utf-8")
            return None
        if resp.status_code != 200:
            raise NetworkError(f"HTTP {resp.status_code} for {url}")
        path.write_text(resp.text, encoding="utf-8")
        return resp.text

    def fetch_schedule(self, symbol: str, consolidated: bool, company_id: str, parent: str, section: str) -> dict:
        view = "consolidated" if consolidated else "standalone"
        path = self._path(symbol, view, f"schedule_{section}_{_slug(parent)}.json")
        if self._fresh(path):
            return json.loads(path.read_text(encoding="utf-8"))
        params = {"parent": parent, "section": section}
        if consolidated:
            params["consolidated"] = ""
        resp = self._http_get(f"{BASE_URL}/api/company/{company_id}/schedules/", params=params)
        if resp.status_code != 200:
            raise NetworkError(f"HTTP {resp.status_code} for schedule {parent}/{section}")
        try:
            payload = resp.json()
        except ValueError as exc:
            raise DataValidationError(f"Schedule {parent}/{section} did not return JSON: {exc}") from exc
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(json.dumps(payload, ensure_ascii=False, indent=1), encoding="utf-8")
        return payload

    def fetch_binary(self, symbol: str, name: str, url: str) -> Optional[bytes]:
        """Fetch a small binary (logo). Returns None on any failure."""
        path = self._path(symbol, "assets", name)
        if self._fresh(path):
            return path.read_bytes()
        if self.offline:
            return None
        try:
            resp = self._http_get(url)
        except NetworkError:
            return None
        if resp.status_code != 200 or not resp.content:
            return None
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(resp.content)
        return resp.content

    def fetch_text(self, symbol: str, name: str, url: str) -> Optional[str]:
        data = self.fetch_binary(symbol, name, url)
        if data is None:
            return None
        return data.decode("utf-8", errors="replace")

    # -- high level --------------------------------------------------------- #
    def fetch_company(self, symbol: str) -> RawCompany:
        symbol = symbol.strip().upper()
        for suffix in (".NS", ".BO"):
            if symbol.endswith(suffix):
                symbol = symbol[: -len(suffix)]
        if not re.fullmatch(r"[A-Z0-9&\-_.]{1,20}", symbol):
            raise SymbolNotFound(f"'{symbol}' is not a valid NSE/BSE symbol.")

        parsed = None
        consolidated = True
        html = self.fetch_page(symbol, consolidated=True)
        if html is not None:
            parsed = parse_company_page(html, symbol)
            if is_bank_layout(parsed["tables"]):
                raise BankNotSupported(f"{parsed['name']} ({symbol}) is a bank/financial company; banks are not yet supported.")
        if parsed is None or not _has_annual_data(parsed["tables"]):
            consolidated = False
            html = self.fetch_page(symbol, consolidated=False)
            if html is None:
                raise SymbolNotFound(f"Screener.in has no company page for '{symbol}'. Pass the exact NSE symbol (e.g. TATATECH).")
            parsed = parse_company_page(html, symbol)
        if is_bank_layout(parsed["tables"]):
            raise BankNotSupported(f"{parsed['name']} ({symbol}) is a bank/financial company; banks are not yet supported.")
        if not _has_annual_data(parsed["tables"]):
            raise DataValidationError(f"Screener.in shows no annual Profit & Loss data for {symbol}.")

        schedules: Dict[str, Dict[str, Dict[str, ScheduleItem]]] = {}
        for section in ANNUAL_SECTIONS:
            table = parsed["tables"].get(section)
            if table is None:
                continue
            for parent in table.schedules.values():
                payload = self.fetch_schedule(symbol, consolidated, parsed["company_id"], parent, section)
                schedules.setdefault(section, {})[parent] = parse_schedule(payload)

        view = "consolidated" if consolidated else "standalone"
        page_path = self._path(symbol, view, "page.html")
        fetched = time.strftime("%Y-%m-%d %H:%M", time.localtime(page_path.stat().st_mtime)) if page_path.exists() else ""
        return RawCompany(
            symbol=symbol,
            consolidated=consolidated,
            name=parsed["name"],
            company_id=parsed["company_id"],
            warehouse_id=parsed["warehouse_id"],
            website=parsed["website"],
            bse_code=parsed["bse_code"],
            nse_symbol=parsed["nse_symbol"],
            logo_url=parsed["logo_url"],
            sector_path=parsed["sector_path"],
            top_ratios=parsed["top_ratios"],
            tables=parsed["tables"],
            schedules=schedules,
            source_url=f"{BASE_URL}/company/{symbol}/" + ("consolidated/" if consolidated else ""),
            fetched_at=fetched,
        )


def _main(argv: List[str]) -> int:
    import argparse

    ap = argparse.ArgumentParser(description="Print parsed Screener.in tables for a symbol.")
    ap.add_argument("symbol")
    ap.add_argument("--cache-dir", default="cache")
    ap.add_argument("--offline", action="store_true")
    args = ap.parse_args(argv)
    raw = ScreenerClient(cache_dir=args.cache_dir, offline=args.offline).fetch_company(args.symbol)
    print(f"{raw.name} ({raw.symbol}) id={raw.company_id} consolidated={raw.consolidated}")
    print(f"website={raw.website} bse={raw.bse_code} sector={' > '.join(raw.sector_path)}")
    for k, v in raw.top_ratios.items():
        print(f"  {k}: {v}")
    for sec, t in raw.tables.items():
        print(f"\n[{sec}] {' | '.join(t.headers)}")
        for label, vals in t.rows.items():
            print(f"  {label:<32} " + " ".join("" if v is None else f"{v:,.0f}" for v in vals))
    for sec, parents in raw.schedules.items():
        for parent, items in parents.items():
            print(f"\n[{sec} / {parent}] " + ", ".join(items))
    return 0


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
