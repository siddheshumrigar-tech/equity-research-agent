"""Price-based inputs from yfinance: beta, daily price history, fiscal-year-end prices.

Only used for beta, price charts, technicals, price performance and the historical
P/E band. Fundamentals come from Screener. Every failure is reported as a warning
and the dependent output is left out; no series is invented.

Beta: OLS regression of 2-year weekly returns against the Nifty 50, then Blume
adjusted (0.67 x raw + 0.33) when ``beta_blume`` is true, then clamped to
[beta_min, beta_max].
"""
from __future__ import annotations

import json
import math
import time
from dataclasses import dataclass, field, fields
from pathlib import Path
from typing import Dict, List, Optional


@dataclass
class MarketData:
    beta: Optional[float] = None
    beta_raw: Optional[float] = None
    beta_adjusted: Optional[float] = None
    beta_obs: int = 0
    beta_source: str = ""
    # daily closes (up to 2 years), aligned with the benchmark index
    price_dates: List[str] = field(default_factory=list)
    price_close: List[float] = field(default_factory=list)
    index_close: List[float] = field(default_factory=list)
    month_end_prices: Dict[str, float] = field(default_factory=dict)  # "YYYY-MM" -> close (split + dividend adjusted)
    month_end_prices_unadj: Dict[str, float] = field(default_factory=dict)  # split adjusted only (actual traded)
    fetched_at: str = ""
    warnings: List[str] = field(default_factory=list)

    def to_dict(self) -> dict:
        return self.__dict__.copy()

    @classmethod
    def from_dict(cls, data: dict) -> "MarketData":
        names = {f.name for f in fields(cls)}
        return cls(**{k: v for k, v in data.items() if k in names})

    def last_year(self, days: int = 252):
        """(dates, close, index) for the most recent ``days`` sessions."""
        return self.price_dates[-days:], self.price_close[-days:], self.index_close[-days:]


def compute_beta(stock: List[float], index: List[float]) -> Optional[float]:
    """OLS beta of simple returns of ``stock`` on ``index`` (aligned price lists)."""
    if len(stock) != len(index) or len(stock) < 3:
        return None
    rs = [stock[i] / stock[i - 1] - 1 for i in range(1, len(stock))]
    ri = [index[i] / index[i - 1] - 1 for i in range(1, len(index))]
    n = len(rs)
    ms, mi = sum(rs) / n, sum(ri) / n
    cov = sum((a - ms) * (b - mi) for a, b in zip(rs, ri)) / (n - 1)
    var = sum((b - mi) ** 2 for b in ri) / (n - 1)
    return cov / var if var > 0 else None


def blume(beta_raw: float) -> float:
    return 0.67 * beta_raw + 0.33


def _cache_file(cache_dir: Path, symbol: str) -> Path:
    return cache_dir / symbol.upper() / "market" / "yfinance.json"


def _series(yf, tick, period, interval, auto_adjust=True):
    s = yf.Ticker(tick).history(period=period, interval=interval, auto_adjust=auto_adjust)["Close"].dropna()
    s.index = s.index.tz_localize(None)
    return s


def download_market_data(symbol: str, cfg: dict) -> MarketData:
    md = MarketData(fetched_at=time.strftime("%Y-%m-%d %H:%M"))
    try:
        import yfinance as yf
    except ImportError:
        md.warnings.append("yfinance not installed; beta default used, price chart omitted.")
        return md
    tick = f"{symbol.upper()}.NS"
    bench = cfg.get("benchmark_index", "^NSEI")
    try:
        s = _series(yf, tick, "2y", "1wk").copy()
        i = _series(yf, bench, "2y", "1wk").copy()
        s.index, i.index = s.index.normalize(), i.index.normalize()
        j = s.to_frame("s").join(i.to_frame("i"), how="inner").dropna()
        if len(j) >= 50:
            md.beta_raw = compute_beta(list(j["s"]), list(j["i"]))
            md.beta_obs = len(j) - 1
            md.beta_source = f"2-yr weekly returns of {tick} vs {bench} ({md.beta_obs} weeks)"
        else:
            md.warnings.append(f"Only {len(j)} weekly prices for {tick}; beta default used.")
    except Exception as exc:  # network / ticker problems
        md.warnings.append(f"Beta download failed for {tick} ({exc.__class__.__name__}); beta default used.")
    try:
        s = _series(yf, tick, "2y", "1d").copy()
        i = _series(yf, bench, "2y", "1d").copy()
        s.index, i.index = s.index.normalize(), i.index.normalize()
        j = s.to_frame("s").join(i.to_frame("i"), how="inner").dropna()
        if len(j) >= 20:
            md.price_dates = [d.strftime("%Y-%m-%d") for d in j.index]
            md.price_close = [float(v) for v in j["s"]]
            md.index_close = [float(v) for v in j["i"]]
        else:
            md.warnings.append("Less than 20 days of price history; price chart omitted.")
    except Exception as exc:
        md.warnings.append(f"Price history download failed ({exc.__class__.__name__}); price chart omitted.")
    try:
        m = _series(yf, tick, "10y", "1mo")
        md.month_end_prices = {d.strftime("%Y-%m"): float(v) for d, v in m.items() if not math.isnan(v)}
        mu = _series(yf, tick, "10y", "1mo", auto_adjust=False)
        md.month_end_prices_unadj = {d.strftime("%Y-%m"): float(v) for d, v in mu.items() if not math.isnan(v)}
    except Exception as exc:
        md.warnings.append(f"Monthly price history failed ({exc.__class__.__name__}); historical P/E band omitted.")
    return md


def fetch_market_data(symbol: str, cfg: dict, cache_dir: str | Path = "cache", offline: bool = False,
                      ttl_seconds: Optional[int] = 24 * 3600, refresh: bool = False) -> MarketData:
    cache = _cache_file(Path(cache_dir), symbol)
    fresh = cache.exists() and (offline or ttl_seconds is None
                                or (not refresh and time.time() - cache.stat().st_mtime < ttl_seconds))
    if fresh and not offline:
        # caches written before unadjusted prices / 2-year daily history were stored are refreshed
        raw = json.loads(cache.read_text(encoding="utf-8"))
        if "month_end_prices_unadj" not in raw:
            fresh = False
    if fresh:
        md = MarketData.from_dict(json.loads(cache.read_text(encoding="utf-8")))
        if not md.fetched_at:
            md.fetched_at = time.strftime("%Y-%m-%d %H:%M", time.localtime(cache.stat().st_mtime))
        return _apply_beta_bounds(md, cfg)
    if offline:
        md = MarketData()
        md.warnings.append("Offline run: no cached yfinance data; beta default used, price chart and P/E band omitted.")
        return _apply_beta_bounds(md, cfg)
    md = download_market_data(symbol, cfg)
    cache.parent.mkdir(parents=True, exist_ok=True)
    raw = md.to_dict()
    raw["warnings"] = list(md.warnings)
    cache.write_text(json.dumps(raw, indent=1), encoding="utf-8")
    return _apply_beta_bounds(md, cfg)


def _apply_beta_bounds(md: MarketData, cfg: dict) -> MarketData:
    lo, hi, default = cfg.get("beta_min", 0.5), cfg.get("beta_max", 2.0), cfg.get("beta_default", 1.0)
    if md.beta_raw is None:
        md.beta = default
        md.beta_adjusted = None
        md.beta_source = md.beta_source or f"Default {default:.2f} (no price history available)"
        return md
    adj = blume(md.beta_raw) if cfg.get("beta_blume", True) else md.beta_raw
    md.beta_adjusted = adj
    md.beta = min(max(adj, lo), hi)
    if "Blume" not in md.beta_source and cfg.get("beta_blume", True):
        md.beta_source = f"{md.beta_source}; raw {md.beta_raw:.2f}, Blume adjusted 0.67 x raw + 0.33"
    if md.beta != adj:
        md.warnings.append(f"Adjusted beta {adj:.2f} clamped to {md.beta:.2f}.")
    return md
