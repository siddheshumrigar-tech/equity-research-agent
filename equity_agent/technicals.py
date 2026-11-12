"""Price-based indicators computed from yfinance daily closes (reproducible, no hand-drawn levels).

* 50- and 200-day simple moving averages and the date of the last 50/200 cross
* RSI(14), Wilder smoothing
* distance from the 52-week high and low
* support / resistance: swing pivots (a close that is the extreme of +/- 10 sessions)
  over the last year, clustered within 2%; nearest levels below / above the last close
* 1M / 3M / 12M absolute returns and returns relative to the Nifty 50
"""
from __future__ import annotations

from dataclasses import asdict, dataclass, field
from typing import Dict, List, Optional

SESSIONS = {"1M": 21, "3M": 63, "6M": 126, "12M": 252}


def sma(values: List[float], n: int) -> List[Optional[float]]:
    out: List[Optional[float]] = []
    s = 0.0
    for i, v in enumerate(values):
        s += v
        if i >= n:
            s -= values[i - n]
        out.append(s / n if i >= n - 1 else None)
    return out


def rsi(values: List[float], n: int = 14) -> List[Optional[float]]:
    out: List[Optional[float]] = [None] * len(values)
    if len(values) <= n:
        return out
    gains = [max(0.0, values[i] - values[i - 1]) for i in range(1, len(values))]
    losses = [max(0.0, values[i - 1] - values[i]) for i in range(1, len(values))]
    ag, al = sum(gains[:n]) / n, sum(losses[:n]) / n
    out[n] = 100.0 if al == 0 else 100 - 100 / (1 + ag / al)
    for i in range(n + 1, len(values)):
        ag = (ag * (n - 1) + gains[i - 1]) / n
        al = (al * (n - 1) + losses[i - 1]) / n
        out[i] = 100.0 if al == 0 else 100 - 100 / (1 + ag / al)
    return out


def pivots(values: List[float], window: int = 10):
    highs, lows = [], []
    for i in range(window, len(values) - window):
        seg = values[i - window: i + window + 1]
        if values[i] == max(seg):
            highs.append(values[i])
        if values[i] == min(seg):
            lows.append(values[i])
    return highs, lows


def cluster(levels: List[float], tol: float = 0.02) -> List[float]:
    out: List[List[float]] = []
    for v in sorted(levels):
        if out and abs(v / (sum(out[-1]) / len(out[-1])) - 1) <= tol:
            out[-1].append(v)
        else:
            out.append([v])
    return [sum(c) / len(c) for c in out]


def period_return(values: List[float], sessions: int) -> Optional[float]:
    if len(values) <= sessions:
        return None
    return values[-1] / values[-1 - sessions] - 1


@dataclass
class Technicals:
    last_date: str
    last_close: float
    sma50: Optional[float]
    sma200: Optional[float]
    last_cross: Optional[str]  # "golden" / "death"
    last_cross_date: Optional[str]
    rsi14: Optional[float]
    high_52w: float
    low_52w: float
    from_high: float
    from_low: float
    supports: List[float]
    resistances: List[float]
    returns: Dict[str, Optional[float]] = field(default_factory=dict)
    relative: Dict[str, Optional[float]] = field(default_factory=dict)
    trend: str = ""
    sma50_series: List[Optional[float]] = field(default_factory=list)
    sma200_series: List[Optional[float]] = field(default_factory=list)
    rsi_series: List[Optional[float]] = field(default_factory=list)

    def to_dict(self, series: bool = False) -> dict:
        d = asdict(self)
        if not series:
            for k in ("sma50_series", "sma200_series", "rsi_series"):
                d.pop(k)
        return d


def compute_technicals(dates: List[str], close: List[float], index: Optional[List[float]] = None) -> Optional[Technicals]:
    if len(close) < 60:
        return None
    s50, s200, r = sma(close, 50), sma(close, 200), rsi(close, 14)
    cross = cross_date = None
    for i in range(len(close) - 1, 0, -1):
        a0, b0, a1, b1 = s50[i - 1], s200[i - 1], s50[i], s200[i]
        if None in (a0, b0, a1, b1):
            break
        if (a0 - b0) * (a1 - b1) < 0 or (a0 == b0 and a1 != b1):
            cross, cross_date = ("golden" if a1 > b1 else "death"), dates[i]
            break
    yr = close[-252:]
    hi, lo, last = max(yr), min(yr), close[-1]
    ph, pl = pivots(yr)
    res = [x for x in cluster(ph) if x > last * 1.005]
    sup = [x for x in cluster(pl) if x < last * 0.995]
    rets = {k: period_return(close, n) for k, n in SESSIONS.items()}
    rel = {}
    if index is not None and len(index) == len(close):
        for k, n in SESSIONS.items():
            a, b = period_return(close, n), period_return(index, n)
            rel[k] = (a - b) if a is not None and b is not None else None
    trend = ""
    if s50[-1] is not None and s200[-1] is not None:
        if last > s50[-1] > s200[-1]:
            trend = "uptrend (price above 50-day above 200-day average)"
        elif last < s50[-1] < s200[-1]:
            trend = "downtrend (price below 50-day below 200-day average)"
        else:
            trend = "range-bound (moving averages not aligned)"
    return Technicals(last_date=dates[-1], last_close=last, sma50=s50[-1], sma200=s200[-1], last_cross=cross,
                      last_cross_date=cross_date, rsi14=r[-1], high_52w=hi, low_52w=lo, from_high=last / hi - 1,
                      from_low=last / lo - 1, supports=sorted(sup, reverse=True)[:3], resistances=sorted(res)[:3],
                      returns=rets, relative=rel, trend=trend, sma50_series=s50, sma200_series=s200, rsi_series=r)
