"""Technical view: price with 50/200-day averages, support / resistance, RSI(14), returns, computed commentary."""
from __future__ import annotations

from datetime import datetime
from typing import List, Optional

from reportlab.platypus import Paragraph

from ..common import R, pair
from ..layout import section_title
from ..styles import CONTENT_W, HALF_W, RS, data_table, fmt_num, fmt_pct

CHART_DAYS = 252
PERIODS = ("1M", "3M", "6M", "12M")


def _date(s: Optional[str]) -> str:
    if not s:
        return "n/a"
    try:
        return datetime.strptime(str(s)[:10], "%Y-%m-%d").strftime("%d %b %Y")
    except ValueError:
        return str(s)


def _vs(price: float, level: Optional[float]) -> Optional[float]:
    """Price relative to a level: +5% = price 5% above it."""
    return price / level - 1 if level else None


def rsi_zone(v: Optional[float]) -> str:
    if v is None:
        return "not available"
    if v >= 70:
        return "overbought (70 or above)"
    if v <= 30:
        return "oversold (30 or below)"
    return "neutral (between 30 and 70)"


def technical_commentary(t) -> List[str]:
    """Plain statements computed from the indicators; no forecast and no judgement beyond the definitions."""
    last = t.last_close
    out = []
    if t.trend:
        s = f"<b>Trend.</b> {t.trend[0].upper()}{t.trend[1:]}."
    else:
        s = "<b>Trend.</b> Fewer than 200 sessions of history, so the long-term average is not available."
    parts = []
    for lab, lvl in (("50-day", t.sma50), ("200-day", t.sma200)):
        d = _vs(last, lvl)
        if d is not None:
            parts.append(f"{abs(d):.1%} {'above' if d >= 0 else 'below'} the {lab} average ({RS}{lvl:,.0f})")
    if parts:
        s += f" The close of {RS}{last:,.0f} on {_date(t.last_date)} is " + " and ".join(parts) + "."
    out.append(s)
    if t.last_cross:
        side = "above" if t.last_cross == "golden" else "below"
        out.append(f"<b>Moving-average cross.</b> The 50-day average last crossed {side} the 200-day average on "
                   f"{_date(t.last_cross_date)} (a {'golden' if side == 'above' else 'death'} cross).")
    elif t.sma200 is not None:
        out.append("<b>Moving-average cross.</b> No cross of the 50-day and 200-day averages in the history held.")
    if t.rsi14 is not None:
        out.append(f"<b>Momentum.</b> RSI(14) is {t.rsi14:.0f}, {rsi_zone(t.rsi14)}.")
    lv = []
    if t.supports:
        lv.append(f"nearest support {RS}{t.supports[0]:,.0f} ({t.supports[0] / last - 1:+.1%})")
    else:
        lv.append(f"no swing low below the price in the last year (52-week low {RS}{t.low_52w:,.0f})")
    if t.resistances:
        lv.append(f"nearest resistance {RS}{t.resistances[0]:,.0f} ({t.resistances[0] / last - 1:+.1%})")
    else:
        lv.append(f"no swing high above the price in the last year (52-week high {RS}{t.high_52w:,.0f})")
    out.append("<b>Levels.</b> " + lv[0][0].upper() + lv[0][1:] + "; " + lv[1] + f". The stock is "
               f"{abs(t.from_high):.1%} below its 52-week high and {t.from_low:.1%} above its 52-week low.")
    a, rel = t.returns.get("12M"), t.relative.get("12M")
    if a is not None:
        s = f"<b>Relative performance.</b> 12-month return {a:+.1%}"
        if rel is not None:
            s += (f", {abs(rel) * 100:.1f} pp {'ahead of' if rel >= 0 else 'behind'} the Nifty 50 "
                  f"({a - rel:+.1%})")
        out.append(s + ".")
    return out


def _levels_table(r: R, t):
    last = t.last_close
    rows = [[r.P("Indicator", "thl"), r.P(f"Level ({RS})", "th"), r.P("Price vs level", "th")]]

    def add(lab, lvl, show_vs=True, bold=False):
        rows.append([r.P(f"<b>{lab}</b>" if bold else lab, "tdl"), r.P(fmt_num(lvl, 0) if lvl is not None else "n/a",
                                                                      "tdb" if bold else "td"),
                     r.P(fmt_pct(_vs(last, lvl), 1, True) if show_vs else "", "td")])

    add(f"Last close ({_date(t.last_date)})", last, show_vs=False, bold=True)
    for k in range(len(t.resistances) - 1, -1, -1):
        add(f"Resistance R{k + 1}", t.resistances[k])
    for k, lvl in enumerate(t.supports):
        add(f"Support S{k + 1}", lvl)
    add("50-day average", t.sma50)
    add("200-day average", t.sma200)
    add("52-week high", t.high_52w)
    add("52-week low", t.low_52w)
    rows.append([r.P("RSI(14)", "tdl"), r.P(f"{t.rsi14:.1f}" if t.rsi14 is not None else "n/a", "td"),
                 r.P(rsi_zone(t.rsi14).split(" (")[0], "td")])
    cross = (f"{'Golden' if t.last_cross == 'golden' else 'Death'}, {_date(t.last_cross_date)}"
             if t.last_cross else "None in history")
    rows.append([r.P("Last 50/200-day cross", "tdl"), r.P("", "td"), r.P(cross, "td")])
    return data_table(r.theme, rows, [HALF_W * 0.44, HALF_W * 0.22, HALF_W * 0.34], total_rows=())


def _returns_table(r: R, t):
    rows = [[r.P("%", "thl")] + [r.P(k, "th") for k in PERIODS]]
    absr = [t.returns.get(k) for k in PERIODS]
    rel = [t.relative.get(k) for k in PERIODS]
    idx = [a - b if a is not None and b is not None else None for a, b in zip(absr, rel)]
    pg = lambda x: fmt_pct(x, 1, True)
    rows.append([r.P(r.res.fin.symbol, "tdl")] + [r.P(pg(x), "td") for x in absr])
    rows.append([r.P("Nifty 50", "tdl")] + [r.P(pg(x), "td") for x in idx])
    rows.append([r.P("<b>Relative (pp)</b>", "tdl")] +
                [r.P(f"<b>{x * 100:+.1f}</b>" if x is not None else "n/a", "tdb") for x in rel])
    return data_table(r.theme, rows, [HALF_W * 0.32] + [HALF_W * 0.17] * 4, total_rows=(3,))


def technical_view(r: R, num: int) -> list:
    res = r.res
    t, md, f = res.technicals, res.market, res.fin
    out = section_title(r.theme, "Technical view", num)
    if t is None:
        out.append(r.P("Fewer than 60 sessions of price history from yfinance; technical indicators are left out.",
                       "small"))
        return out
    out.append(r.P("The rating and target rest on the 12-month valuation. This page shows where the price stands "
                   "on standard indicators, computed from daily closes, as context for timing; it is not an input "
                   "to the target.", "small"))
    for s in technical_commentary(t):
        out.append(Paragraph(s, r.st["bul"], bulletText="■"))
    n = min(CHART_DAYS, len(md.price_close))
    cut = lambda vals: list(vals[-n:])
    h = 330
    p = r.charts.technical("tech_price", cut(md.price_dates), cut(md.price_close), cut(t.sma50_series),
                           cut(t.sma200_series), cut(t.rsi_series), t.supports, t.resistances, CONTENT_W, h,
                           label=f"{f.symbol} close")
    out.append(r.ex.make(f"Price, 50/200-day averages, support / resistance and RSI(14), last {n} sessions",
                         r.img(p, CONTENT_W, h),
                         f"yfinance daily closes ({f.symbol}.NS) to {_date(t.last_date)}. Averages are simple; RSI "
                         "uses Wilder smoothing. Support / resistance = closes that are the low / high of the "
                         "surrounding 21 sessions over the last year, grouped within 2%."))
    left = r.ex.make("Technical levels", _levels_table(r, t),
                     "yfinance, computed. Resistances above and supports below the last close.")
    right = r.ex.make("Price performance vs the Nifty 50", _returns_table(r, t),
                      "yfinance; 1M / 3M / 6M / 12M = 21 / 63 / 126 / 252 sessions. Relative = stock minus index "
                      "return.")
    out.append(pair(left, right))
    return out
