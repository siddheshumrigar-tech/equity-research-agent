"""Charts in one house style (brand colour + greys). Every function returns a PNG path or None."""
from __future__ import annotations

from pathlib import Path
from typing import List, Optional, Sequence

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.ticker import FuncFormatter, MaxNLocator  # noqa: E402

from .styles import RS  # noqa: E402

DPI = 200
GRID = "#E5E7EB"
INK = "#1F2937"
MUTED = "#6B7280"


def _style():
    plt.rcParams.update({
        "font.family": "DejaVu Sans", "font.size": 7, "axes.titlesize": 7.5, "axes.labelsize": 6.8,
        "xtick.labelsize": 6.3, "ytick.labelsize": 6.3, "legend.fontsize": 6.2, "axes.edgecolor": "#9CA3AF",
        "axes.linewidth": 0.6, "xtick.color": MUTED, "ytick.color": MUTED, "axes.labelcolor": MUTED,
        "xtick.major.size": 2, "ytick.major.size": 0, "axes.spines.top": False, "axes.spines.right": False,
        "axes.spines.left": False, "legend.frameon": False,
    })


class Charts:
    def __init__(self, theme, out_dir: Path):
        _style()
        self.t = theme
        self.dir = Path(out_dir)
        self.P = "#" + theme.brand
        self.D = "#" + theme.chart[1]
        self.L = "#" + theme.chart[3]
        self.G = "#9CA3AF"

    def _fig(self, w_pt: float, h_pt: float, ncols: int = 1):
        fig, ax = plt.subplots(1, ncols, figsize=(w_pt / 72, h_pt / 72))
        return fig, ax

    def _ax(self, ax, ylab: Optional[str] = None):
        ax.grid(axis="y", color=GRID, lw=0.5)
        ax.set_axisbelow(True)
        if ylab:
            ax.set_ylabel(ylab)

    def _save(self, fig, name: str) -> str:
        p = self.dir / f"{name}.png"
        fig.tight_layout(pad=0.3)
        fig.savefig(p, dpi=DPI)
        plt.close(fig)
        return str(p)

    @staticmethod
    def _k(v, _=None):
        return f"{v / 1000:,.0f}k" if abs(v) >= 10000 else f"{v:,.0f}"

    # ------------------------------------------------------------------ charts
    def price_vs_index(self, name, dates, close, index, w, h, label="Stock", bench="Nifty 50") -> Optional[str]:
        if len(close) < 20:
            return None
        fig, ax = self._fig(w, h)
        s0, i0 = close[0], index[0]
        x = range(len(close))
        ax.plot(x, [p / s0 * 100 for p in close], color=self.P, lw=1.3, label=label)
        ax.plot(x, [p / i0 * 100 for p in index], color=self.G, lw=1.0, label=bench)
        ax.axhline(100, color=GRID, lw=0.6)
        step = max(1, len(dates) // 4)
        ax.set_xticks(list(x)[::step])
        ax.set_xticklabels([dates[k][2:7] for k in range(0, len(dates), step)])
        self._ax(ax)
        ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=2, handlelength=1.2, borderaxespad=0.2)
        return self._save(fig, name)

    def technical(self, name, dates, close, sma50, sma200, rsi, supports, resistances, w, h,
                  label="Close") -> Optional[str]:
        """Price with 50/200-day averages and support / resistance lines, RSI(14) in a panel below."""
        if len(close) < 20:
            return None
        from datetime import datetime

        fig, (ax, ax2) = plt.subplots(2, 1, figsize=(w / 72, h / 72), sharex=True,
                                      gridspec_kw={"height_ratios": [3, 1], "hspace": 0.08})
        nan = float("nan")
        x = list(range(len(close)))
        clean = lambda vals: [nan if v is None else v for v in vals]
        ax.plot(x, close, color=self.P, lw=1.2, label=label)
        ax.plot(x, clean(sma50), color=INK, lw=0.9, label="50-day average")
        ax.plot(x, clean(sma200), color=self.G, lw=0.9, ls="--", label="200-day average")
        marks = [(lvl, f"R{k + 1} {RS}{lvl:,.0f}", "#B91C1C") for k, lvl in enumerate(resistances)]
        marks += [(lvl, f"S{k + 1} {RS}{lvl:,.0f}", "#15803D") for k, lvl in enumerate(supports)]
        for lvl, _, col in marks:
            ax.axhline(lvl, color=col, ls=":", lw=0.7)
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
        # labels in the right margin, nudged apart so close levels stay legible
        lo, hi = ax.get_ylim()
        gap = (hi - lo) * 0.05
        placed: List[float] = []
        for lvl, text, col in sorted(marks):
            y = max(lvl, placed[-1] + gap) if placed else lvl
            placed.append(y)
            ax.annotate(text, xy=(len(x) - 1, lvl), xytext=(len(x) + len(x) * 0.012, y), textcoords="data",
                        color=col, fontsize=5.8, va="center", ha="left", annotation_clip=False)
        self._ax(ax)
        ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=3, handlelength=1.6, borderaxespad=0.2)
        ax2.plot(x, clean(rsi), color=INK, lw=0.8)
        ax2.axhline(70, color="#B91C1C", ls="--", lw=0.6)
        ax2.axhline(30, color="#15803D", ls="--", lw=0.6)
        ax2.set_ylim(0, 100)
        ax2.set_yticks([30, 50, 70])
        ax2.set_ylabel("RSI(14)")
        self._ax(ax2)
        # a tick at the first session of every other month
        ticks, labs, prev = [], [], None
        for i, d in enumerate(dates):
            try:
                dt = datetime.strptime(str(d)[:10], "%Y-%m-%d")
            except ValueError:
                continue
            if prev is not None and dt.month != prev and dt.month % 2 == 1:
                ticks.append(i)
                labs.append(dt.strftime("%b %y"))
            prev = dt.month
        ax2.set_xticks(ticks)
        ax2.set_xticklabels(labs)
        ax2.set_xlim(0, len(x) - 1)
        p = self.dir / f"{name}.png"
        fig.subplots_adjust(left=0.07, right=0.91, top=0.93, bottom=0.07)
        fig.savefig(p, dpi=DPI)
        plt.close(fig)
        return str(p)

    def bars_line(self, name, labels, bars, line, n_actual, w, h, bar_label, line_label, line_pct=True) -> str:
        fig, ax = self._fig(w, h)
        cols = [self.D] * n_actual + [self.L] * (len(labels) - n_actual)
        ax.bar(labels, bars, color=cols, width=0.65, label=bar_label)
        ax.yaxis.set_major_formatter(FuncFormatter(self._k))
        self._ax(ax)
        ax2 = ax.twinx()
        nan = float("nan")
        ax2.plot(labels, [nan if v is None else (v * 100 if line_pct else v) for v in line], color=INK, marker="o",
                 ms=2.2, lw=1.0, label=line_label)
        ax2.spines["right"].set_visible(False)
        ax2.tick_params(labelsize=6.2, colors=MUTED, length=0)
        if line_pct:
            ax2.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:.0f}%"))
        # headroom so the legend above the plot never sits on the tallest bar or the highest point
        top = max([b for b in bars if b is not None] or [0])
        if top > 0:
            ax.set_ylim(top=top * 1.15)
        lv = [v * 100 if line_pct else v for v in line if v is not None]
        if lv:
            lo, hi = min(lv), max(lv)
            span = (hi - lo) or abs(hi) or 1
            ax2.set_ylim(lo - 0.1 * span, hi + 0.2 * span)
        ax.tick_params(axis="x", rotation=90 if len(labels) > 10 else 0)
        h1, l1 = ax.get_legend_handles_labels()
        h2, l2 = ax2.get_legend_handles_labels()
        ax.legend(h1 + h2, l1 + l2, loc="lower left", bbox_to_anchor=(0, 1.0), ncol=2, handlelength=1.2,
                  borderaxespad=0.2)
        return self._save(fig, name)

    def lines(self, name, labels, series: Sequence[tuple], w, h, pct=True, split: Optional[int] = None,
              legend_cols=3, hlines: Sequence[tuple] = ()) -> str:
        """``hlines`` = [(value, label)] reference lines in the same units as the series."""
        fig, ax = self._fig(w, h)
        palette = [self.P, INK, self.G, self.L]
        nan = float("nan")
        for k, (lab, vals) in enumerate(series):
            ys = [nan if v is None else (v * 100 if pct else v) for v in vals]
            ax.plot(labels, ys, color=palette[k % len(palette)], marker="o", ms=2, lw=1.1, label=lab)
        if split is not None and 0 < split < len(labels):
            ax.axvline(split - 0.5, color=MUTED, ls=":", lw=0.7)
        for val, lab in hlines:
            y = val * 100 if pct else val
            ax.axhline(y, color="#B91C1C", ls="--", lw=0.8)
            ax.text(len(labels) - 1, y, f" {lab}", color="#B91C1C", fontsize=5.8, va="bottom", ha="right")
        if pct:
            ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:.0f}%"))
        self._ax(ax)
        ax.tick_params(axis="x", rotation=90 if len(labels) > 10 else 0)
        # legend above the plot so it never covers a line
        ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=legend_cols, handlelength=1.2, borderaxespad=0.2)
        return self._save(fig, name)

    def bars(self, name, labels, values, w, h, n_actual: Optional[int] = None, pct=False, colours=None) -> str:
        fig, ax = self._fig(w, h)
        if colours is None:
            na = len(labels) if n_actual is None else n_actual
            colours = [self.D] * na + [self.L] * (len(labels) - na)
        ax.bar(labels, [v * 100 if pct else v for v in values], color=colours, width=0.65)
        ax.axhline(0, color=MUTED, lw=0.5)
        ax.yaxis.set_major_formatter(FuncFormatter((lambda v, _: f"{v:.0f}%") if pct else self._k))
        self._ax(ax)
        ax.tick_params(axis="x", rotation=90 if len(labels) > 10 else 0)
        return self._save(fig, name)

    def grouped_bars(self, name, labels, series: Sequence[tuple], w, h, n_actual: Optional[int] = None) -> str:
        """Side-by-side bars per label; forecast columns (index >= n_actual) in lighter shades."""
        fig, ax = self._fig(w, h)
        k = max(1, len(series))
        width = 0.8 / k
        na = len(labels) if n_actual is None else n_actual
        dark, light = [self.D, self.G, INK], [self.L, "#D1D5DB", "#9CA3AF"]
        x = list(range(len(labels)))
        for j, (lab, vals) in enumerate(series):
            cols = [dark[j % 3] if i < na else light[j % 3] for i in x]
            ax.bar([i + (j - (k - 1) / 2) * width for i in x], [v if v is not None else 0 for v in vals],
                   width=width * 0.95, color=cols, label=lab)
        ax.set_xticks(x)
        ax.set_xticklabels(labels)
        ax.axhline(0, color=MUTED, lw=0.5)
        ax.yaxis.set_major_formatter(FuncFormatter(self._k))
        self._ax(ax)
        ax.tick_params(axis="x", rotation=90 if len(labels) > 10 else 0)
        ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), ncol=len(series), handlelength=1.2, borderaxespad=0.2)
        return self._save(fig, name)

    def pe_band(self, name, labels, values, mean: float, current: Optional[float], w, h) -> Optional[str]:
        """Fiscal-year-end P/E bars, the current P/E and the mean as a reference line."""
        if not values:
            return None
        labs, vals = list(labels), list(values)
        cols = [self.D] * len(vals)
        if current:
            labs.append("Now")
            vals.append(current)
            cols.append(self.P)
        fig, ax = self._fig(w, h)
        ax.bar(labs, vals, color=cols, width=0.6)
        for i, v in enumerate(vals):
            ax.text(i, v, f"{v:.0f}x", ha="center", va="bottom", fontsize=5.8, color=INK)
        ax.axhline(mean, color="#B91C1C", ls="--", lw=0.8, label=f"Mean of fiscal year-ends {mean:.0f}x")
        ax.legend(loc="lower left", bbox_to_anchor=(0, 1.0), handlelength=1.6, borderaxespad=0.2)
        ax.set_ylim(0, max(vals + [mean]) * 1.15)
        ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:.0f}x"))
        self._ax(ax)
        return self._save(fig, name)

    def peer_bars(self, name, names: Sequence[str], values: Sequence[Optional[float]], highlight: int,
                  median: Optional[float], w, h, unit: str = "pct") -> str:
        """Horizontal bars, one per company; the subject (index ``highlight``) in the brand colour.

        ``unit``: "pct" = fractions shown as %, "pctpt" = numbers already in %, "x" = multiples."""
        scale = 100 if unit == "pct" else 1
        fmt = (lambda v: f"{v:.0f}x") if unit == "x" else (lambda v: f"{v:.1f}%")
        rows = [(n, v * scale, i == highlight) for i, (n, v) in enumerate(zip(names, values)) if v is not None]
        fig, ax = self._fig(w, h)
        if not rows:
            ax.axis("off")
            return self._save(fig, name)
        ys = list(range(len(rows)))[::-1]
        ax.barh(ys, [v for _, v, _ in rows], color=[self.P if hi else "#D1D5DB" for _, _, hi in rows], height=0.62)
        lo = min(0.0, min(v for _, v, _ in rows))
        hi_v = max(0.0, max(v for _, v, _ in rows))
        span = (hi_v - lo) or 1
        for y, (_, v, hi) in zip(ys, rows):
            ax.text(v + (0.02 * span if v >= 0 else -0.02 * span), y, fmt(v), va="center",
                    ha="left" if v >= 0 else "right", fontsize=5.4, color=INK, fontweight="bold" if hi else "normal",
                    bbox=dict(facecolor="white", edgecolor="none", pad=0.4), zorder=3)
        if median is not None:
            ax.axvline(median * scale, color="#B91C1C", ls="--", lw=0.7)
        ax.set_yticks(ys)
        ax.set_yticklabels([n for n, _, _ in rows], fontsize=5.4)
        # room for the value labels at both ends
        ax.set_xlim(lo - (0.38 * span if lo < 0 else 0), hi_v + 0.34 * span)
        ax.xaxis.set_major_locator(MaxNLocator(nbins=3))
        ax.tick_params(axis="x", labelsize=5.2)
        ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: fmt(v)))
        ax.grid(axis="x", color=GRID, lw=0.4)
        ax.set_axisbelow(True)
        return self._save(fig, name)

    def peer_history(self, name, years: Sequence[str], panels: Sequence[tuple], w, h, subject: str = "Company") -> str:
        """One panel per ratio: each peer as a thin grey line, the peer median dashed, the company in brand colour.

        panels = [(title, company_values, median_values, [peer_values, ...])], values as fractions."""
        fig, axes = self._fig(w, h, ncols=len(panels))
        axes = axes if len(panels) > 1 else [axes]
        nan = float("nan")
        pc = lambda vals: [nan if v is None else v * 100 for v in vals]
        x = list(range(len(years)))
        for ax, (title, comp, med, others) in zip(axes, panels):
            for k, o in enumerate(others):
                ax.plot(x, pc(o), color="#D1D5DB", lw=0.7, label="Peers" if k == 0 else None)
            ax.plot(x, pc(med), color=INK, lw=1.0, ls="--", label="Peer median")
            ax.plot(x, pc(comp), color=self.P, lw=1.6, marker="o", ms=2.2, label=subject)
            # scale to the company and the median; outlier peers run off the panel instead of flattening it
            core = [v * 100 for v in list(comp) + list(med) if v is not None]
            if core:
                lo, hi = min(core), max(core)
                pad = max(hi - lo, 4.0) * 0.6
                ax.set_ylim(lo - pad, hi + pad)
            ax.set_title(title, loc="left", color=INK, fontweight="bold", fontsize=6.6)
            ax.set_xticks(x)
            ax.set_xticklabels(years, rotation=90, fontsize=5.4)
            ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:.0f}%"))
            ax.tick_params(axis="y", labelsize=5.4)
            self._ax(ax)
        h_, l_ = axes[0].get_legend_handles_labels()
        fig.legend(h_, l_, loc="upper center", ncol=3, handlelength=1.6, bbox_to_anchor=(0.5, 1.0))
        fig.tight_layout(pad=0.3, rect=(0, 0, 1, 0.9))
        p = self.dir / f"{name}.png"
        fig.savefig(p, dpi=DPI)
        plt.close(fig)
        return str(p)

    def stacked(self, name, labels, series: Sequence[tuple], w, h, pct=True) -> str:
        fig, ax = self._fig(w, h)
        bottom = [0.0] * len(labels)
        shades = [self.D, self.P, self.L, self.G, "#D1D5DB", "#4B5563", "#E5E7EB"]
        for k, (lab, vals) in enumerate(series):
            ys = [v * 100 if pct else v for v in vals]
            ax.bar(labels, ys, bottom=bottom, color=shades[k % len(shades)], width=0.65, label=lab)
            bottom = [b + y for b, y in zip(bottom, ys)]
        if pct:
            ax.yaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:.0f}%"))
        self._ax(ax)
        ax.tick_params(axis="x", rotation=90 if len(labels) > 10 else 0)
        ax.legend(loc="upper center", bbox_to_anchor=(0.5, -0.18), ncol=min(4, len(series)), handlelength=1.0)
        return self._save(fig, name)

    def small_multiples(self, name, labels, panels: Sequence[tuple], w, h) -> str:
        fig, axes = self._fig(w, h, ncols=len(panels))
        for ax, (title, vals, kind) in zip(axes, panels):
            ys = [None if v is None else (v * 100 if kind == "pct" else v) for v in vals]
            ax.plot(labels, ys, color=self.P, marker="o", ms=2, lw=1.1)
            ax.set_title(title, loc="left", color=INK, fontweight="bold")
            self._ax(ax)
            ax.tick_params(axis="x", rotation=90, labelsize=5.6)
        return self._save(fig, name)

    def football(self, name, rows: List[tuple], price: float, target: float, w, h) -> Optional[str]:
        """rows = [(label, low, high)]"""
        rows = [r for r in rows if r[1] is not None and r[2] is not None]
        if not rows:
            return None
        fig, ax = self._fig(w, h)
        for k, (lab, lo, hi) in enumerate(rows):
            lo, hi = min(lo, hi), max(lo, hi)
            ax.barh(k, max(hi - lo, (hi or 1) * 0.004), left=lo, color=self.P if k % 2 == 0 else self.L, height=0.5)
            ax.text(hi, k, f"  {RS}{lo:,.0f} - {RS}{hi:,.0f}", va="center", fontsize=6, color=INK)
        ax.set_yticks(range(len(rows)))
        ax.set_yticklabels([r[0] for r in rows])
        ax.axvline(price, color=INK, lw=1.0, ls="--")
        ax.axvline(target, color=self.D, lw=1.4)
        ax.text(price, len(rows) - 0.4, f"CMP {RS}{price:,.0f}", fontsize=6, ha="center", color=INK)
        ax.text(target, -0.9, f"Target {RS}{target:,.0f}", fontsize=6, ha="center", color=self.D)
        ax.set_ylim(-1.2, len(rows) - 0.1)
        ax.grid(axis="x", color=GRID, lw=0.5)
        ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:,.0f}"))
        lo_all = min(min(r[1], r[2]) for r in rows + [("", price, target)])
        hi_all = max(max(r[1], r[2]) for r in rows + [("", price, target)])
        ax.set_xlim(lo_all - 0.05 * abs(hi_all - lo_all), hi_all + 0.32 * abs(hi_all - lo_all))
        return self._save(fig, name)

    def risk_reward(self, name, bear, base, bull, price, w, h) -> str:
        fig, ax = self._fig(w, h)
        vals = [("Bear", bear, "#B91C1C"), ("Base", base, self.D), ("Bull", bull, "#15803D")]
        lo = min(v for _, v, _ in vals + [("", price, "")])
        hi = max(v for _, v, _ in vals + [("", price, "")])
        span = hi - lo or abs(hi) or 1
        ax.hlines(0, lo - 0.08 * span, hi + 0.08 * span, color="#D1D5DB", lw=3)
        for k, (lab, v, col) in enumerate(vals):
            ax.plot([v], [0], "o", color=col, ms=5)
            y = 1.25 if k == 1 else 0.45  # base label on a higher row so close values do not overlap
            ax.text(v, y, f"{lab} {RS}{v:,.0f}", ha="center", va="bottom", fontsize=5.8, color=col)
        ax.plot([price], [0], "|", color=INK, ms=12, mew=1.4)
        ax.text(price, -0.55, f"CMP {RS}{price:,.0f}", ha="center", va="top", fontsize=5.8, color=INK)
        ax.set_ylim(-1.6, 2.3)
        ax.set_xlim(lo - 0.15 * span, hi + 0.15 * span)
        ax.axis("off")
        return self._save(fig, name)

    def impact_bars(self, name, labels: Sequence[str], values: Sequence[float], w, h) -> Optional[str]:
        """Horizontal bars of a % impact (fractions); negative red, positive green, largest at the top."""
        rows = [(lab, v * 100) for lab, v in zip(labels, values) if v is not None]
        if not rows:
            return None
        rows.sort(key=lambda x: abs(x[1]))
        fig, ax = self._fig(w, h)
        ys = list(range(len(rows)))
        ax.barh(ys, [v for _, v in rows], color=["#B91C1C" if v < 0 else "#15803D" for _, v in rows], height=0.6)
        lo, hi = min(0.0, min(v for _, v in rows)), max(0.0, max(v for _, v in rows))
        span = (hi - lo) or 1
        for y, (_, v) in zip(ys, rows):
            ax.text(v + (0.02 * span if v >= 0 else -0.02 * span), y, f"{v:+.0f}%", va="center",
                    ha="left" if v >= 0 else "right", fontsize=5.8, color=INK)
        ax.axvline(0, color=MUTED, lw=0.6)
        ax.set_yticks(ys)
        ax.set_yticklabels([lab for lab, _ in rows], fontsize=5.8)
        ax.set_xlim(lo - 0.25 * span, hi + 0.25 * span)
        ax.xaxis.set_major_formatter(FuncFormatter(lambda v, _: f"{v:.0f}%"))
        ax.grid(axis="x", color=GRID, lw=0.4)
        ax.set_axisbelow(True)
        return self._save(fig, name)

    def timeline(self, name, items: Sequence[tuple], w, h) -> Optional[str]:
        """items = [(datetime, label, impact)] with impact positive / negative / neutral; labels alternate sides."""
        items = sorted(items, key=lambda x: x[0])
        if not items:
            return None
        import matplotlib.dates as mdates

        fig, ax = self._fig(w, h)
        cols = {"positive": "#15803D", "negative": "#B91C1C"}
        xs = [float(mdates.date2num(d)) for d, _, _ in items]
        ax.axhline(0, color="#9CA3AF", lw=1.4, zorder=1)
        for k, ((_, lab, imp), x) in enumerate(zip(items, xs)):
            col = cols.get(str(imp).lower(), self.D)
            y = 0.55 if k % 2 == 0 else -0.55
            ax.plot([x], [0], "o", color=col, ms=5, zorder=3)
            ax.plot([x, x], [0, y * 0.75], color=col, lw=0.7, zorder=2)
            ax.text(x, y, lab, ha="center", va="bottom" if y > 0 else "top", fontsize=5.6, color=INK)
        span = xs[-1] - xs[0]
        pad = max(span * 0.12, 30)
        ax.set_xlim(xs[0] - pad, xs[-1] + pad)
        ax.set_ylim(-1.4, 1.4)
        # one tick per month (or per quarter / year on long ranges) so labels never repeat
        if span + 2 * pad <= 400:
            ax.xaxis.set_major_locator(mdates.MonthLocator())
        elif span + 2 * pad <= 1200:
            ax.xaxis.set_major_locator(mdates.MonthLocator(bymonth=(1, 4, 7, 10)))
        else:
            ax.xaxis.set_major_locator(mdates.YearLocator())
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %y"))
        ax.set_yticks([])
        ax.spines["bottom"].set_visible(False)
        return self._save(fig, name)

    def donut(self, name, labels, values, w, h, title: str = "") -> Optional[str]:
        vals = [max(0.0, v or 0.0) for v in values]
        if sum(vals) <= 0:
            return None
        fig, ax = self._fig(w, h)
        cols = [self.D, self.L, self.G, "#4B5563", "#D1D5DB", self.P]
        ax.pie(vals, colors=cols[:len(vals)], startangle=90, counterclock=False,
               wedgeprops=dict(width=0.38, edgecolor="white"), autopct=lambda p: f"{p:.0f}%" if p >= 4 else "",
               pctdistance=0.8, textprops={"fontsize": 5.8, "color": "white"})
        ax.text(0, 0, title, ha="center", va="center", fontsize=6.5, fontweight="bold", color=INK)
        ax.legend(labels, loc="center left", bbox_to_anchor=(1.0, 0.5), fontsize=5.8, handlelength=0.9)
        ax.set_aspect("equal")
        return self._save(fig, name)
