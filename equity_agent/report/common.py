"""Shared context for report sections."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

from reportlab.platypus import Image, Paragraph, Table, TableStyle

from .charts import Charts
from .layout import Exhibits
from .styles import CONTENT_W, Theme, fmt_num, fmt_pct


@dataclass
class R:
    """Everything a section needs: the run result, theme, charts, exhibit counter, narrative helpers."""
    res: object
    theme: Theme
    charts: Charts
    ex: Exhibits
    tmp: Path
    front_h: float = 720.0
    auto_notes: List[str] = field(default_factory=list)

    @property
    def st(self):
        return self.theme.st

    @property
    def nar(self):
        return self.res.narrative

    def P(self, text: str, style: str = "body") -> Paragraph:
        return Paragraph(text, self.st[style])

    def N(self, text) -> str:
        """Narrative text with tokens filled and citation markers styled."""
        if self.nar is None or text is None:
            return text or ""
        import re

        s = self.nar.text(str(text))
        s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
        # a non-breaking space keeps the marker on the same line as the word it cites
        def _cite_repl(m):
            clean_digits = re.sub(r"\s+", "", m.group(1))
            return f"\u00a0<super><font size=5.5>[{clean_digits}]</font></super>"

        return re.sub(r"\s*\[(\d+(?:\s*,\s*\d+)*)\]", _cite_repl, s)

    def img(self, path: Optional[str], w: float, h: float):
        return Image(path, width=w, height=h) if path else self.P("Chart not available (missing data).", "small")


def pair(left, right, w_left: float = 0.5, gap: float = 12, width: float = CONTENT_W) -> Table:
    """Two flowables (or lists) side by side."""
    from reportlab.platypus import KeepTogether

    unwrap = lambda x: list(x._content) if isinstance(x, KeepTogether) else x
    left, right = unwrap(left), unwrap(right)
    wl = (width - gap) * w_left
    wr = width - gap - wl
    t = Table([[left, "", right]], colWidths=[wl, gap, wr])
    t.setStyle(TableStyle([("VALIGN", (0, 0), (-1, -1), "TOP"), ("LEFTPADDING", (0, 0), (-1, -1), 0),
                           ("RIGHTPADDING", (0, 0), (-1, -1), 0), ("TOPPADDING", (0, 0), (-1, -1), 0),
                           ("BOTTOMPADDING", (0, 0), (-1, -1), 0)]))
    return t


def growth(series: List[Optional[float]]) -> List[Optional[float]]:
    out: List[Optional[float]] = [None]
    for a, b in zip(series, series[1:]):
        out.append(b / a - 1 if a and a > 0 and b is not None else None)
    return out


def ratio(a, b):
    return a / b if a is not None and b not in (None, 0) else None


def row_cells(r: R, label: str, vals, fmt=fmt_num, bold=False, label_style="tdl"):
    st = "tdb" if bold else "td"
    return [r.P(f"<b>{label}</b>" if bold else label, label_style)] + [r.P(fmt(v) if v is not None else "n/a", st)
                                                                       for v in vals]


__all__ = ["R", "pair", "growth", "ratio", "row_cells", "fmt_num", "fmt_pct"]
