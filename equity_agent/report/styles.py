"""House style: fonts, colours (brand colour with a WCAG contrast check), paragraph and table styles."""
from __future__ import annotations

import colorsys
from pathlib import Path
from typing import Dict, Iterable, List, Optional, Sequence

import matplotlib
from reportlab.lib import colors
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Table, TableStyle

PAGE_W, PAGE_H = A4
MARGIN = 34
CONTENT_W = PAGE_W - 2 * MARGIN
GUTTER = 14
HALF_W = (CONTENT_W - GUTTER) / 2

INK = "1F2937"
MUTED = "6B7280"
RULE = "D1D5DB"
SHADE = "F3F4F6"
RATING_COLOURS = {"BUY": "15803D", "ADD": "3F6212", "REDUCE": "B45309", "SELL": "B91C1C"}


# --------------------------------------------------------------------------- #
# fonts (DejaVu ships with matplotlib and has the ₹ glyph)
# --------------------------------------------------------------------------- #
def _register_fonts():
    if "DejaVu" in pdfmetrics.getRegisteredFontNames():
        return "DejaVu", "DejaVu-Bold", "DejaVu-Oblique"
    ttf = Path(matplotlib.get_data_path()) / "fonts" / "ttf"
    try:
        pdfmetrics.registerFont(TTFont("DejaVu", str(ttf / "DejaVuSans.ttf")))
        pdfmetrics.registerFont(TTFont("DejaVu-Bold", str(ttf / "DejaVuSans-Bold.ttf")))
        pdfmetrics.registerFont(TTFont("DejaVu-Oblique", str(ttf / "DejaVuSans-Oblique.ttf")))
        pdfmetrics.registerFont(TTFont("DejaVu-BoldOblique", str(ttf / "DejaVuSans-BoldOblique.ttf")))
        from reportlab.pdfbase.pdfmetrics import registerFontFamily
        registerFontFamily("DejaVu", normal="DejaVu", bold="DejaVu-Bold", italic="DejaVu-Oblique",
                           boldItalic="DejaVu-BoldOblique")
        return "DejaVu", "DejaVu-Bold", "DejaVu-Oblique"
    except Exception:
        return "Helvetica", "Helvetica-Bold", "Helvetica-Oblique"


FONT, FONT_B, FONT_I = _register_fonts()
RS = "₹" if FONT == "DejaVu" else "Rs."


# --------------------------------------------------------------------------- #
# colour
# --------------------------------------------------------------------------- #
def _rgb(hex_: str):
    h = hex_.lstrip("#")
    return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))


def _hex(rgb) -> str:
    return "".join(f"{max(0, min(255, int(round(c * 255)))):02X}" for c in rgb)


def luminance(hex_: str) -> float:
    def ch(c):
        return c / 12.92 if c <= 0.03928 else ((c + 0.055) / 1.055) ** 2.4
    r, g, b = (ch(c) for c in _rgb(hex_))
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast(a: str, b: str = "FFFFFF") -> float:
    la, lb = sorted((luminance(a), luminance(b)), reverse=True)
    return (la + 0.05) / (lb + 0.05)


def readable_on_white(hex_: str, target: float = 4.5) -> str:
    """Darken ``hex_`` (keeping hue) until it reaches ``target`` contrast on white."""
    hex_ = hex_.lstrip("#").upper()
    if contrast(hex_) >= target:
        return hex_
    h, l, s = colorsys.rgb_to_hls(*_rgb(hex_))
    while l > 0 and contrast(_hex(colorsys.hls_to_rgb(h, l, s))) < target:
        l -= 0.01
    return _hex(colorsys.hls_to_rgb(h, max(l, 0), s))


def tint(hex_: str, share: float) -> str:
    """Mix ``hex_`` with white; ``share`` = weight of white (0..1)."""
    r, g, b = _rgb(hex_)
    return _hex((r + (1 - r) * share, g + (1 - g) * share, b + (1 - b) * share))


def c(hex_: str):
    return colors.HexColor("#" + hex_.lstrip("#"))


class Theme:
    """Colours and paragraph styles for one report (brand colour applied)."""

    def __init__(self, palette):
        self.brand = palette.primary.upper()
        self.accent = readable_on_white(self.brand)          # text, rules, headings
        self.accent_fill = self.brand                         # bars and boxes
        self.on_fill = "FFFFFF" if contrast(self.brand) >= 3.0 else INK
        self.band = readable_on_white(palette.dark, 7.0)      # header band (white text on it)
        self.light = tint(self.brand, 0.88)                   # estimate columns, sidebar
        self.lighter = tint(self.brand, 0.94)
        self.chart = [self.brand, readable_on_white(self.brand, 7.0), "9CA3AF", tint(self.brand, 0.5), "4B5563",
                      tint(self.brand, 0.7)]
        self.st = self._styles()

    def _styles(self) -> Dict[str, ParagraphStyle]:
        A, I, M = c(self.accent), c(INK), c(MUTED)
        base = dict(fontName=FONT, textColor=I)
        return {
            "headline": ParagraphStyle("headline", fontName=FONT_B, fontSize=15, leading=18.5, textColor=A, spaceAfter=5),
            "h1": ParagraphStyle("h1", fontName=FONT_B, fontSize=12, leading=15, textColor=A, spaceBefore=6, spaceAfter=2),
            "h2": ParagraphStyle("h2", fontName=FONT_B, fontSize=9, leading=11.5, textColor=I, spaceBefore=5, spaceAfter=2),
            "h3": ParagraphStyle("h3", fontName=FONT_B, fontSize=8, leading=10, textColor=A, spaceBefore=3, spaceAfter=1),
            "body": ParagraphStyle("body", fontSize=8.2, leading=11.3, alignment=TA_JUSTIFY, spaceAfter=4, **base),
            "bodyl": ParagraphStyle("bodyl", fontSize=8.2, leading=11.3, alignment=TA_LEFT, spaceAfter=3, **base),
            "bul": ParagraphStyle("bul", fontSize=8.2, leading=11.2, leftIndent=9, bulletIndent=0, spaceAfter=3,
                                  alignment=TA_JUSTIFY, **base),
            "small": ParagraphStyle("small", fontName=FONT, fontSize=6.8, leading=8.6, textColor=M),
            "source": ParagraphStyle("source", fontName=FONT_I, fontSize=6.3, leading=8, textColor=M, spaceBefore=1),
            "exhibit": ParagraphStyle("exhibit", fontName=FONT_B, fontSize=7.4, leading=9.2, textColor=A, spaceAfter=2),
            "th": ParagraphStyle("th", fontName=FONT_B, fontSize=6.8, leading=8.4, textColor=I, alignment=TA_RIGHT),
            "thl": ParagraphStyle("thl", fontName=FONT_B, fontSize=6.8, leading=8.4, textColor=I),
            "td": ParagraphStyle("td", fontName=FONT, fontSize=6.8, leading=8.4, textColor=I, alignment=TA_RIGHT),
            "tdl": ParagraphStyle("tdl", fontName=FONT, fontSize=6.8, leading=8.4, textColor=I),
            "tdb": ParagraphStyle("tdb", fontName=FONT_B, fontSize=6.8, leading=8.4, textColor=I, alignment=TA_RIGHT),
            "tdlb": ParagraphStyle("tdlb", fontName=FONT_B, fontSize=6.8, leading=8.4, textColor=I),
            "side_h": ParagraphStyle("side_h", fontName=FONT_B, fontSize=7.4, leading=9.4, textColor=A, spaceBefore=4,
                                     spaceAfter=1),
            "side": ParagraphStyle("side", fontName=FONT, fontSize=6.9, leading=8.6, textColor=I),
            "side_r": ParagraphStyle("side_r", fontName=FONT, fontSize=6.9, leading=8.6, textColor=I, alignment=TA_RIGHT),
            "side_rb": ParagraphStyle("side_rb", fontName=FONT_B, fontSize=6.9, leading=8.6, textColor=I, alignment=TA_RIGHT),
            "rating": ParagraphStyle("rating", fontName=FONT_B, fontSize=19, leading=22, textColor=I),
            "label": ParagraphStyle("label", fontName=FONT_B, fontSize=6.4, leading=8, textColor=M),
        }


# --------------------------------------------------------------------------- #
# tables: horizontal rules only, bold totals, shaded estimate columns
# --------------------------------------------------------------------------- #
def data_table(theme: Theme, rows: List[list], widths: Sequence[float], header_rows: int = 1,
               estimate_cols: Iterable[int] = (), bold_rows: Iterable[int] = (), total_rows: Iterable[int] = (),
               section_rows: Iterable[int] = (), zebra: bool = False) -> Table:
    t = Table(rows, colWidths=list(widths), repeatRows=header_rows)
    A = c(theme.accent)
    style = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 1.6), ("BOTTOMPADDING", (0, 0), (-1, -1), 1.6),
        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3),
        ("LINEBELOW", (0, header_rows - 1), (-1, header_rows - 1), 0.9, A),
        ("LINEBELOW", (0, header_rows), (-1, -2), 0.25, c(RULE)),
        ("LINEBELOW", (0, -1), (-1, -1), 0.6, c(MUTED)),
    ]
    for col in estimate_cols:
        style.append(("BACKGROUND", (col, 0), (col, -1), c(theme.lighter)))
    for r in total_rows:
        style.append(("LINEABOVE", (0, r), (-1, r), 0.6, c(INK)))
    for r in section_rows:
        style.append(("BACKGROUND", (0, r), (-1, r), c(SHADE)))
    if zebra:
        style.append(("ROWBACKGROUNDS", (0, header_rows), (-1, -1), [colors.white, c(SHADE)]))
    t.setStyle(TableStyle(style))
    return t


# --------------------------------------------------------------------------- #
# number formatting
# --------------------------------------------------------------------------- #
def fmt_num(v: Optional[float], d: int = 0) -> str:
    if v is None:
        return "n/a"
    s = f"{abs(v):,.{d}f}"
    return f"({s})" if v < 0 else s


def fmt_pct(v: Optional[float], d: int = 1, signed: bool = False) -> str:
    if v is None:
        return "n/a"
    return f"{v * 100:+.{d}f}%" if signed else f"{v * 100:.{d}f}%"


def fmt_x(v: Optional[float], d: int = 1) -> str:
    return "n/a" if v is None else f"{v:.{d}f}x"


def fmt_rs(v: Optional[float], d: int = 0) -> str:
    return "n/a" if v is None else f"{RS}{v:,.{d}f}"
