"""Page templates, running header / footer, draft watermark and numbered exhibits."""
from __future__ import annotations

import io
from datetime import date
from typing import List, Optional

from reportlab.lib import colors
from reportlab.lib.utils import ImageReader
from reportlab.platypus import (BaseDocTemplate, Flowable, Frame, KeepTogether, NextPageTemplate, PageTemplate,
                                Paragraph, Spacer)

from .styles import CONTENT_W, FONT, FONT_B, GUTTER, HALF_W, INK, MARGIN, MUTED, PAGE_H, PAGE_W, RS, c

FRONT_BAND_TOP = 14  # gap between the page edge and the top of the front-page band
FRONT_BAND_H = 46
FRONT_BAND_GAP = 12  # space between the bottom of the band and the first line of content
BODY_TOP = 44
FOOT_H = 34


class ReportDoc(BaseDocTemplate):
    """A4 document with three page templates: ``front`` (header band), ``body`` (one column), ``two_col``."""

    def __init__(self, path, res, theme, final: bool, **kw):
        super().__init__(str(path), pagesize=(PAGE_W, PAGE_H), leftMargin=MARGIN, rightMargin=MARGIN,
                         topMargin=BODY_TOP, bottomMargin=FOOT_H, title=f"{res.fin.name} - Equity Research",
                         author="equity-research-agent", **kw)
        self.res, self.theme, self.final = res, theme, final
        h_body = PAGE_H - BODY_TOP - FOOT_H
        front_top = FRONT_BAND_TOP + FRONT_BAND_H + FRONT_BAND_GAP
        front = Frame(MARGIN, FOOT_H, CONTENT_W, PAGE_H - front_top - FOOT_H, id="front", leftPadding=0,
                      rightPadding=0, topPadding=0, bottomPadding=0)
        body = Frame(MARGIN, FOOT_H, CONTENT_W, h_body, id="body", leftPadding=0, rightPadding=0, topPadding=0,
                     bottomPadding=0)
        left = Frame(MARGIN, FOOT_H, HALF_W, h_body, id="left", leftPadding=0, rightPadding=0, topPadding=0,
                     bottomPadding=0)
        right = Frame(MARGIN + HALF_W + GUTTER, FOOT_H, HALF_W, h_body, id="right", leftPadding=0, rightPadding=0,
                      topPadding=0, bottomPadding=0)
        self.addPageTemplates([
            PageTemplate(id="front", frames=[front], onPage=self._front_page),
            PageTemplate(id="body", frames=[body], onPage=self._body_page),
            PageTemplate(id="two_col", frames=[left, right], onPage=self._body_page),
        ])
        self.front_height = PAGE_H - front_top - FOOT_H

    # ------------------------------------------------------------------ decorations
    def _footer(self, cv):
        t = self.theme
        cv.setStrokeColor(c("D1D5DB"))
        cv.setLineWidth(0.4)
        cv.line(MARGIN, FOOT_H - 8, PAGE_W - MARGIN, FOOT_H - 8)
        cv.setFont(FONT, 6.2)
        cv.setFillColor(c(MUTED))
        src = "Sources: Screener.in (reported financials), yfinance (prices), model estimates."
        if self.final and self.res.narrative is not None:
            src += " Narrative and analyst inputs by the host agent, sourced on the Sources page."
        cv.drawString(MARGIN, FOOT_H - 17, src)
        cv.drawString(MARGIN, FOOT_H - 25, "Model-generated research for education; not investment advice and not a "
                                           "SEBI-registered research report. See Disclosures.")
        cv.setFont(FONT_B, 7)
        cv.setFillColor(c(INK))
        cv.drawRightString(PAGE_W - MARGIN, FOOT_H - 17, f"{cv.getPageNumber()}")

    def _watermark(self, cv):
        if self.final:
            return
        cv.saveState()
        cv.setFillColor(colors.Color(0.85, 0.15, 0.15, alpha=0.10))
        cv.setFont(FONT_B, 42)
        cv.translate(PAGE_W / 2, PAGE_H / 2)
        cv.rotate(35)
        cv.drawCentredString(0, 0, "DRAFT - narrative pending")
        cv.restoreState()

    def _front_page(self, cv, doc):
        t, f, v = self.theme, self.res.fin, self.res.valuation
        cv.saveState()
        top = PAGE_H - FRONT_BAND_TOP
        cv.setFillColor(c(t.band))
        cv.rect(0, top - FRONT_BAND_H, PAGE_W, FRONT_BAND_H, stroke=0, fill=1)
        x = MARGIN
        logo = self.res.palette.logo_png
        if logo:
            try:
                img = ImageReader(io.BytesIO(logo))
                iw, ih = img.getSize()
                h = FRONT_BAND_H - 14
                w = min(70, iw * h / ih)
                cv.setFillColor(colors.white)
                cv.roundRect(x, top - FRONT_BAND_H + 7, w + 8, h, 3, stroke=0, fill=1)
                cv.drawImage(img, x + 4, top - FRONT_BAND_H + 9, width=w, height=h - 4, preserveAspectRatio=True,
                             mask="auto")
                x += w + 16
            except Exception:
                pass
        cv.setFillColor(colors.white)
        cv.setFont(FONT_B, 10.5)
        cv.drawString(x, top - 19, "EQUITY RESEARCH  |  INDIA")
        cv.setFont(FONT, 7.6)
        cv.drawString(x, top - 32, f"{f.sector}  >  {f.industry}")
        stage = "Initiating coverage" if self.final else "Model draft (narrative pending)"
        cv.setFont(FONT_B, 8.5)
        cv.drawRightString(PAGE_W - MARGIN, top - 19, stage)
        cv.setFont(FONT, 7.6)
        cv.drawRightString(PAGE_W - MARGIN, top - 32, date.fromisoformat(v.valuation_date).strftime("%d %B %Y"))
        cv.restoreState()
        self._footer(cv)
        self._watermark(cv)

    def _body_page(self, cv, doc):
        t, f, v = self.theme, self.res.fin, self.res.valuation
        cv.saveState()
        y = PAGE_H - 24
        cv.setFillColor(c(t.accent_fill))
        cv.rect(MARGIN, y - 2, 3, 11, stroke=0, fill=1)
        cv.setFont(FONT_B, 8)
        cv.setFillColor(c(INK))
        cv.drawString(MARGIN + 7, y, f"{f.name} ({f.symbol})")
        cv.setFont(FONT, 7.4)
        cv.setFillColor(c(MUTED))
        cv.drawRightString(PAGE_W - MARGIN, y, f"{v.rating}  |  CMP {RS}{f.price:,.0f}  |  12m target {RS}"
                                               f"{v.target_price:,.0f}  |  {date.fromisoformat(v.valuation_date):%d %b %Y}")
        cv.setStrokeColor(c(t.accent))
        cv.setLineWidth(0.7)
        cv.line(MARGIN, y - 7, PAGE_W - MARGIN, y - 7)
        cv.restoreState()
        self._footer(cv)
        self._watermark(cv)


class Exhibits:
    """Numbers exhibits in order of appearance."""

    def __init__(self, theme):
        self.theme = theme
        self.n = 0
        self.titles: List[str] = []

    def make(self, title: str, content, source: str = "Screener.in, model estimates", keep: bool = True,
             width: Optional[float] = None):
        self.n += 1
        self.titles.append(title)
        st = self.theme.st
        head = Paragraph(f"Exhibit {self.n}: {title}", st["exhibit"])
        head.keepWithNext = True  # never leave an exhibit title alone at the foot of a page
        parts = [head]
        parts += content if isinstance(content, list) else [content]
        parts.append(Paragraph(f"Source: {source}", st["source"]))
        parts.append(Spacer(1, 5))
        return KeepTogether(parts) if keep else parts


class HRule(Flowable):
    def __init__(self, width: float, colour: str, thickness: float = 0.8, space: float = 3):
        super().__init__()
        self.width, self.colour, self.thickness, self.space = width, colour, thickness, space

    def wrap(self, aw, ah):
        self._w = min(self.width, aw)
        return self._w, self.thickness + 2 * self.space

    def draw(self):
        self.canv.setStrokeColor(c(self.colour))
        self.canv.setLineWidth(self.thickness)
        self.canv.line(0, self.space, getattr(self, "_w", self.width), self.space)


def section_title(theme, text: str, number: Optional[int] = None):
    """Heading and rule, kept on the same page as the flowable that follows them."""
    st = theme.st
    label = f"{number}. {text}" if number is not None else text
    head, rule = Paragraph(label, st["h1"]), HRule(CONTENT_W, theme.accent, 0.8, 2)
    head.keepWithNext = rule.keepWithNext = True
    return [head, rule]


__all__ = ["ReportDoc", "Exhibits", "HRule", "section_title", "NextPageTemplate"]
