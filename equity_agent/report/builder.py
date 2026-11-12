"""Assemble the PDF: front page, then sections in one-column and two-column page templates."""
from __future__ import annotations

import shutil
import tempfile
from pathlib import Path

from reportlab.platypus import CondPageBreak, NextPageTemplate, PageBreak

from .charts import Charts
from .common import R
from .layout import Exhibits, ReportDoc
from .sections import appendix, financials, forecasts, front, peers, quarterly, research, statements, technical, valuation
from .sections.story import story_in_charts as story_charts
from .styles import Theme


def write_report(path, res) -> str:
    theme = Theme(res.palette)
    tmp = Path(tempfile.mkdtemp(prefix="eq_report_"))
    try:
        final = res.narrative is not None
        doc = ReportDoc(path, res, theme, final=final)
        r = R(res=res, theme=theme, charts=Charts(theme, tmp), ex=Exhibits(theme), tmp=tmp, front_h=doc.front_height)
        story = [NextPageTemplate("body")]
        story += front.front_page(r)
        story.append(PageBreak())
        num = 0

        def nxt():
            nonlocal num
            num += 1
            return num

        def maybe(fn, brk=None):
            """Narrative sections: added (and numbered) only when they have content."""
            nonlocal num
            items = fn(r, num + 1)
            if not items:
                return []
            num += 1
            return ([brk] if brk is not None else []) + items

        story += story_charts(r, nxt())
        story.append(PageBreak())
        narrative = []
        narrative += maybe(research.thesis)
        narrative += maybe(research.company_overview, CondPageBreak(300))
        narrative += maybe(research.industry, CondPageBreak(300))
        narrative += maybe(research.pestle_esg, CondPageBreak(300))
        if narrative:
            story += narrative + [PageBreak()]
        story += quarterly.quarterly_results(r, nxt())
        story += financials.performance(r, nxt())
        story += financials.costs_and_working_capital(r, nxt())
        story += financials.balance_sheet_cash_flow(r, nxt())
        story += financials.returns(r, nxt())
        story.append(CondPageBreak(260))
        story += peers.peer_section(r, nxt())
        story.append(CondPageBreak(260))
        story += forecasts.forecasts_vs_history(r, nxt())
        story += appendix.assumptions(r, nxt())
        story.append(PageBreak())
        story += valuation.valuation(r, nxt())
        story.append(PageBreak())
        story += technical.technical_view(r, nxt())
        story.append(PageBreak())
        later = maybe(research.risks_catalysts) + maybe(research.corporate_actions, CondPageBreak(200))
        if later:
            story += later + [PageBreak()]
        story += statements.financial_statements(r, nxt())
        story.append(PageBreak())
        story += appendix.data_notes(r, nxt())
        story += appendix.disclosures(r)
        doc.build(story)
        res.report_exhibits = list(r.ex.titles)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)
    return str(path)
