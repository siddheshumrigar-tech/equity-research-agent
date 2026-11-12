"""PDF research report. Thin wrapper kept for callers of ``write_pdf(path, res)``.

The report itself lives in ``equity_agent/report/``: an institutional layout (front page
with rating box and data sidebar, numbered exhibits with sources, house style in the
company's brand colour), built from the same data, forecast and valuation as the Excel
model. Qualitative sections come only from the host agent's ``narrative.json``; without
it the report is a labelled draft.
"""
from __future__ import annotations

from .report import write_report


def write_pdf(path, res) -> str:
    return write_report(path, res)
