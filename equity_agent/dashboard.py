"""Tesla-style executive Dashboard + hidden Dashboard_Calc sheet.

Layout follows the reference ``Tesla Financial model`` dashboard: banner with logo
(B1:E5), "Financial Year" selector (B6:C7 / E6), seven KPI cards across F1:S7,
share-info box (T2:X7) and eight charts in two rows (rows 8-24 and 25-39).
Every figure is a formula on Dashboard_Calc, which looks values up by the year
in Dashboard!E6 with INDEX/MATCH. Colours come from the company's logo palette.
"""
from __future__ import annotations

import io
from typing import List

from openpyxl.chart import BarChart, LineChart, PieChart, Reference, Series
from openpyxl.chart.data_source import AxDataSource, NumDataSource, NumRef, StrRef
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.legend import Legend
from openpyxl.chart.marker import Marker
from openpyxl.chart.series import DataPoint, SeriesLabel
from openpyxl.chart.shapes import GraphicalProperties
from openpyxl.chart.text import RichText, Text
from openpyxl.chart.title import Title
from openpyxl.drawing.line import LineProperties
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor, TwoCellAnchor
from openpyxl.drawing.text import CharacterProperties, Paragraph, ParagraphProperties, RegularTextRun
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.units import pixels_to_EMU
from openpyxl.worksheet.datavalidation import DataValidation

from .excel_writer import (F_NUM, S_BS, S_CALC, S_CB, S_CF, S_DASH, S_DRV, S_IS, S_WC, Ctx, _fill, q)

RED = "C00000"
GREEN = "00B050"
CHART_BG = "F2F2F2"

# Tesla column widths; F..S hold the seven 2-column KPI cards.
COL_WIDTHS = {"A": 1.5, "B": 11.2, "C": 11.2, "D": 1.5, "E": 12.8, "T": 19, "U": 11.2, "V": 4.8, "W": 13,
              "X": 15, "Y": 11.2, "Z": 11.2}
CARD_COL_WIDTH = 10.5
CARDS = [  # (title, calc row key, number format, is_ratio)
    ("Revenue", "rev", '#,##0', False),
    ("COGS", "cogs", '#,##0', False),
    ("OPEX", "opex", '#,##0', False),
    ("Gross Profit", "gp", '#,##0', False),
    ("Net Profit", "np", '#,##0', False),
    ("ROA", "roa", '0.0%', True),
    ("ROE", "roe", '0.0%', True),
]
# (from col, from row, to col, to row) as 0-based marker positions
CHART_ANCHORS = {
    "revenue": (1, 7, 9, 24), "expenses": (9, 7, 14, 24), "pie": (14, 7, 19, 24), "margins": (19, 7, 26, 24),
    "waterfall": (1, 24, 9, 39), "trade": (9, 24, 14, 39), "liquidity": (14, 24, 19, 39), "fcf": (19, 24, 26, 39),
}


# --------------------------------------------------------------------------- #
# Dashboard_Calc
# --------------------------------------------------------------------------- #
class CalcLayout:
    def __init__(self, n_items: int):
        self.year = 3
        self.kpi = {"rev": 5, "cogs": 6, "opex": 7, "gp": 8, "np": 9, "roa": 10, "roe": 11}
        self.margins = 14  # 4 rows
        self.pie = 20  # 7 rows
        self.items = 29
        self.n_items = n_items
        self.wf = self.items + n_items + 2  # header row; 9 steps follow
        self.trade = self.wf + 12  # 4 rows
        self.liq = self.trade + 6  # 3 rows
        self.hist = self.liq + 5  # labels row, values row
        self.fcf = self.hist + 4  # labels, actual, forecast
        self.info = self.fcf + 5  # share info rows


def _lookup(ctx: Ctx, sheet: str, key: str, year_cell: str) -> str:
    r = ctx.r(sheet, key)
    last = ctx.last_col
    return f"INDEX({q(sheet)}!$B${r}:${last}${r},MATCH({year_cell},{q(sheet)}!$B$4:${last}$4,0))"


def build_calc(wb, ctx: Ctx) -> CalcLayout:
    ws = wb.create_sheet(S_CALC)
    fin = ctx.fin
    lay = CalcLayout(len(fin.expense_items))
    ws.column_dimensions["A"].width = 34
    for c in "BCDEFGHIJ":
        ws.column_dimensions[c].width = 12
    ws["A1"] = "DASHBOARD CALCULATIONS - driven by Dashboard!E6 (do not edit)"
    ws["A1"].font = Font(bold=True)
    ws.cell(lay.year, 1, "Selected year / prior year")
    ws.cell(lay.year, 2, f"={q(S_DASH)}!$E$6")
    ws.cell(lay.year, 3, f"=B{lay.year}-1")
    ws.cell(4, 1, "KPI")
    for j, h in enumerate(("Selected", "Prior", "Change"), 2):
        ws.cell(4, j, h).font = Font(bold=True)
    cur, pri = f"$B${lay.year}", f"$C${lay.year}"
    kpi_src = {"rev": (S_IS, "sales", "Revenue"), "cogs": (S_IS, "cogs", "COGS"), "opex": (S_IS, "opex", "OPEX"),
               "gp": (S_IS, "gross_profit", "Gross Profit"), "np": (S_IS, "net_profit", "Net Profit"),
               "roa": (S_BS, "roa", "ROA"), "roe": (S_BS, "roe", "ROE")}
    for key, r in lay.kpi.items():
        sheet, src, label = kpi_src[key]
        ws.cell(r, 1, label)
        ws.cell(r, 2, f"={_lookup(ctx, sheet, src, cur)}")
        ws.cell(r, 3, f'=IFERROR({_lookup(ctx, sheet, src, pri)},"n/a")')
        if key in ("roa", "roe"):
            ws.cell(r, 4, f'=IFERROR((B{r}-C{r})*100,"n/a")')
        else:
            ws.cell(r, 4, f'=IFERROR(B{r}/C{r}-1,"n/a")')

    ws.cell(lay.margins - 1, 1, "Margins (selected year)").font = Font(bold=True)
    for k, (label, key) in enumerate((("Gross margin", "gross_margin"), ("EBITDA margin", "ebitda_margin"),
                                      ("PBT margin", "pbt_margin"), ("Net margin", "net_margin"))):
        ws.cell(lay.margins + k, 1, label)
        ws.cell(lay.margins + k, 2, f"={_lookup(ctx, S_IS, key, cur)}").number_format = '0.0%'

    ws.cell(lay.pie - 1, 1, "Where revenue goes (selected year, negatives shown as 0)").font = Font(bold=True)
    for k, (label, key) in enumerate((("COGS", "cogs"), ("Operating expenses", "opex"), ("Depreciation", "depreciation"),
                                      ("Interest", "interest"), ("Tax", "tax"), ("Minority & other", "minority_other"),
                                      ("Net profit", "net_profit"))):
        ws.cell(lay.pie + k, 1, label)
        ws.cell(lay.pie + k, 2, f"=MAX(0,{_lookup(ctx, S_IS, key, cur)})")

    ws.cell(lay.items - 1, 1, "Expense breakup (₹ Cr, selected year)").font = Font(bold=True)
    for k, it in enumerate(fin.expense_items):
        ws.cell(lay.items + k, 1, it.name)
        ws.cell(lay.items + k, 2, f"={_lookup(ctx, S_CB, f'item_{k}', cur)}")

    # waterfall
    w0 = lay.wf
    for j, h in enumerate(("P&L waterfall", "Delta", "Start", "End", "Base", "Up", "Up <0", "Down", "Down <0", "Total"), 1):
        ws.cell(w0, j, h).font = Font(bold=True)
    steps = [("Revenue", "sales", "total"), ("COGS", "cogs", "-"), ("OPEX", "opex", "-"),
             ("Other income", "other_income", "+"), ("Depreciation", "depreciation", "-"), ("Interest", "interest", "-"),
             ("Tax", "tax", "-"), ("Other items", None, "plug"), ("Net profit", "net_profit", "total")]
    for k, (label, key, kind) in enumerate(steps):
        r = w0 + 1 + k
        ws.cell(r, 1, label)
        if kind == "total":
            ws.cell(r, 2, f"={_lookup(ctx, S_IS, key, cur)}")
            ws.cell(r, 3, 0)
            ws.cell(r, 4, f"=B{r}")
            for col in range(5, 10):
                ws.cell(r, col, 0)
            ws.cell(r, 10, f"=B{r}")
            continue
        if kind == "-":
            ws.cell(r, 2, f"=-{_lookup(ctx, S_IS, key, cur)}")
        elif kind == "+":
            ws.cell(r, 2, f"={_lookup(ctx, S_IS, key, cur)}")
        else:  # balancing step so the bars end exactly at reported net profit
            ws.cell(r, 2, f"={_lookup(ctx, S_IS, 'net_profit', cur)}-D{r - 1}")
        ws.cell(r, 3, f"=D{r - 1}")
        ws.cell(r, 4, f"=C{r}+B{r}")
        S, E = f"C{r}", f"D{r}"
        lo, hi = f"MIN({S},{E})", f"MAX({S},{E})"
        cross = f"AND({lo}<0,{hi}>0)"
        ws.cell(r, 5, f"=IF({lo}>=0,{lo},IF({hi}<=0,{hi},0))")
        ws.cell(r, 6, f"=IF({E}>{S},IF({lo}>=0,{E}-{S},IF({hi}<=0,{S}-{E},{hi})),0)")
        ws.cell(r, 7, f"=IF({E}>{S},IF({cross},{lo},0),0)")
        ws.cell(r, 8, f"=IF({E}<{S},IF({lo}>=0,{S}-{E},IF({hi}<=0,{E}-{S},{hi})),0)")
        ws.cell(r, 9, f"=IF({E}<{S},IF({cross},{lo},0),0)")
        ws.cell(r, 10, 0)

    ws.cell(lay.trade - 1, 1, "Trade cycle (days)").font = Font(bold=True)
    for k, (label, key) in enumerate((("DSO", "dso"), ("DIO", "dio"), ("DPO", "dpo"), ("Cash cycle", "ccc"))):
        ws.cell(lay.trade + k, 1, label)
        ws.cell(lay.trade + k, 2, f"=ROUND({_lookup(ctx, S_WC, key, cur)},0)")
    ws.cell(lay.liq - 1, 1, "Liquidity ratios (approx.)").font = Font(bold=True)
    for k, (label, key) in enumerate((("Current ratio", "current_ratio"), ("Quick ratio", "quick_ratio"),
                                      ("Cash ratio", "cash_ratio"))):
        ws.cell(lay.liq + k, 1, label)
        ws.cell(lay.liq + k, 2, f"=ROUND({_lookup(ctx, S_BS, key, cur)},2)")

    ws.cell(lay.hist - 1, 1, "Revenue history (all years)").font = Font(bold=True)
    ws.cell(lay.hist, 1, "Year")
    ws.cell(lay.hist + 1, 1, "Revenue")
    ws.cell(lay.fcf - 1, 1, "Free cash flow (actual = CFO - capex; forecast = FCFF)").font = Font(bold=True)
    ws.cell(lay.fcf, 1, "Year")
    ws.cell(lay.fcf + 1, 1, "Actual")
    ws.cell(lay.fcf + 2, 1, "Forecast")
    for i in range(ctx.n_all):
        col = 2 + i
        src = ctx.col(i)
        ws.cell(lay.hist, col, ctx.labels[i])
        ws.cell(lay.hist + 1, col, f"={q(S_IS)}!{src}{ctx.r(S_IS, 'sales')}")
        ws.cell(lay.fcf, col, ctx.labels[i])
        fcf_ref = f"={q(S_CF)}!{src}{ctx.r(S_CF, 'fcf')}"
        actual = i < ctx.n_act
        ws.cell(lay.fcf + 1, col, fcf_ref if actual else "=NA()")
        ws.cell(lay.fcf + 2, col, fcf_ref if (not actual or i == ctx.n_act - 1) else "=NA()")

    info = [("Sector", f"{fin.sector} / {fin.industry}"), ("Market cap (₹ Cr)", f"={ctx.drv['mcap']}"),
            ("Price (₹)", f"={ctx.drv['price']}"), ("Fair value (₹)", f"={ctx.refs['target']}"),
            ("Status", f"={ctx.refs['rating']}"), ("Upside", f"={ctx.refs['upside']}")]
    for k, (label, f) in enumerate(info):
        ws.cell(lay.info + k, 1, label)
        ws.cell(lay.info + k, 2, f)
    # number formats (also used by chart data labels)
    fmt_ranges = [(lay.kpi["rev"], lay.kpi["np"], "BC", "#,##0"), (lay.kpi["roa"], lay.kpi["roe"], "BC", "0.0%"),
                  (lay.kpi["rev"], lay.kpi["np"], "D", "0.0%"), (lay.kpi["roa"], lay.kpi["roe"], "D", "0.0"),
                  (lay.pie, lay.pie + 6, "B", "#,##0"), (lay.items, lay.items + max(lay.n_items, 1) - 1, "B", "#,##0"),
                  (lay.wf + 1, lay.wf + 9, "BCDEFGHIJ", "#,##0"), (lay.trade, lay.trade + 3, "B", "0"),
                  (lay.liq, lay.liq + 2, "B", "0.00"), (lay.info + 1, lay.info + 3, "B", "#,##0.00"),
                  (lay.info + 5, lay.info + 5, "B", "+0.0%;-0.0%")]
    for r1, r2, cols, fmt in fmt_ranges:
        for r in range(r1, r2 + 1):
            for col in cols:
                ws[f"{col}{r}"].number_format = fmt
    for r in (lay.hist + 1, lay.fcf + 1, lay.fcf + 2):
        for i in range(ctx.n_all):
            ws.cell(r, 2 + i).number_format = "#,##0"
    ws.sheet_state = "hidden"
    return lay


# --------------------------------------------------------------------------- #
# chart helpers
# --------------------------------------------------------------------------- #
def _cp(sz: int, color: str = "404040", bold: bool = False) -> CharacterProperties:
    return CharacterProperties(sz=sz, b=bold, solidFill=color)


def _rich(sz: int = 900, color: str = "404040", bold: bool = False) -> RichText:
    cp = _cp(sz, color, bold)
    return RichText(p=[Paragraph(pPr=ParagraphProperties(defRPr=cp), endParaRPr=cp)])


def _chart_title(text: str, color: str) -> Title:
    cp = _cp(1600, color, True)
    para = Paragraph(pPr=ParagraphProperties(defRPr=cp), r=[RegularTextRun(rPr=cp, t=text)])
    return Title(tx=Text(rich=RichText(p=[para])), overlay=False)


def _base_style(ch, title: str, ctx: Ctx):
    ch.title = _chart_title(title, ctx.palette.dark)
    ch.graphical_properties = GraphicalProperties(solidFill=CHART_BG, ln=LineProperties(noFill=True))
    ch.plot_area.graphicalProperties = GraphicalProperties(noFill=True)
    ch.legend = None
    ch.roundedCorners = False


def _axes(ch, value_fmt: str = '#,##0', hide_value_axis: bool = False, hide_cat_axis: bool = False):
    ch.x_axis.delete = hide_cat_axis
    ch.y_axis.delete = hide_value_axis
    ch.y_axis.majorGridlines = None
    ch.y_axis.number_format = value_fmt
    ch.x_axis.txPr = _rich(800)
    ch.y_axis.txPr = _rich(800)
    ch.x_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill="BFBFBF"))
    ch.y_axis.graphicalProperties = GraphicalProperties(ln=LineProperties(noFill=True))


def _series(ws, col: int, r1: int, r2: int, cat_col: int, title: str, color: str) -> Series:
    vals = Reference(ws, min_col=col, min_row=r1, max_row=r2)
    s = Series(vals, title=title)
    s.cat = AxDataSource(strRef=StrRef(f=f"{q(ws.title)}!${get_column_letter(cat_col)}${r1}:${get_column_letter(cat_col)}${r2}"))
    s.graphicalProperties = GraphicalProperties(solidFill=color, ln=LineProperties(solidFill=color))
    return s


def _row_series(ws, row: int, c1: int, c2: int, cat_row: int, title: str) -> Series:
    vals = Reference(ws, min_col=c1, max_col=c2, min_row=row, max_row=row)
    s = Series(vals, title=title)
    s.cat = AxDataSource(strRef=StrRef(
        f=f"{q(ws.title)}!${get_column_letter(c1)}${cat_row}:${get_column_letter(c2)}${cat_row}"))
    return s


def _labels(fmt: str, pos: str = "outEnd", sz: int = 800, color: str = "404040") -> DataLabelList:
    d = DataLabelList()
    d.showVal = True
    d.showCatName = d.showSerName = d.showPercent = d.showLegendKey = False
    d.showLeaderLines = False
    # number format comes from the source cells on Dashboard_Calc (label numFmt without
    # sourceLinked is not schema-valid and can trigger Excel's repair prompt)
    d.position = pos
    d.txPr = _rich(sz, color)
    return d


def _place(ws, ch, key: str):
    c1, r1, c2, r2 = CHART_ANCHORS[key]
    ch.anchor = TwoCellAnchor(editAs="twoCell", _from=AnchorMarker(col=c1, colOff=0, row=r1, rowOff=0),
                              to=AnchorMarker(col=c2, colOff=0, row=r2, rowOff=0))
    ws.add_chart(ch)


# --------------------------------------------------------------------------- #
# charts
# --------------------------------------------------------------------------- #
def add_charts(ws, calc, lay: CalcLayout, ctx: Ctx):
    pal = ctx.palette
    n_all, n_act = ctx.n_all, ctx.n_act

    # 1. Revenue growth Y-o-Y (column, all years; forecast bars lighter)
    ch = BarChart()
    ch.type, ch.grouping, ch.gapWidth = "col", "clustered", 75
    s = _row_series(calc, lay.hist + 1, 2, 1 + n_all, lay.hist, "Revenue (₹ Cr)")
    s.graphicalProperties = GraphicalProperties(solidFill=pal.dark, ln=LineProperties(solidFill=pal.dark))
    for i in range(n_act, n_all):
        dp = DataPoint(idx=i)
        dp.graphicalProperties = GraphicalProperties(solidFill=pal.light, ln=LineProperties(solidFill=pal.light))
        s.dPt.append(dp)
    ch.series.append(s)
    _base_style(ch, "Revenue Growth Y-o-Y (₹ Cr)", ctx)
    _axes(ch, '#,##0')
    _place(ws, ch, "revenue")

    # 2. Expense breakup (horizontal bar)
    ch = BarChart()
    ch.type, ch.grouping, ch.gapWidth = "bar", "clustered", 78
    r2 = lay.items + max(lay.n_items, 1) - 1
    s = _series(calc, 2, lay.items, r2, 1, "₹ Cr", pal.primary)
    s.dLbls = _labels('#,##0')
    ch.series.append(s)
    _base_style(ch, "Expense Breakup (₹ Cr)", ctx)
    _axes(ch, '#,##0', hide_value_axis=True)
    ch.x_axis.scaling.orientation = "maxMin"
    _place(ws, ch, "expenses")

    # 3. Where revenue goes (pie)
    ch = PieChart()
    s = _series(calc, 2, lay.pie, lay.pie + 6, 1, "Share", pal.primary)
    s.graphicalProperties = None
    for i, colr in enumerate(pal.shades(7)):
        dp = DataPoint(idx=i)
        dp.graphicalProperties = GraphicalProperties(solidFill=colr, ln=LineProperties(solidFill="FFFFFF"))
        s.dPt.append(dp)
    lbl = DataLabelList()
    lbl.showPercent = True
    lbl.showVal = lbl.showCatName = lbl.showSerName = lbl.showLegendKey = False
    lbl.showLeaderLines = True
    lbl.position = "bestFit"
    lbl.txPr = _rich(800, "000000", True)
    s.dLbls = lbl
    ch.series.append(s)
    _base_style(ch, "Where Revenue Goes", ctx)
    ch.legend = Legend()
    ch.legend.position = "b"
    ch.legend.txPr = _rich(750)
    _place(ws, ch, "pie")

    # 4. Margins (horizontal bar)
    ch = BarChart()
    ch.type, ch.grouping, ch.gapWidth = "bar", "clustered", 40
    s = _series(calc, 2, lay.margins, lay.margins + 3, 1, "Margin", pal.primary)
    s.dLbls = _labels('0.0%', pos="ctr", color=pal.text_on_primary)
    ch.series.append(s)
    _base_style(ch, "Margins", ctx)
    _axes(ch, '0.0%')
    ch.x_axis.scaling.orientation = "maxMin"
    ch.graphical_properties = GraphicalProperties(solidFill=CHART_BG, ln=LineProperties(solidFill="D9D9D9"))
    _place(ws, ch, "margins")

    # 5. P&L waterfall (stacked column with invisible base)
    ch = BarChart()
    ch.type, ch.grouping, ch.overlap, ch.gapWidth = "col", "stacked", 100, 50
    w1, w2 = lay.wf + 1, lay.wf + 9
    colours = {5: None, 6: pal.primary, 7: pal.primary, 8: RED, 9: RED, 10: pal.dark}
    names = {5: "Base", 6: "Increase", 7: "Increase", 8: "Decrease", 9: "Decrease", 10: "Total"}
    for col in (5, 6, 7, 8, 9, 10):
        s = _series(calc, col, w1, w2, 1, names[col], colours[col] or "FFFFFF")
        if colours[col] is None:
            s.graphicalProperties = GraphicalProperties(noFill=True, ln=LineProperties(noFill=True))
        ch.series.append(s)
    _base_style(ch, "P&L Waterfall (₹ Cr)", ctx)
    _axes(ch, '#,##0')
    _place(ws, ch, "waterfall")

    # 6. Trade cycle days (horizontal bar, funnel-like)
    ch = BarChart()
    ch.type, ch.grouping, ch.gapWidth = "bar", "clustered", 40
    s = _series(calc, 2, lay.trade, lay.trade + 3, 1, "Days", pal.primary)
    s.dLbls = _labels('0', pos="outEnd")
    ch.series.append(s)
    _base_style(ch, "Trade Cycle (Days)", ctx)
    _axes(ch, '0', hide_value_axis=True)
    ch.x_axis.scaling.orientation = "maxMin"
    _place(ws, ch, "trade")

    # 7. Liquidity ratios (horizontal bar, value axis deleted)
    ch = BarChart()
    ch.type, ch.grouping, ch.gapWidth = "bar", "clustered", 30
    s = _series(calc, 2, lay.liq, lay.liq + 2, 1, "x", pal.dark)
    s.dLbls = _labels('0.00', pos="inEnd", color=pal.text_on_dark)
    ch.series.append(s)
    _base_style(ch, "Liquidity Ratios (approx.)", ctx)
    _axes(ch, '0.00', hide_value_axis=True)
    ch.x_axis.scaling.orientation = "maxMin"
    _place(ws, ch, "liquidity")

    # 8. Free cash flow (line, actual vs forecast; not year-dependent)
    ch = LineChart()
    for row, title, colr, dash in ((lay.fcf + 1, "Actual FCF", pal.dark, None), (lay.fcf + 2, "Forecast FCFF", pal.primary, "dash")):
        s = _row_series(calc, row, 2, 1 + n_all, lay.fcf, title)
        s.graphicalProperties = GraphicalProperties(ln=LineProperties(solidFill=colr, w=28575, prstDash=dash))
        s.marker = Marker(symbol="none")
        s.smooth = False
        ch.series.append(s)
    _base_style(ch, "Free Cash Flow (₹ Cr)", ctx)
    ch.legend = Legend()
    ch.legend.position = "b"
    ch.legend.txPr = _rich(750)
    _axes(ch, '#,##0')
    ch.display_blanks = "gap"
    _place(ws, ch, "fcf")


# --------------------------------------------------------------------------- #
# dashboard sheet
# --------------------------------------------------------------------------- #
def _logo_image(png: bytes):
    from openpyxl.drawing.image import Image as XLImage
    from PIL import Image

    im = Image.open(io.BytesIO(png))
    max_w, max_h = 76, 64
    scale = min(max_w / im.width, max_h / im.height)
    im = im.resize((max(1, int(im.width * scale)), max(1, int(im.height * scale))))
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    buf.seek(0)
    xl = XLImage(buf)
    xl.width, xl.height = im.width, im.height
    off_x = max(2, (84 - im.width) // 2)
    off_y = max(2, (104 - im.height) // 2)
    xl.anchor = OneCellAnchor(_from=AnchorMarker(col=1, colOff=pixels_to_EMU(off_x), row=0, rowOff=pixels_to_EMU(off_y)),
                              ext=XDRPositiveSize2D(pixels_to_EMU(im.width), pixels_to_EMU(im.height)))
    return xl


def build_dashboard(wb, ctx: Ctx):
    pal, fin = ctx.palette, ctx.fin
    lay = build_calc(wb, ctx)
    calc = wb[S_CALC]
    ws = wb[S_DASH]
    ws.sheet_view.showGridLines = False
    ws.sheet_view.showRowColHeaders = False
    ws.sheet_view.zoomScale = 90
    ws.sheet_format.defaultColWidth = 11.2
    ws.sheet_format.defaultRowHeight = 15.6
    ws.sheet_properties.tabColor = pal.primary
    for k, w in COL_WIDTHS.items():
        ws.column_dimensions[k].width = w
    for i in range(14):
        ws.column_dimensions[get_column_letter(6 + i)].width = CARD_COL_WIDTH
    for r in range(1, 8):
        ws.row_dimensions[r].height = 15.6 if r != 3 else 18
    ws.row_dimensions[4].height = 18

    # banner B1:E5
    dark = _fill(pal.dark)
    for r in range(1, 6):
        for col in range(2, 6):
            ws.cell(r, col).fill = dark
    name_cells = "C1:E5" if pal.logo_png else "B1:E5"
    ws.merge_cells(name_cells)
    nm = ws[name_cells.split(":")[0]]
    nm.value = fin.name
    size = 20 if len(fin.name) <= 14 else (16 if len(fin.name) <= 24 else 13)
    nm.font = Font(name="Arial", size=size, bold=True, color=pal.text_on_dark)
    nm.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    if pal.logo_png:
        try:
            ws.add_image(_logo_image(pal.logo_png))
        except Exception:
            pass

    # financial year selector
    ws.merge_cells("B6:C7")
    ws.merge_cells("E6:E7")
    for addr in ("B6", "C6", "B7", "C7", "E6", "E7"):
        ws[addr].fill = _fill(pal.dark)
    ws["B6"] = "Financial Year"
    ws["B6"].font = Font(name="Arial", size=14, bold=True, color="E7E6E6" if pal.text_on_dark == "FFFFFF" else "000000")
    ws["B6"].alignment = Alignment(horizontal="center", vertical="center")
    ws["E6"] = fin.years[-1]
    ws["E6"].font = Font(name="Calibri", size=18, bold=True, color="E7E6E6" if pal.text_on_dark == "FFFFFF" else "000000")
    ws["E6"].alignment = Alignment(horizontal="center", vertical="center")
    ws["E6"].number_format = "0"
    dv = DataValidation(type="list", formula1='"' + ",".join(str(y) for y in ctx.years) + '"', allow_blank=False,
                        showDropDown=False)
    dv.promptTitle = "Financial Year"
    dv.prompt = "Pick a fiscal year (A = reported, E = forecast)"
    ws.add_data_validation(dv)
    dv.add("E6")

    # KPI cards F1:S7
    card_fill = _fill(pal.card_bg)
    sep = Side(style="medium", color="FFFFFF")
    for k, (title, key, fmt, is_ratio) in enumerate(CARDS):
        c1, c2 = 6 + 2 * k, 7 + 2 * k
        L1, L2 = get_column_letter(c1), get_column_letter(c2)
        for r in range(1, 8):
            for col in (c1, c2):
                cell = ws.cell(r, col)
                cell.fill = card_fill
                cell.border = Border(left=sep if col == c1 else None, right=sep if col == c2 else None)
        crow = lay.kpi[key]
        ws.merge_cells(f"{L1}1:{L2}2")
        t = ws[f"{L1}1"]
        t.value = title + ("" if is_ratio else " (₹ Cr)")
        t.font = Font(name="Arial", size=12, bold=True, color=pal.dark)
        t.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(f"{L1}3:{L2}4")
        v = ws[f"{L1}3"]
        v.value = f"={q(S_CALC)}!$B${crow}"
        v.number_format = fmt
        v.font = Font(name="Calibri", size=22 if is_ratio else 18, bold=True, color="000000")
        v.alignment = Alignment(horizontal="center", vertical="center")
        ws.merge_cells(f"{L1}5:{L2}5")
        ch = ws[f"{L1}5"]
        ch.value = f"={q(S_CALC)}!$D${crow}"
        ch.number_format = '+0.0" pp";-0.0" pp";0.0" pp"' if is_ratio else '+0.0%;-0.0%;0.0%'
        ch.font = Font(name="Calibri", size=12, bold=True, color=GREEN)
        ch.alignment = Alignment(horizontal="center", vertical="center")
        ws.conditional_formatting.add(f"{L1}5", FormulaRule(formula=[f"AND(ISNUMBER({L1}5),{L1}5<0)"], font=Font(color=RED, bold=True)))
        ws.conditional_formatting.add(f"{L1}5", FormulaRule(formula=[f"AND(ISNUMBER({L1}5),{L1}5>=0)"], font=Font(color=GREEN, bold=True)))
        ws.merge_cells(f"{L1}6:{L2}6")
        lab = ws[f"{L1}6"]
        lab.value = "VS Pre Year"
        lab.font = Font(name="Calibri", size=10, color="595959")
        lab.alignment = Alignment(horizontal="center", vertical="center")

    # share price info T2:X7
    ws.merge_cells("T2:V6")
    hdr = ws["T2"]
    short = fin.name.replace(" Ltd", "").replace(" Limited", "")
    hdr.value = f"{short} Share Price Info."
    hdr.font = Font(name="Arial", size=14, bold=True, color=pal.text_on_light)
    hdr.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    for r in range(2, 7):
        for col in range(20, 23):
            ws.cell(r, col).fill = _fill(pal.light)
    i0 = lay.info
    rows = [("NSE -", fin.symbol, None), ("Sector -", f"={q(S_CALC)}!$B${i0}", None),
            ("Market Cap (₹ Cr) -", f"={q(S_CALC)}!$B${i0 + 1}", F_NUM),
            ("Price (₹) -", f"={q(S_CALC)}!$B${i0 + 2}", '#,##0.00'),
            ("Fair Value (₹) -", f"={q(S_CALC)}!$B${i0 + 3}", '#,##0.00'),
            ("STATUS", f"={q(S_CALC)}!$B${i0 + 4}", None)]
    box = Side(style="thin", color="D9D9D9")
    for k, (label, value, fmt) in enumerate(rows):
        r = 2 + k
        a = ws.cell(r, 23, label)
        a.font = Font(name="Calibri", size=11, bold=label == "STATUS")
        b = ws.cell(r, 24, value)
        b.font = Font(name="Calibri", size=11, bold=True)
        b.alignment = Alignment(horizontal="center" if label == "STATUS" else "left", shrink_to_fit=True)
        if fmt:
            b.number_format = fmt
        for col in (23, 24):
            ws.cell(r, col).border = Border(top=box, bottom=box, left=box, right=box)
    ws.conditional_formatting.add("X7", FormulaRule(formula=['ISNUMBER(SEARCH("ACCUMULATE",$X$7))'],
                                                    font=Font(color="006100", bold=True), fill=_fill("C6EFCE")))
    ws.conditional_formatting.add("X7", FormulaRule(formula=['ISNUMBER(SEARCH("HOLD",$X$7))'],
                                                    font=Font(color="9C5700", bold=True), fill=_fill("FFEB9C")))
    ws.conditional_formatting.add("X7", FormulaRule(formula=['OR(ISNUMBER(SEARCH("SELL",$X$7)),ISNUMBER(SEARCH("REDUCE",$X$7)))'],
                                                    font=Font(color="9C0006", bold=True), fill=_fill("FFC7CE")))

    add_charts(ws, calc, lay, ctx)

    view = "consolidated" if fin.consolidated else "standalone"
    note = ws.cell(41, 2, f"Source: Screener.in ({view}) fetched {fin.fetched_at}. A = reported, E = model forecast "
                          f"({'Base' if ctx.forecast.scenario == 'base' else ctx.forecast.scenario} scenario switch in "
                          f"Drivers!C4). Liquidity ratios are approximations. Not investment advice.")
    note.font = Font(size=8, italic=True, color="7F7F7F")
    return lay
