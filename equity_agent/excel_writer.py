"""Write the financial model workbook.

Actual years are typed-in reported values (blue font). Forecast years are Excel
formulas (black font) that reference the Drivers sheet, so the scenario switch
(Drivers!C4: 1 Base, 2 Bull, 3 Bear) re-runs the whole forecast and valuation.
Each driver also has one cell per forecast year (Drivers!G:K): working-capital
days glide from the last reported year, and sourced analyst overrides are typed
there (orange fill, cell comment with the source). The Python engine in
``model_engine.py`` / ``valuation.py`` computes the same numbers; tests
recalculate the workbook and compare.

The only stored (non-formula) model results are the two non-live scenarios on the
Scenarios sheet; its check cell proves the live scenario equals its stored value.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from datetime import date
from typing import Callable, Dict, List, Optional, Union

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation

from .brand import Palette
from .drivers import DRIVER_ORDER, GLIDE_KEYS, YEAR_DRIVERS, Drivers
from .financials import CompanyFinancials, cost_of_revenue_label, opex_label
from .market_data import MarketData
from .model_engine import Forecast
from .valuation import G_STEPS, WACC_STEPS, Valuation

S_COVER, S_DASH, S_DRV, S_CB = "Cover Page", "Dashboard", "Drivers", "Cost Breakdown"
S_IS, S_PPE, S_WC, S_CF, S_BS = "Income Statement", "PP&E Schedule", "Working Capital", "Cash Flow", "Balance Sheet"
S_WACC, S_DCF, S_CALC = "CAPM & WACC", "DCF Valuation", "Dashboard_Calc"
S_VAL, S_PEERS, S_SCEN, S_AI = "Valuation Summary", "Peers", "Scenarios", "Analyst Inputs"

F_NUM = '#,##0;(#,##0);0'
F_NUM1 = '#,##0.0;(#,##0.0);0.0'
F_PCT = '0.0%'
F_DAYS = '0'
F_X = '0.00"x"'
F_X1 = '0.0"x"'
F_EPS = '#,##0.00'
F_RS = '"₹ "#,##0.00'

BLUE, BLACK, GREY = "0000FF", "000000", "7F7F7F"
OVERRIDE_FILL = "FCE4D6"
THIN = Side(style="thin", color="D9D9D9")
YEAR_COL0 = 7  # Drivers!G = first forecast year

Value = Union[float, int, str, None]


def q(sheet: str) -> str:
    return f"'{sheet}'"


@dataclass
class Ctx:
    fin: CompanyFinancials
    drivers: Drivers
    forecast: Forecast
    val: Valuation
    palette: Palette
    cfg: dict
    md: Optional[MarketData]
    peers: object = None
    analyst: object = None
    flags: list = field(default_factory=list)
    rows: Dict[str, Dict[str, int]] = field(default_factory=dict)
    drv: Dict[str, str] = field(default_factory=dict)  # driver/input key -> absolute ref (active value)
    drow: Dict[str, int] = field(default_factory=dict)  # driver key -> Drivers row
    refs: Dict[str, str] = field(default_factory=dict)  # named results (target, rating, upside, ...)

    @property
    def n_act(self) -> int:
        return self.fin.n

    @property
    def n_all(self) -> int:
        return self.fin.n + len(self.forecast.years)

    @property
    def years(self) -> List[int]:
        return list(self.fin.years) + list(self.forecast.years)

    @property
    def labels(self) -> List[str]:
        return self.fin.year_labels + self.forecast.year_labels

    @staticmethod
    def col(i: int) -> str:
        return get_column_letter(2 + i)

    @property
    def last_col(self) -> str:
        return self.col(self.n_all - 1)

    def r(self, sheet: str, key: str) -> int:
        return self.rows[sheet][key]

    def ref(self, sheet: str, key: str, i: int, absolute: bool = False) -> str:
        c = self.col(i)
        row = self.r(sheet, key)
        return f"{q(sheet)}!${c}${row}" if absolute else f"{q(sheet)}!{c}{row}"

    def dy(self, key: str, i: int) -> str:
        """Per-year driver cell for overall column index ``i`` (a forecast year)."""
        t = i - self.n_act
        return f"Drivers!${get_column_letter(YEAR_COL0 + t)}${self.drow[key]}"


# --------------------------------------------------------------------------- #
# styling helpers
# --------------------------------------------------------------------------- #
def _fill(hex_: str) -> PatternFill:
    return PatternFill("solid", start_color=hex_, end_color=hex_)


def _title(ws, text: str, sub: str, pal: Palette):
    ws["A1"] = text
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color=pal.dark)
    ws["A2"] = sub
    ws["A2"].font = Font(name="Calibri", size=9, italic=True, color=GREY)


def _period_header(ws, ctx: Ctx, n_cols: Optional[int] = None):
    pal = ctx.palette
    n = ctx.n_all if n_cols is None else n_cols
    ws.column_dimensions["A"].width = 46
    ws["A3"] = "₹ Cr (unless stated)"
    ws["A3"].font = Font(size=9, italic=True, color=GREY)
    ws["A4"] = "Fiscal year"
    ws["A4"].font = Font(size=8, color=GREY)
    for i in range(n):
        c = ctx.col(i)
        ws.column_dimensions[c].width = 11.5
        actual = i < ctx.n_act
        h = ws[f"{c}3"]
        h.value = ctx.labels[i]
        h.font = Font(bold=True, color=pal.text_on_dark if actual else pal.text_on_light)
        h.fill = _fill(pal.dark if actual else pal.light)
        h.alignment = Alignment(horizontal="center")
        y = ws[f"{c}4"]
        y.value = ctx.years[i]
        y.font = Font(size=8, color=GREY)
        y.alignment = Alignment(horizontal="center")
    ws.freeze_panes = "B5"
    ws.sheet_view.showGridLines = False


def _layout(keys: List[Optional[Union[str, tuple]]], start: int = 5) -> Dict[str, int]:
    out, r = {}, start
    for k in keys:
        if isinstance(k, str):
            out[k] = r
        elif isinstance(k, tuple):
            out[k[0]] = r
        r += 1
    return out


def _write_row(ws, ctx: Ctx, sheet: str, key: str, label: str, actual: Union[List[Value], Callable[[int], Value], None],
               forecast: Union[Callable[[int], Value], None], fmt: str = F_NUM, bold: bool = False,
               indent: int = 0, top_border: bool = False):
    r = ctx.r(sheet, key)
    a = ws.cell(r, 1, label)
    a.font = Font(bold=bold)
    a.alignment = Alignment(indent=indent)
    for i in range(ctx.n_all):
        is_act = i < ctx.n_act
        src = actual if is_act else forecast
        if src is None:
            continue
        v = src[i] if isinstance(src, list) else src(i)
        if v is None:
            continue
        cell = ws.cell(r, 2 + i, v)
        is_formula = isinstance(v, str) and v.startswith("=")
        cell.font = Font(bold=bold, color=BLACK if is_formula else (BLUE if not isinstance(v, str) else BLACK))
        cell.number_format = fmt
        if top_border:
            cell.border = Border(top=Side(style="thin", color="808080"))
    if top_border:
        a.border = Border(top=Side(style="thin", color="808080"))


def _section(ws, row: int, text: str, pal: Palette, width: int):
    c = ws.cell(row, 1, text)
    c.font = Font(bold=True, color=pal.dark)
    for i in range(1, width + 2):
        ws.cell(row, i).border = Border(bottom=Side(style="thin", color=pal.primary))


def _head_row(ws, row: int, heads: List[str], pal: Palette, start_col: int = 1):
    for j, h in enumerate(heads, start_col):
        c = ws.cell(row, j, h)
        c.font = Font(bold=True, color=pal.text_on_dark)
        c.fill = _fill(pal.dark)
        c.alignment = Alignment(horizontal="center", wrap_text=True)


def _put(ws, row: int, col: int, value, fmt: Optional[str] = None, bold: bool = False, color: Optional[str] = None):
    c = ws.cell(row, col, value)
    if fmt:
        c.number_format = fmt
    is_formula = isinstance(value, str) and value.startswith("=")
    c.font = Font(bold=bold, color=color or (BLACK if is_formula or isinstance(value, str) else BLUE))
    return c


# --------------------------------------------------------------------------- #
# layouts (row numbers must be known before formulas reference other sheets)
# --------------------------------------------------------------------------- #
IS_KEYS = ["sales", "growth", "cogs", "gross_profit", "gross_margin", "opex", "ebitda", "ebitda_margin", None,
           "other_income", "depreciation", "ebit", "interest", "pbt", "pbt_margin", "tax", "tax_rate",
           "minority_other", "net_profit", "net_margin", None, "eps", "dividend_payout", "dividends"]
PPE_KEYS = ["opening_net", "capex", "cwip_transfer", "depreciation", "other_movements", "closing_net", None,
            "gross_block", "acc_dep", "cwip", None, "capex_pct", "dep_rate_hist"]
WC_KEYS = ["receivables", "inventory", "trade_payables", "other_ca", "other_liab", "nwc", "delta_nwc", None,
           "loans_advances", None, "dso", "dio", "dpo", "ccc"]
CF_KEYS = ["net_profit", "depreciation", "wc_other", "cfo", None, "capex", "other_investing", "cfi", None,
           "dividends", "other_financing", "cff", None, "net_cash_flow", "opening_cash", "closing_cash", None, "fcf"]
BS_KEYS = [("h_le",), "equity_capital", "reserves", "total_equity", "minority_interest", "lt_borrowings",
           "st_borrowings", "lease_liabilities", "other_borrowings", "borrowings", "trade_payables", "other_liab",
           "total_liabilities", None, ("h_a",), "fixed_assets", "cwip", "investments", "receivables", "inventory", "cash",
           "other_ca", "loans_advances", "total_assets", None, "balance_check", None, ("h_r",), "roa", "roe", "roce",
           "debt_equity", "invested_capital", "current_assets", "current_liabilities", "current_ratio", "quick_ratio",
           "cash_ratio"]


def _cb_keys(fin: CompanyFinancials) -> List:
    items = sorted(range(len(fin.expense_items)), key=lambda k: fin.expense_items[k].kind != "COGS")
    keys: List = [("h_items",)]
    keys += [f"item_{k}" for k in items]
    keys += [None, "total_cogs", "total_opex", "total_items", "reported_expenses", "rounding", None, ("h_pct",)]
    keys += [f"pct_{k}" for k in items]
    return keys


def build_layouts(ctx: Ctx):
    ctx.rows[S_IS] = _layout(IS_KEYS)
    ctx.rows[S_PPE] = _layout(PPE_KEYS)
    ctx.rows[S_WC] = _layout(WC_KEYS)
    ctx.rows[S_CF] = _layout(CF_KEYS)
    ctx.rows[S_BS] = _layout(BS_KEYS)
    ctx.rows[S_CB] = _layout(_cb_keys(ctx.fin))


# --------------------------------------------------------------------------- #
# Drivers
# --------------------------------------------------------------------------- #
def write_drivers(wb, ctx: Ctx):
    ws = wb.create_sheet(S_DRV)
    pal, fin, n = ctx.palette, ctx.fin, ctx.n_act
    _title(ws, f"{fin.name} - Forecast Drivers & Valuation Inputs",
           f"Drivers are {len(ctx.drivers.window_years)}-year averages of reported history "
           f"(FY{ctx.drivers.window_years[0] % 100:02d}-FY{ctx.drivers.window_years[-1] % 100:02d}). "
           "Blue cells are inputs you can edit; orange cells are sourced analyst inputs.", pal)
    ws.column_dimensions["A"].width = 46
    for c, w in zip("BCDEF", (12, 11, 11, 11, 11)):
        ws.column_dimensions[c].width = w
    nf = len(ctx.forecast.years)
    for t in range(nf):
        ws.column_dimensions[get_column_letter(YEAR_COL0 + t)].width = 10
    basis_col = YEAR_COL0 + nf
    ws.column_dimensions[get_column_letter(basis_col)].width = 80
    ws["A4"] = "Scenario (1 = Base, 2 = Bull, 3 = Bear)"
    ws["A4"].font = Font(bold=True)
    ws["C4"] = 1
    ws["C4"].font = Font(bold=True, color=BLUE)
    ws["C4"].fill = _fill("FFF2CC")
    ws["C4"].alignment = Alignment(horizontal="center")
    ws["D4"] = '=CHOOSE($C$4,"Base","Bull","Bear")'
    ws["D4"].font = Font(bold=True)
    dv = DataValidation(type="list", formula1='"1,2,3"', allow_blank=False)
    ws.add_data_validation(dv)
    dv.add("C4")
    heads = ["Driver", "Historical", "Base", "Bull", "Bear", "Active"] + ctx.forecast.year_labels + ["Basis"]
    _head_row(ws, 6, heads, pal)

    # input block row numbers first (glide formulas need the glide-years cell)
    r_inputs0 = 7 + len(DRIVER_ORDER) + 2
    glide_ref = f"Drivers!$C${r_inputs0 + 13}"
    r = 7
    for key in DRIVER_ORDER:
        ctx.drow[key] = r
        ctx.drv[key] = f"Drivers!$F${r}"
        r += 1
    wc_rows = _layout(WC_KEYS)
    last_a = ctx.col(n - 1)
    for key in DRIVER_ORDER:
        d = ctx.drivers.items[key]
        r = ctx.drow[key]
        fmt = F_PCT if d.unit == "pct" else '0.0'
        ws.cell(r, 1, d.label)
        if d.historical is not None:
            h = ws.cell(r, 2, d.historical)
            h.number_format = fmt
            h.font = Font(color=GREY)
        for j, v in enumerate((d.base, d.bull, d.bear), 3):
            c = ws.cell(r, j, v)
            c.number_format = fmt
            c.font = Font(color=BLUE)
        a = ws.cell(r, 6, f"=CHOOSE($C$4,C{r},D{r},E{r})")
        a.number_format = fmt
        a.font = Font(bold=True)
        if key in YEAR_DRIVERS:
            for t in range(nf):
                col = YEAR_COL0 + t
                ov = ctx.drivers.overrides.get(key, {}).get(t)
                if ov is not None:
                    f = f"={ov.value}+($F${r}-$C${r})"
                elif key in GLIDE_KEYS:
                    base0 = f"{q(S_WC)}!${last_a}${wc_rows[key]}"
                    f = f"={base0}+($F${r}-{base0})*MIN(1,{t + 1}/{glide_ref})"
                else:
                    f = f"=$F${r}"
                c = ws.cell(r, col, f)
                c.number_format = fmt
                if ov is not None:
                    c.fill = _fill(OVERRIDE_FILL)
                    c.font = Font(bold=True, color="C55A11")
                    c.comment = Comment(f"Analyst input: {ov.value}\nSource: {ov.source}", "equity-research-agent")
        basis = d.basis + (" (clamped)" if d.clamped else "")
        if key in GLIDE_KEYS:
            basis += f"; per-year values glide from FY{fin.years[-1] % 100:02d} actual"
        if key in ctx.drivers.overrides:
            basis += "; analyst override in orange cells"
        ws.cell(r, basis_col, basis).font = Font(size=9, color=GREY)

    r = r_inputs0 - 1
    _section(ws, r, "Valuation & market inputs", pal, 6)
    r = r_inputs0
    cfg, val, ai = ctx.cfg, ctx.val, ctx.analyst
    w = cfg["target_weights"]
    g_input = ai.terminal_growth if ai is not None and ai.terminal_growth is not None else cfg["terminal_growth"]
    g_note = ("Analyst input: " + ai._terminal_src) if ai is not None and ai.terminal_growth is not None \
        else cfg.get("terminal_growth_source", "")
    ronic_ovr = ai.terminal_ronic if ai is not None and ai.terminal_ronic is not None else None
    beta_note = val.beta_source + (f" (raw regression beta {val.beta_raw:.2f})" if val.beta_raw is not None else "")
    inputs = [
        ("rf", "Risk-free rate (INR)", cfg["risk_free_rate"], F_PCT, cfg.get("risk_free_source", "")),
        ("erp", "Equity risk premium (India)", cfg["equity_risk_premium"], F_PCT, cfg.get("equity_risk_premium_source", "")),
        ("beta", "Beta (Blume-adjusted regression vs Nifty 50)", val.beta, '0.00', beta_note),
        ("g", "Terminal growth", g_input, F_PCT, g_note),
        ("w_dcf", "Target weight: DCF", w.get("dcf", 0.5), '0%', "Renormalised over available methods"),
        ("w_pe", "Target weight: P/E vs own history", w.get("pe", 0.25), '0%', ""),
        ("w_ev", "Target weight: EV/EBITDA vs peers", w.get("ev_ebitda", 0.25), '0%', ""),
        ("ntm_w", "NTM weight of first forecast year", val.ntm_weight, '0.00',
         f"Share of the next 12 months in {ctx.forecast.year_labels[0]} (valuation date {val.valuation_date})"),
        ("div_yield", "Dividend yield", val.dividend_yield, '0.00%', "Screener.in; used for roll-forward and total return"),
        ("roic_cap", "Cap on current ROIC", cfg.get("roic_cap", 1.0), '0%', "Limits reinvestment math for asset-light firms"),
        ("ronic_ovr", "Terminal RONIC override (blank = rule)", ronic_ovr, F_PCT,
         ("Analyst input: " + ai._terminal_src) if ronic_ovr is not None else "Blank: RONIC = WACC + share x (ROIC - WACC)"),
        ("cwip_years", "Years to capitalise CWIP", int(cfg.get("cwip_years", 2)), '0', "CWIP moves into the block evenly"),
        ("fade_years", "Fade years (structural; edit in config)", int(cfg.get("fade_years", 5)), '0', ""),
        ("glide", "Working-capital glide years", int(cfg.get("wc_glide_years", 2)), '0',
         "Days move from the last reported year to the driver over this many years"),
        ("price", "Current share price (₹)", ctx.fin.price, F_EPS, f"Screener.in close, fetched {ctx.fin.fetched_at}"),
        ("mcap", "Market capitalisation (₹ Cr)", ctx.fin.market_cap, F_NUM, "Screener.in"),
    ]
    for key, label, value, fmt, note in inputs:
        ws.cell(r, 1, label)
        c = ws.cell(r, 3, value)
        c.number_format = fmt
        c.font = Font(color=BLUE)
        if key in ("ronic_ovr", "g") and ai is not None and ((key == "g" and ai.terminal_growth is not None)
                                                             or (key == "ronic_ovr" and ronic_ovr is not None)):
            c.fill = _fill(OVERRIDE_FILL)
        ws.cell(r, basis_col, note).font = Font(size=9, color=GREY)
        ctx.drv[key] = f"Drivers!$C${r}"
        r += 1
    assert ctx.drv["glide"] == glide_ref.replace("Drivers!", "Drivers!"), "glide cell moved"
    ws.cell(r, 1, "Shares outstanding (Cr) = market cap / price")
    c = ws.cell(r, 3, f"={ctx.drv['mcap']}/{ctx.drv['price']}")
    c.number_format = '#,##0.00'
    ctx.drv["shares"] = f"Drivers!$C${r}"
    r += 2
    if ctx.drivers.notes:
        ws.cell(r, 1, "Notes").font = Font(bold=True)
        for note in ctx.drivers.notes:
            r += 1
            ws.cell(r, 1, note).font = Font(size=9, color=GREY)
    ws.freeze_panes = "B7"
    ws.sheet_view.showGridLines = False


# --------------------------------------------------------------------------- #
# Income statement
# --------------------------------------------------------------------------- #
def write_income_statement(wb, ctx: Ctx):
    ws = wb.create_sheet(S_IS)
    fin, L, n = ctx.fin, ctx.fin.lines, ctx.n_act
    view = "consolidated" if fin.consolidated else "standalone"
    _title(ws, f"{fin.name} - Income Statement",
           f"Actuals: Screener.in {view} figures (reported). Forecast: formulas on the Drivers sheet.", ctx.palette)
    _period_header(ws, ctx)
    R = ctx.rows[S_IS]
    c = ctx.col
    D, dy = ctx.drv, ctx.dy

    def prev(key, i):
        return f"{c(i - 1)}{R[key]}"

    w = lambda *a, **k: _write_row(ws, ctx, S_IS, *a, **k)
    w("sales", "Revenue (Sales)", L["sales"], lambda i: f"={prev('sales', i)}*(1+{dy('sales_growth', i)})", bold=True)
    w("growth", "  Revenue growth %", lambda i: None if i == 0 else f"=IFERROR({c(i)}{R['sales']}/{prev('sales', i)}-1,0)",
      lambda i: f"=IFERROR({c(i)}{R['sales']}/{prev('sales', i)}-1,0)", fmt=F_PCT)
    w("cogs", cost_of_revenue_label(fin), L["cogs"], lambda i: f"={c(i)}{R['sales']}*{dy('cogs_pct', i)}")
    w("gross_profit", "Gross profit", lambda i: f"={c(i)}{R['sales']}-{c(i)}{R['cogs']}",
      lambda i: f"={c(i)}{R['sales']}-{c(i)}{R['cogs']}", bold=True, top_border=True)
    pct = lambda key: (lambda i: f"=IFERROR({c(i)}{R[key]}/{c(i)}{R['sales']},0)")
    w("gross_margin", "  Gross margin %", pct("gross_profit"), pct("gross_profit"), fmt=F_PCT)
    w("opex", opex_label(fin), L["opex"], lambda i: f"={c(i)}{R['gross_profit']}-{c(i)}{R['ebitda']}")
    w("ebitda", "EBITDA (Operating profit)", L["ebitda"], lambda i: f"={c(i)}{R['sales']}*{dy('ebitda_margin', i)}",
      bold=True, top_border=True)
    w("ebitda_margin", "  EBITDA margin %", pct("ebitda"), pct("ebitda"), fmt=F_PCT)
    w("other_income", "Other income (incl. exceptional items in actuals)", L["other_income"],
      lambda i: f"={c(i)}{R['sales']}*{dy('other_income_pct', i)}")
    if ctx.drivers.dep_basis == "gross":
        dep_f = lambda i: f"={q(S_PPE)}!{c(i - 1)}{ctx.r(S_PPE, 'gross_block')}*{dy('dep_rate', i)}"
    else:
        dep_f = lambda i: f"={q(S_PPE)}!{c(i - 1)}{ctx.r(S_PPE, 'closing_net')}*{dy('dep_rate', i)}"
    w("depreciation", "Depreciation & amortisation", L["depreciation"], dep_f)
    ebit = lambda i: f"={c(i)}{R['ebitda']}+{c(i)}{R['other_income']}-{c(i)}{R['depreciation']}"
    w("ebit", "EBIT (after other income)", ebit, ebit, bold=True, top_border=True)
    w("interest", "Interest (finance cost)", L["interest"],
      lambda i: f"={q(S_BS)}!{c(i - 1)}{ctx.r(S_BS, 'borrowings')}*{dy('interest_rate', i)}")
    w("pbt", "Profit before tax", L["pbt"], lambda i: f"={c(i)}{R['ebit']}-{c(i)}{R['interest']}", bold=True, top_border=True)
    w("pbt_margin", "  PBT margin %", pct("pbt"), pct("pbt"), fmt=F_PCT)
    w("tax", "Tax (actual: PBT x reported tax %)", L["tax"], lambda i: f"=MAX({c(i)}{R['pbt']},0)*{dy('tax_rate', i)}")
    w("tax_rate", "  Effective tax rate %", lambda i: f"=IFERROR({c(i)}{R['tax']}/{c(i)}{R['pbt']},0)",
      lambda i: f"=IFERROR({c(i)}{R['tax']}/{c(i)}{R['pbt']},0)", fmt=F_PCT)
    w("minority_other", "Minority interest, associates & rounding", L["minority_other"],
      lambda i: f"=({c(i)}{R['pbt']}-{c(i)}{R['tax']})*{dy('minority_pct', i)}")
    w("net_profit", "Net profit (reported)", L["net_profit"],
      lambda i: f"={c(i)}{R['pbt']}-{c(i)}{R['tax']}-{c(i)}{R['minority_other']}", bold=True, top_border=True)
    w("net_margin", "  Net margin %", pct("net_profit"), pct("net_profit"), fmt=F_PCT)
    eps_act = [fin.eps[i] if fin.eps[i] is not None else f"={c(i)}{R['net_profit']}/{D['shares']}" for i in range(n)]
    w("eps", "EPS (₹) (actual: Screener; forecast: net profit / current shares)", eps_act,
      lambda i: f"={c(i)}{R['net_profit']}/{D['shares']}", fmt=F_EPS)
    w("dividend_payout", "Dividend payout %", L["dividend_payout"], lambda i: f"={dy('dividend_payout', i)}", fmt=F_PCT)
    div = lambda i: f"=MAX({c(i)}{R['net_profit']},0)*{c(i)}{R['dividend_payout']}"
    w("dividends", "Dividends (net profit x payout)", div, div)


# --------------------------------------------------------------------------- #
# Cost breakdown
# --------------------------------------------------------------------------- #
def write_cost_breakdown(wb, ctx: Ctx):
    ws = wb.create_sheet(S_CB)
    fin, n = ctx.fin, ctx.n_act
    _title(ws, f"{fin.name} - Cost Breakdown",
           "Actuals: Screener.in expense schedule (₹ Cr). Forecast: each item keeps its last reported share of COGS / "
           "operating expenses.", ctx.palette)
    _period_header(ws, ctx)
    R = ctx.rows[S_CB]
    c = ctx.col
    last = n - 1
    _section(ws, R["h_items"], "Expense items (₹ Cr)", ctx.palette, ctx.n_all)
    cogs_r, opex_r = ctx.r(S_IS, "cogs"), ctx.r(S_IS, "opex")
    for k, it in enumerate(fin.expense_items):
        key = f"item_{k}"
        base_row = cogs_r if it.kind == "COGS" else opex_r
        rr = R[key]
        fc = lambda i, rr=rr, base_row=base_row: (
            f"=IFERROR(${c(last)}{rr}/{q(S_IS)}!${c(last)}${base_row}*{q(S_IS)}!{c(i)}{base_row},0)")
        _write_row(ws, ctx, S_CB, key, f"{it.name} [{it.kind}]", it.values, fc)
    cogs_keys = [R[f"item_{k}"] for k, it in enumerate(fin.expense_items) if it.kind == "COGS"]
    item_rows = [R[f"item_{k}"] for k in range(len(fin.expense_items))]

    def sum_of(rows):
        return lambda i: "=" + ("+".join(f"{c(i)}{r}" for r in rows) if rows else "0")

    _write_row(ws, ctx, S_CB, "total_cogs", "Total COGS items", sum_of(cogs_keys), sum_of(cogs_keys), bold=True, top_border=True)
    _write_row(ws, ctx, S_CB, "total_opex", "Operating expenses (Income Statement)",
               lambda i: f"={q(S_IS)}!{c(i)}{opex_r}", lambda i: f"={q(S_IS)}!{c(i)}{opex_r}")
    _write_row(ws, ctx, S_CB, "total_items", "Sum of expense items", sum_of(item_rows), sum_of(item_rows), bold=True)
    _write_row(ws, ctx, S_CB, "reported_expenses", "Reported total expenses (Screener)", fin.lines["expenses"], None)
    _write_row(ws, ctx, S_CB, "rounding", "Rounding difference (reported - items)",
               lambda i: f"={c(i)}{R['reported_expenses']}-{c(i)}{R['total_items']}", None)
    _section(ws, R["h_pct"], "Share of revenue", ctx.palette, ctx.n_all)
    for k, it in enumerate(fin.expense_items):
        f = lambda i, k=k: f"=IFERROR({c(i)}{R[f'item_{k}']}/{q(S_IS)}!{c(i)}{ctx.r(S_IS, 'sales')},0)"
        _write_row(ws, ctx, S_CB, f"pct_{k}", f"{it.name} % of revenue", f, f, fmt=F_PCT)


# --------------------------------------------------------------------------- #
# PP&E
# --------------------------------------------------------------------------- #
def write_ppe(wb, ctx: Ctx):
    ws = wb.create_sheet(S_PPE)
    fin, L, n = ctx.fin, ctx.fin.lines, ctx.n_act
    _title(ws, f"{fin.name} - PP&E Schedule",
           "Net block roll-forward. 'Other movements' in actual years = acquisitions, leases (ROU), disposals and "
           "revaluations implied by the reported balances. Forecast: CWIP is capitalised evenly over "
           "Drivers 'Years to capitalise CWIP'.", ctx.palette)
    _period_header(ws, ctx)
    R = ctx.rows[S_PPE]
    c = ctx.col
    is_r = ctx.rows[S_IS]
    last_a = c(n - 1)
    cwip0 = f"${last_a}${R['cwip']}"
    w = lambda *a, **k: _write_row(ws, ctx, S_PPE, *a, **k)
    w("opening_net", "Opening net block", lambda i: None if i == 0 else f"={c(i - 1)}{R['closing_net']}",
      lambda i: f"={c(i - 1)}{R['closing_net']}")
    w("capex", "Add: capex (fixed assets purchased)", L["capex"],
      lambda i: f"={q(S_IS)}!{c(i)}{is_r['sales']}*{ctx.dy('capex_pct', i)}")
    w("cwip_transfer", "Add: CWIP capitalised (forecast)", None,
      lambda i: f"=IF({cwip0}>0,MIN({c(i - 1)}{R['cwip']},{cwip0}/{ctx.drv['cwip_years']}),0)")
    w("depreciation", "Less: depreciation", lambda i: f"={q(S_IS)}!{c(i)}{is_r['depreciation']}",
      lambda i: f"={q(S_IS)}!{c(i)}{is_r['depreciation']}")
    w("other_movements", "Other movements (balancing, actuals only)",
      lambda i: None if i == 0 else f"={c(i)}{R['closing_net']}-{c(i)}{R['opening_net']}-{c(i)}{R['capex']}+{c(i)}{R['depreciation']}",
      lambda i: 0)
    w("closing_net", "Closing net block (Fixed assets)", L["fixed_assets"],
      lambda i: (f"={c(i)}{R['opening_net']}+{c(i)}{R['capex']}+{c(i)}{R['cwip_transfer']}-{c(i)}{R['depreciation']}"
                 f"+{c(i)}{R['other_movements']}"),
      bold=True, top_border=True)
    if fin.gross_block is not None:
        w("gross_block", "Gross block", fin.gross_block,
          lambda i: f"={c(i - 1)}{R['gross_block']}+{c(i)}{R['capex']}+{c(i)}{R['cwip_transfer']}")
        w("acc_dep", "Accumulated depreciation", fin.acc_dep, lambda i: f"={c(i - 1)}{R['acc_dep']}+{c(i)}{R['depreciation']}")
    else:
        ws.cell(R["gross_block"], 1, "Gross block (not reported by Screener)")
        ws.cell(R["acc_dep"], 1, "Accumulated depreciation (not reported by Screener)")
    w("cwip", "Capital work in progress", L["cwip"], lambda i: f"={c(i - 1)}{R['cwip']}-{c(i)}{R['cwip_transfer']}")
    w("capex_pct", "Capex % of revenue", lambda i: f"=IFERROR({c(i)}{R['capex']}/{q(S_IS)}!{c(i)}{is_r['sales']},0)",
      lambda i: f"=IFERROR({c(i)}{R['capex']}/{q(S_IS)}!{c(i)}{is_r['sales']},0)", fmt=F_PCT)
    base_key = "gross_block" if fin.gross_block is not None else "closing_net"
    w("dep_rate_hist", f"Depreciation % of opening {'gross' if base_key == 'gross_block' else 'net'} block",
      lambda i: None if i == 0 else f"=IFERROR({c(i)}{R['depreciation']}/{c(i - 1)}{R[base_key]},0)",
      lambda i: f"=IFERROR({c(i)}{R['depreciation']}/{c(i - 1)}{R[base_key]},0)", fmt=F_PCT)


# --------------------------------------------------------------------------- #
# Working capital
# --------------------------------------------------------------------------- #
def write_working_capital(wb, ctx: Ctx):
    ws = wb.create_sheet(S_WC)
    fin, L, n = ctx.fin, ctx.fin.lines, ctx.n_act
    _title(ws, f"{fin.name} - Working Capital",
           "Days on year-end balances: DSO = receivables / revenue x 365; DIO = inventories / COGS x 365; DPO = "
           "trade payables / " + ("COGS" if ctx.drivers.dpo_basis == "cogs" else "total expenses") + " x 365 "
           "(the same definitions as the PDF report). Loans & advances are non-operating (held flat, outside working "
           "capital).", ctx.palette)
    _period_header(ws, ctx)
    R = ctx.rows[S_WC]
    c, dy = ctx.col, ctx.dy
    is_r = ctx.rows[S_IS]
    sales = lambda i: f"{q(S_IS)}!{c(i)}{is_r['sales']}"
    cogs = lambda i: f"{q(S_IS)}!{c(i)}{is_r['cogs']}"
    if ctx.drivers.dpo_basis == "cogs":
        pay_base = cogs
    else:
        pay_base = lambda i: f"({q(S_IS)}!{c(i)}{is_r['cogs']}+{q(S_IS)}!{c(i)}{is_r['opex']})"
    w = lambda *a, **k: _write_row(ws, ctx, S_WC, *a, **k)
    w("receivables", "Trade receivables", L["receivables"], lambda i: f"={sales(i)}*{dy('dso', i)}/365")
    w("inventory", "Inventories", L["inventory"], lambda i: f"={cogs(i)}*{dy('dio', i)}/365")
    w("trade_payables", "Trade payables", L["trade_payables"], lambda i: f"={pay_base(i)}*{dy('dpo', i)}/365")
    ol = [a + b for a, b in zip(L["other_liability_items"], L["advance_from_customers"])]
    w("other_ca", "Other operating assets (other asset items)", L["other_asset_items"],
      lambda i: f"={sales(i)}*{dy('other_ca_pct', i)}")
    w("other_liab", "Other liabilities (other items + customer advances)", ol,
      lambda i: f"={sales(i)}*{dy('other_liab_pct', i)}")
    nwc = lambda i: (f"={c(i)}{R['receivables']}+{c(i)}{R['inventory']}+{c(i)}{R['other_ca']}"
                     f"-{c(i)}{R['trade_payables']}-{c(i)}{R['other_liab']}")
    w("nwc", "Operating net working capital", nwc, nwc, bold=True, top_border=True)
    w("delta_nwc", "Change in net working capital", lambda i: None if i == 0 else f"={c(i)}{R['nwc']}-{c(i - 1)}{R['nwc']}",
      lambda i: f"={c(i)}{R['nwc']}-{c(i - 1)}{R['nwc']}")
    w("loans_advances", "Loans & advances (non-operating, held flat)", L["loans_advances"],
      lambda i: f"={c(i - 1)}{R['loans_advances']}")
    # n/a when the base is not positive (equity_agent.ratios.days returns None there)
    day = lambda bal, base: (lambda i: f'=IF({base(i)}>0,{c(i)}{R[bal]}/{base(i)}*365,"n/a")')
    ccc = lambda i: (f'=IF(OR(ISTEXT({c(i)}{R["dso"]}),ISTEXT({c(i)}{R["dpo"]})),"n/a",'
                     f'{c(i)}{R["dso"]}+N({c(i)}{R["dio"]})-{c(i)}{R["dpo"]})')
    dso, dio, dpo = day("receivables", sales), day("inventory", cogs), day("trade_payables", pay_base)
    w("dso", "Receivable days (DSO)", dso, dso, fmt=F_DAYS)
    w("dio", "Inventory days (DIO)", dio, dio, fmt=F_DAYS)
    w("dpo", "Payable days (DPO)" + ("" if ctx.drivers.dpo_basis == "cogs" else " on total expenses"), dpo, dpo, fmt=F_DAYS)
    w("ccc", "Cash conversion cycle (days)", ccc, ccc, fmt=F_DAYS, bold=True)


# --------------------------------------------------------------------------- #
# Cash flow
# --------------------------------------------------------------------------- #
def write_cash_flow(wb, ctx: Ctx):
    ws = wb.create_sheet(S_CF)
    fin, L, n = ctx.fin, ctx.fin.lines, ctx.n_act
    _title(ws, f"{fin.name} - Cash Flow Statement",
           "Actual CFO/CFI/CFF are reported totals; component lines in actual years are derived from them. "
           "Forecast: indirect method from the model.", ctx.palette)
    _period_header(ws, ctx)
    R = ctx.rows[S_CF]
    c = ctx.col
    is_r, wc_r, ppe_r, bs_r = ctx.rows[S_IS], ctx.rows[S_WC], ctx.rows[S_PPE], ctx.rows[S_BS]
    w = lambda *a, **k: _write_row(ws, ctx, S_CF, *a, **k)
    np_ = lambda i: f"={q(S_IS)}!{c(i)}{is_r['net_profit']}"
    dep = lambda i: f"={q(S_IS)}!{c(i)}{is_r['depreciation']}"
    w("net_profit", "Net profit", np_, np_)
    w("depreciation", "Add: depreciation", dep, dep)
    w("wc_other", "Working capital & other items",
      lambda i: f"={c(i)}{R['cfo']}-{c(i)}{R['net_profit']}-{c(i)}{R['depreciation']}",
      lambda i: f"=-{q(S_WC)}!{c(i)}{wc_r['delta_nwc']}")
    w("cfo", "Cash from operating activities", L["cfo"],
      lambda i: f"={c(i)}{R['net_profit']}+{c(i)}{R['depreciation']}+{c(i)}{R['wc_other']}", bold=True, top_border=True)
    w("capex", "Capex", [-v for v in L["capex"]], lambda i: f"=-{q(S_PPE)}!{c(i)}{ppe_r['capex']}")
    w("other_investing", "Other investing (investments, acquisitions, interest received)",
      lambda i: f"={c(i)}{R['cfi']}-{c(i)}{R['capex']}", lambda i: 0)
    w("cfi", "Cash from investing activities", L["cfi"], lambda i: f"={c(i)}{R['capex']}+{c(i)}{R['other_investing']}",
      bold=True, top_border=True)
    w("dividends", "Dividends paid", [-v for v in L["dividends_paid"]], lambda i: f"=-{q(S_IS)}!{c(i)}{is_r['dividends']}")
    w("other_financing", "Borrowings & other financing",
      lambda i: f"={c(i)}{R['cff']}-{c(i)}{R['dividends']}",
      lambda i: f"={q(S_BS)}!{c(i)}{bs_r['borrowings']}-{q(S_BS)}!{c(i - 1)}{bs_r['borrowings']}")
    w("cff", "Cash from financing activities", L["cff"], lambda i: f"={c(i)}{R['dividends']}+{c(i)}{R['other_financing']}",
      bold=True, top_border=True)
    ncf = lambda i: f"={c(i)}{R['cfo']}+{c(i)}{R['cfi']}+{c(i)}{R['cff']}"
    w("net_cash_flow", "Net cash flow", L["net_cash_flow"], ncf, bold=True)
    w("opening_cash", "Opening cash & equivalents", lambda i: None if i == 0 else f"={c(i - 1)}{R['closing_cash']}",
      lambda i: f"={c(i - 1)}{R['closing_cash']}")
    w("closing_cash", "Closing cash & equivalents (actual: balance sheet)", lambda i: f"={q(S_BS)}!{c(i)}{bs_r['cash']}",
      lambda i: f"={c(i)}{R['opening_cash']}+{c(i)}{R['net_cash_flow']}", bold=True)
    w("fcf", "Free cash flow (actual: CFO - capex reported; forecast: FCFF)", L["fcf"],
      lambda i: f"={q(S_DCF)}!{get_column_letter(2 + i - n)}{DCF_ROWS['fcff']}", bold=True)


# --------------------------------------------------------------------------- #
# Balance sheet
# --------------------------------------------------------------------------- #
def write_balance_sheet(wb, ctx: Ctx):
    ws = wb.create_sheet(S_BS)
    fin, L, n = ctx.fin, ctx.fin.lines, ctx.n_act
    _title(ws, f"{fin.name} - Balance Sheet",
           "Actuals reported by Screener.in. Forecast cash comes from the Cash Flow sheet, so the balance check is a "
           "real check (no plug).", ctx.palette)
    _period_header(ws, ctx)
    R = ctx.rows[S_BS]
    c = ctx.col
    is_r, wc_r, ppe_r, cf_r = ctx.rows[S_IS], ctx.rows[S_WC], ctx.rows[S_PPE], ctx.rows[S_CF]
    w = lambda *a, **k: _write_row(ws, ctx, S_BS, *a, **k)
    flat = lambda key: (lambda i: f"={c(i - 1)}{R[key]}")
    _section(ws, R["h_le"], "Equity & liabilities", ctx.palette, ctx.n_all)
    w("equity_capital", "Equity share capital", L["equity_capital"], flat("equity_capital"))
    w("reserves", "Reserves", L["reserves"],
      lambda i: f"={c(i - 1)}{R['reserves']}+{q(S_IS)}!{c(i)}{is_r['net_profit']}-{q(S_IS)}!{c(i)}{is_r['dividends']}")
    te = lambda i: f"={c(i)}{R['equity_capital']}+{c(i)}{R['reserves']}"
    w("total_equity", "Total shareholders' equity", te, te, bold=True)
    w("minority_interest", "Minority interest", L["minority_interest"], flat("minority_interest"))
    w("lt_borrowings", "Long-term borrowings", L["lt_borrowings"], flat("lt_borrowings"), indent=1)
    w("st_borrowings", "Short-term borrowings", L["st_borrowings"], flat("st_borrowings"), indent=1)
    w("lease_liabilities", "Lease liabilities", L["lease_liabilities"], flat("lease_liabilities"), indent=1)
    w("other_borrowings", "Other borrowings", L["other_borrowings"], flat("other_borrowings"), indent=1)
    br = lambda i: (f"={c(i)}{R['lt_borrowings']}+{c(i)}{R['st_borrowings']}+{c(i)}{R['lease_liabilities']}"
                    f"+{c(i)}{R['other_borrowings']}")
    w("borrowings", "Total borrowings", br, br, bold=True)
    w("trade_payables", "Trade payables", lambda i: f"={q(S_WC)}!{c(i)}{wc_r['trade_payables']}",
      lambda i: f"={q(S_WC)}!{c(i)}{wc_r['trade_payables']}")
    w("other_liab", "Other liabilities & provisions", lambda i: f"={q(S_WC)}!{c(i)}{wc_r['other_liab']}",
      lambda i: f"={q(S_WC)}!{c(i)}{wc_r['other_liab']}")
    tl = lambda i: (f"={c(i)}{R['total_equity']}+{c(i)}{R['minority_interest']}+{c(i)}{R['borrowings']}"
                    f"+{c(i)}{R['trade_payables']}+{c(i)}{R['other_liab']}")
    w("total_liabilities", "TOTAL EQUITY & LIABILITIES", tl, tl, bold=True, top_border=True)
    _section(ws, R["h_a"], "Assets", ctx.palette, ctx.n_all)
    w("fixed_assets", "Net fixed assets", lambda i: f"={q(S_PPE)}!{c(i)}{ppe_r['closing_net']}",
      lambda i: f"={q(S_PPE)}!{c(i)}{ppe_r['closing_net']}")
    w("cwip", "Capital work in progress", lambda i: f"={q(S_PPE)}!{c(i)}{ppe_r['cwip']}",
      lambda i: f"={q(S_PPE)}!{c(i)}{ppe_r['cwip']}")
    w("investments", "Investments", L["investments"], flat("investments"))
    for key, lab in (("receivables", "Trade receivables"), ("inventory", "Inventories"),
                     ("other_ca", "Other operating assets"), ("loans_advances", "Loans & advances")):
        f = lambda i, key=key: f"={q(S_WC)}!{c(i)}{wc_r[key]}"
        w(key, lab, f, f)
    w("cash", "Cash & equivalents", L["cash"], lambda i: f"={q(S_CF)}!{c(i)}{cf_r['closing_cash']}")
    ta = lambda i: "=" + "+".join(f"{c(i)}{R[k]}" for k in ("fixed_assets", "cwip", "investments", "receivables",
                                                             "inventory", "cash", "other_ca", "loans_advances"))
    w("total_assets", "TOTAL ASSETS", ta, ta, bold=True, top_border=True)
    chk = lambda i: f"=ROUND({c(i)}{R['total_assets']}-{c(i)}{R['total_liabilities']},2)"
    w("balance_check", "Balance check (assets - liabilities) = 0", chk, chk, fmt='0.00;-0.00;0.00', bold=True)
    _section(ws, R["h_r"], "Returns & liquidity", ctx.palette, ctx.n_all)
    # Returns on average balances (prior and current year-end), as equity_agent.ratios: n/a in the first year.
    def on_avg(num, bal):
        def f(i):
            if i == 0:
                return "n/a"
            avg = f"({bal(i - 1)}+{bal(i)})/2"
            return f'=IF({avg}>0,{num(i)}/({avg}),"n/a")'
        return f

    np_ = lambda i: f"{q(S_IS)}!{c(i)}{is_r['net_profit']}"
    ebit_ = lambda i: f"({q(S_IS)}!{c(i)}{is_r['pbt']}+{q(S_IS)}!{c(i)}{is_r['interest']})"
    eq_ = lambda i: f"{c(i)}{R['total_equity']}"
    ta_ = lambda i: f"{c(i)}{R['total_assets']}"
    cap_ = lambda i: f"({c(i)}{R['total_equity']}+{c(i)}{R['borrowings']})"
    roa, roe, roce = on_avg(np_, ta_), on_avg(np_, eq_), on_avg(ebit_, cap_)
    de = lambda i: f"=IFERROR({c(i)}{R['borrowings']}/{c(i)}{R['total_equity']},0)"
    ic = lambda i: f"={c(i)}{R['fixed_assets']}+{c(i)}{R['cwip']}+{q(S_WC)}!{c(i)}{wc_r['nwc']}"
    ca = lambda i: (f"={c(i)}{R['receivables']}+{c(i)}{R['inventory']}+{c(i)}{R['cash']}+{c(i)}{R['other_ca']}"
                    f"+{c(i)}{R['loans_advances']}")
    cl = lambda i: f"={c(i)}{R['trade_payables']}+{c(i)}{R['st_borrowings']}+{c(i)}{R['other_liab']}"
    cr = lambda i: f"=IFERROR({c(i)}{R['current_assets']}/{c(i)}{R['current_liabilities']},0)"
    qr = lambda i: f"=IFERROR(({c(i)}{R['current_assets']}-{c(i)}{R['inventory']})/{c(i)}{R['current_liabilities']},0)"
    cashr = lambda i: f"=IFERROR({c(i)}{R['cash']}/{c(i)}{R['current_liabilities']},0)"
    w("roa", "Return on assets (net profit / average total assets)", roa, roa, fmt=F_PCT)
    w("roe", "Return on equity (net profit / average equity)", roe, roe, fmt=F_PCT)
    w("roce", "ROCE, pre-tax ((PBT + interest) / average (equity + borrowings))", roce, roce, fmt=F_PCT)
    w("debt_equity", "Debt / equity", de, de, fmt=F_X)
    w("invested_capital", "Operating invested capital (net block + CWIP + operating NWC)", ic, ic)
    w("current_assets", "Current assets (approx.: receivables + inventory + cash + other assets)", ca, ca)
    w("current_liabilities", "Current liabilities (approx.: payables + ST borrowings + other liabilities)", cl, cl)
    w("current_ratio", "Current ratio (approx.)", cr, cr, fmt=F_X)
    w("quick_ratio", "Quick ratio (approx.)", qr, qr, fmt=F_X)
    w("cash_ratio", "Cash ratio (approx.)", cashr, cashr, fmt=F_X)


# --------------------------------------------------------------------------- #
# CAPM & WACC
# --------------------------------------------------------------------------- #
WACC_ROWS = {"rf": 5, "beta": 6, "erp": 7, "ke": 8, "kd_pre": 9, "tax": 10, "kd_post": 11, "e": 12, "d": 13,
             "we": 14, "wd": 15, "wacc": 16, "g": 17}


def write_wacc(wb, ctx: Ctx):
    ws = wb.create_sheet(S_WACC)
    pal = ctx.palette
    _title(ws, f"{ctx.fin.name} - CAPM & WACC", "Market-value weights: market cap and latest reported borrowings.", pal)
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["C"].width = 14
    ws.column_dimensions["D"].width = 90
    D, R, n = ctx.drv, WACC_ROWS, ctx.n_act
    last = ctx.col(n - 1)
    bs_b = ctx.r(S_BS, "borrowings")
    v = ctx.val
    rows = [
        ("rf", "Risk-free rate", f"={D['rf']}", F_PCT, "Drivers sheet"),
        ("beta", "Beta (Blume-adjusted)", f"={D['beta']}", '0.00',
         v.beta_source + (f"; raw {v.beta_raw:.2f}" if v.beta_raw is not None else "")),
        ("erp", "Equity risk premium", f"={D['erp']}", F_PCT, "Drivers sheet"),
        ("ke", "Cost of equity (Ke = Rf + Beta x ERP)", "=C5+C6*C7", F_PCT, ""),
        ("kd_pre", "Pre-tax cost of debt", f"={D['interest_rate']}", F_PCT, "Interest rate driver"),
        ("tax", "Tax rate", f"={D['tax_rate']}", F_PCT, "Tax rate driver"),
        ("kd_post", "Post-tax cost of debt", "=C9*(1-C10)", F_PCT, ""),
        ("e", "Equity value (market cap, ₹ Cr)", f"={D['mcap']}", F_NUM, ""),
        ("d", "Debt (borrowings incl. leases, ₹ Cr)", f"={q(S_BS)}!${last}${bs_b}", F_NUM, f"FY{ctx.fin.years[-1] % 100:02d} balance sheet"),
        ("we", "Weight of equity", "=C12/(C12+C13)", F_PCT, ""),
        ("wd", "Weight of debt", "=C13/(C12+C13)", F_PCT, ""),
        ("wacc", "WACC", "=C8*C14+C11*C15", F_PCT, ""),
        ("g", "Terminal growth (capped at WACC - 1%)", f"=MIN({D['g']},C16-0.01)", F_PCT, ""),
    ]
    for key, label, f, fmt, note in rows:
        r = R[key]
        ws.cell(r, 1, label).font = Font(bold=key in ("ke", "wacc"))
        cell = ws.cell(r, 3, f)
        cell.number_format = fmt
        cell.font = Font(bold=key in ("ke", "wacc"))
        ws.cell(r, 4, note).font = Font(size=9, color=GREY)
    ws.sheet_view.showGridLines = False


# --------------------------------------------------------------------------- #
# DCF (3-stage)
# --------------------------------------------------------------------------- #
DCF_LINE_KEYS = ["revenue", "growth", "ebitda", "dep", "ebit", "ebit_margin", "tax_rate", "nopat", "add_dep", "capex",
                 "dnwc", "roic", "reinvestment", "fcff", "period", "df", "pv"]
DCF_ROWS = {k: 5 + i for i, k in enumerate(DCF_LINE_KEYS)}
_SUM_KEYS = ["wacc", "g", "ic_prev", "ic_last", "roic_raw", "roic_cap", "roic", "ronic_share", "ronic_ovr", "ronic",
             "sum_pv_explicit", "sum_pv_fade", "sum_pv", "tv", "pv_tv", "ev", "tv_share", "cash", "inv", "debt",
             "minority", "equity", "shares", "dcf_value", "ke", "div_yield", "dcf_12m", "tv_nopat_mult", None,
             "exit_mult", "exit_value"]
DCF_ROWS.update({k: 24 + i for i, k in enumerate(_SUM_KEYS) if k})
DCF_ROWS["sens_head"] = DCF_ROWS["exit_value"] + 4
HELPER_COL0 = 14  # column N: sensitivity helper area


def write_dcf(wb, ctx: Ctx):
    ws = wb.create_sheet(S_DCF)
    pal, fin, n = ctx.palette, ctx.fin, ctx.n_act
    R, D = DCF_ROWS, ctx.drv
    nf = len(ctx.forecast.years)
    nfade = int(ctx.cfg.get("fade_years", 5))
    ntot = nf + nfade
    _title(ws, f"{fin.name} - DCF Valuation (3-stage FCFF, mid-year discounting)",
           "Explicit years: FCFF = EBIT x (1 - t) + D&A - capex - change in NWC (EBIT excludes other income; cash and "
           "investments are added in the bridge). Fade years: growth falls linearly to g, ROIC moves to RONIC, "
           "reinvestment = NOPAT x growth / ROIC. Terminal value = NOPAT x (1 + g) x (1 - g / RONIC) / (WACC - g).", pal)
    ws.column_dimensions["A"].width = 46
    for j in range(ntot + 1):
        ws.column_dimensions[get_column_letter(2 + j)].width = 12
    ws.cell(3, 1, "₹ Cr").font = Font(size=9, italic=True, color=GREY)
    is_r, ppe_r, wc_r, bs_r = ctx.rows[S_IS], ctx.rows[S_PPE], ctx.rows[S_WC], ctx.rows[S_BS]
    first, lastx = "B", get_column_letter(1 + nf)
    lastf = get_column_letter(1 + ntot)
    g_ref, w_ref = f"$C${R['g']}", f"$C${R['wacc']}"
    for j in range(ntot):
        c = get_column_letter(2 + j)
        explicit = j < nf
        h = ws[f"{c}3"]
        h.value = ctx.forecast.year_labels[j] if explicit else f"FY{ctx.val.years[j] % 100:02d}F"
        h.font = Font(bold=True, color=pal.text_on_light if explicit else "FFFFFF")
        h.fill = _fill(pal.light if explicit else "7F7F7F")
        h.alignment = Alignment(horizontal="center")
        ws[f"{c}4"] = "Explicit" if explicit else "Fade"
        ws[f"{c}4"].font = Font(size=8, color=GREY)
        ws[f"{c}4"].alignment = Alignment(horizontal="center")
        p = get_column_letter(1 + j)
        if explicit:
            src = ctx.col(n + j)
            cells = {
                "revenue": f"={q(S_IS)}!{src}{is_r['sales']}",
                "growth": f"={q(S_IS)}!{src}{is_r['growth']}",
                "ebitda": f"={q(S_IS)}!{src}{is_r['ebitda']}",
                "dep": f"={q(S_IS)}!{src}{is_r['depreciation']}",
                "ebit": f"={c}{R['ebitda']}-{c}{R['dep']}",
                "ebit_margin": f"=IFERROR({c}{R['ebit']}/{c}{R['revenue']},0)",
                "tax_rate": f"={ctx.dy('tax_rate', n + j)}",
                "nopat": f"={c}{R['ebit']}*(1-{c}{R['tax_rate']})",
                "add_dep": f"={c}{R['dep']}",
                "capex": f"={q(S_PPE)}!{src}{ppe_r['capex']}",
                "dnwc": f"={q(S_WC)}!{src}{wc_r['delta_nwc']}",
                "reinvestment": f"={c}{R['capex']}-{c}{R['dep']}+{c}{R['dnwc']}",
                "fcff": f"={c}{R['nopat']}+{c}{R['add_dep']}-{c}{R['capex']}-{c}{R['dnwc']}",
            }
            if j == nf - 1:
                cells["roic"] = f"=$C${R['roic']}"
        else:
            k = j - nf + 1
            cells = {
                "growth": f"=${lastx}${R['growth']}+({g_ref}-${lastx}${R['growth']})*{k}/{nfade}",
                "revenue": f"={p}{R['revenue']}*(1+{c}{R['growth']})",
                "ebit_margin": f"=${lastx}${R['ebit_margin']}",
                "ebit": f"={c}{R['revenue']}*{c}{R['ebit_margin']}",
                "tax_rate": f"=${lastx}${R['tax_rate']}",
                "nopat": f"={c}{R['ebit']}*(1-{c}{R['tax_rate']})",
                "roic": f"=$C${R['roic']}+($C${R['ronic']}-$C${R['roic']})*{k}/{nfade}",
                "reinvestment": f"={c}{R['nopat']}*{c}{R['growth']}/{c}{R['roic']}",
                "fcff": f"={c}{R['nopat']}-{c}{R['reinvestment']}",
            }
        cells["period"] = j + 0.5
        cells["df"] = f"=1/(1+{w_ref})^{c}{R['period']}"
        cells["pv"] = f"={c}{R['fcff']}*{c}{R['df']}"
        for key, v in cells.items():
            cell = ws.cell(R[key], 2 + j, v)
            fmt = {"tax_rate": F_PCT, "growth": F_PCT, "ebit_margin": F_PCT, "roic": F_PCT, "df": '0.0000',
                   "period": '0.0'}.get(key, F_NUM)
            cell.number_format = fmt
            if key in ("fcff", "pv"):
                cell.font = Font(bold=True)
    labels = {"revenue": "Revenue", "growth": "  Revenue growth", "ebitda": "EBITDA", "dep": "Depreciation & amortisation",
              "ebit": "EBIT (operating, excl. other income)", "ebit_margin": "  EBIT margin", "tax_rate": "Tax rate",
              "nopat": "NOPAT = EBIT x (1 - t)", "add_dep": "Add: D&A", "capex": "Less: capex",
              "dnwc": "Less: increase in net working capital", "roic": "ROIC (fade: towards RONIC)",
              "reinvestment": "Net reinvestment (capex - D&A + change in NWC; fade: NOPAT x g / ROIC)",
              "fcff": "Free cash flow to firm (FCFF)", "period": "Discount period (mid-year)", "df": "Discount factor",
              "pv": "Present value of FCFF"}
    for key, lab in labels.items():
        ws.cell(R[key], 1, lab).font = Font(bold=key in ("fcff", "pv"))
    last = ctx.col(n - 1)
    lastx_ic = ctx.col(n + nf - 1)
    prevx_ic = ctx.col(n + nf - 2)
    bic = ctx.r(S_BS, "invested_capital")
    avg_ic = f"AVERAGE(C{R['ic_prev']},C{R['ic_last']})"
    summary = [
        ("wacc", "WACC", f"={q(S_WACC)}!$C${WACC_ROWS['wacc']}", F_PCT),
        ("g", "Terminal growth", f"={q(S_WACC)}!$C${WACC_ROWS['g']}", F_PCT),
        ("ic_prev", f"Operating invested capital {ctx.forecast.year_labels[-2]}", f"={q(S_BS)}!${prevx_ic}${bic}", F_NUM),
        ("ic_last", f"Operating invested capital {ctx.forecast.year_labels[-1]}", f"={q(S_BS)}!${lastx_ic}${bic}", F_NUM),
        ("roic_raw", "Current ROIC (final explicit NOPAT / average invested capital)",
         f'=IF({avg_ic}>0,{lastx}{R["nopat"]}/{avg_ic},"n/a")', F_PCT),
        ("roic_cap", "ROIC cap", f"={D['roic_cap']}", '0%'),
        ("roic", "ROIC used (capped; floored at g + 1%)",
         f"=IF({avg_ic}>0,MIN(C{R['roic_cap']},MAX(C{R['roic_raw']},C{R['g']}+0.01)),C{R['roic_cap']})", F_PCT),
        ("ronic_share", "Share of ROIC premium kept (active scenario)", f"={D['ronic_share']}", '0%'),
        ("ronic_ovr", "Terminal RONIC override (Drivers)", f'=IF(ISNUMBER({D["ronic_ovr"]}),{D["ronic_ovr"]},"none")', F_PCT),
        ("ronic", "Terminal RONIC",
         f"=MAX(C{R['g']}+0.01,IF(ISNUMBER(C{R['ronic_ovr']}),C{R['ronic_ovr']},MAX(C{R['wacc']},MIN(C{R['roic']},"
         f"C{R['wacc']}+C{R['ronic_share']}*(C{R['roic']}-C{R['wacc']})))))", F_PCT),
        ("sum_pv_explicit", "PV of FCFF, explicit years", f"=SUM({first}{R['pv']}:{lastx}{R['pv']})", F_NUM),
        ("sum_pv_fade", "PV of FCFF, fade years", f"=SUM({get_column_letter(2 + nf)}{R['pv']}:{lastf}{R['pv']})", F_NUM),
        ("sum_pv", "Sum of PV of FCFF", f"=C{R['sum_pv_explicit']}+C{R['sum_pv_fade']}", F_NUM),
        ("tv", "Terminal value = NOPAT x (1+g) x (1 - g/RONIC) / (WACC - g)",
         f"={lastf}{R['nopat']}*(1+C{R['g']})*(1-C{R['g']}/C{R['ronic']})/(C{R['wacc']}-C{R['g']})", F_NUM),
        ("pv_tv", "PV of terminal value (final-year mid-year factor)", f"=C{R['tv']}*{lastf}{R['df']}", F_NUM),
        ("ev", "Enterprise value", f"=C{R['sum_pv']}+C{R['pv_tv']}", F_NUM),
        ("tv_share", "Terminal value share of EV", f"=C{R['pv_tv']}/C{R['ev']}", '0%'),
        ("cash", "Add: cash & equivalents", f"={q(S_BS)}!${last}${bs_r['cash']}", F_NUM),
        ("inv", "Add: investments", f"={q(S_BS)}!${last}${bs_r['investments']}", F_NUM),
        ("debt", "Less: borrowings (incl. leases)", f"={q(S_BS)}!${last}${bs_r['borrowings']}", F_NUM),
        ("minority", "Less: minority interest", f"={q(S_BS)}!${last}${bs_r['minority_interest']}", F_NUM),
        ("equity", "Equity value", f"=C{R['ev']}+C{R['cash']}+C{R['inv']}-C{R['debt']}-C{R['minority']}", F_NUM),
        ("shares", "Shares outstanding (Cr)", f"={D['shares']}", '#,##0.00'),
        ("dcf_value", "DCF value per share today (₹)", f"=C{R['equity']}/C{R['shares']}", F_RS),
        ("ke", "Cost of equity", f"={q(S_WACC)}!$C${WACC_ROWS['ke']}", F_PCT),
        ("div_yield", "Dividend yield", f"={D['div_yield']}", '0.00%'),
        ("dcf_12m", "DCF value in 12 months (₹) = value x (1 + Ke - dividend yield)",
         f"=C{R['dcf_value']}*(1+C{R['ke']}-C{R['div_yield']})", F_RS),
        ("tv_nopat_mult", "Implied terminal multiple of NOPAT (x)",
         f"=C{R['tv']}/({lastf}{R['nopat']}*(1+C{R['g']}))", F_X1),
        ("exit_mult", "Cross-check: current EV / EBITDA (x)",
         f"=IFERROR(({D['mcap']}+C{R['debt']}+C{R['minority']}-C{R['cash']}-C{R['inv']})/{q(S_IS)}!${last}${is_r['ebitda']},\"n/a\")", F_X1),
        ("exit_value", "Cross-check: exit-multiple value per share (₹)",
         f"=IFERROR((C{R['sum_pv_explicit']}+{lastx}{R['ebitda']}*C{R['exit_mult']}/(1+C{R['wacc']})^{nf}+C{R['cash']}"
         f"+C{R['inv']}-C{R['debt']}-C{R['minority']})/C{R['shares']},\"n/a\")", F_RS),
    ]
    for key, lab, f, fmt in summary:
        r = R[key]
        bold = key in ("ev", "equity", "dcf_value", "dcf_12m")
        ws.cell(r, 1, lab).font = Font(bold=bold)
        cell = ws.cell(r, 3, f)
        if fmt:
            cell.number_format = fmt
        cell.font = Font(bold=bold)
        if key in ("dcf_value", "dcf_12m"):
            cell.fill = _fill(pal.light)
            ws.cell(r, 1).fill = _fill(pal.light)
    _write_sensitivity(ws, ctx, nf, nfade)
    ws.sheet_view.showGridLines = False
    ws.freeze_panes = "B5"


def _write_sensitivity(ws, ctx: Ctx, nf: int, nfade: int):
    """WACC x g grid. Each cell re-runs the fade and terminal stages (explicit FCFF fixed) via a helper row."""
    R, D, pal = DCF_ROWS, ctx.drv, ctx.palette
    h = R["sens_head"]
    ws.cell(h - 1, 1, "Sensitivity: DCF value per share today (₹) - WACC (rows) vs terminal growth (columns)").font = \
        Font(bold=True, color=pal.dark)
    ws.cell(h, 1, "WACC \\ g").font = Font(bold=True)
    lastx = get_column_letter(1 + nf)
    for j, s in enumerate(G_STEPS):
        cell = ws.cell(h, 2 + j, f"=$C${R['g']}+({s})")
        cell.number_format = F_PCT
        cell.font = Font(bold=True, color=pal.text_on_dark)
        cell.fill = _fill(pal.dark)
    bridge = f"$C${R['cash']}+$C${R['inv']}-$C${R['debt']}-$C${R['minority']}"
    avg_ic = f"AVERAGE($C${R['ic_prev']},$C${R['ic_last']})"
    H = HELPER_COL0
    hc = lambda k: get_column_letter(H + k)
    ws.cell(h - 1, H, "Sensitivity helper (one row per grid cell): WACC, g, ROIC, RONIC, then growth / revenue / FCFF "
                      "for each fade year, then value per share").font = Font(size=8, color=GREY)
    for i, s in enumerate(WACC_STEPS):
        r = h + 1 + i
        wc = ws.cell(r, 1, f"=$C${R['wacc']}+({s})")
        wc.number_format = F_PCT
        wc.font = Font(bold=True)
        for j in range(len(G_STEPS)):
            hr = h + 1 + i * len(G_STEPS) + j
            w_, g_ = f"${hc(0)}{hr}", f"${hc(1)}{hr}"
            ws.cell(hr, H, f"=$A{r}")
            ws.cell(hr, H + 1, f"={get_column_letter(2 + j)}${h}")
            ws.cell(hr, H + 2, f"=IF({avg_ic}>0,MIN($C${R['roic_cap']},MAX($C${R['roic_raw']},{g_}+0.01)),$C${R['roic_cap']})")
            roic_ = f"${hc(2)}{hr}"
            ws.cell(hr, H + 3, f"=MAX({g_}+0.01,IF(ISNUMBER($C${R['ronic_ovr']}),$C${R['ronic_ovr']},MAX({w_},MIN({roic_},"
                               f"{w_}+$C${R['ronic_share']}*({roic_}-{w_})))))")
            ronic_ = f"${hc(3)}{hr}"
            pv_terms = []
            for k in range(1, nfade + 1):
                cg, crv, cf = H + 4 + 3 * (k - 1), H + 5 + 3 * (k - 1), H + 6 + 3 * (k - 1)
                ws.cell(hr, cg, f"=${lastx}${R['growth']}+({g_}-${lastx}${R['growth']})*{k}/{nfade}")
                prev_rev = f"${lastx}${R['revenue']}" if k == 1 else f"{hc(crv - H - 3)}{hr}"
                ws.cell(hr, crv, f"={prev_rev}*(1+{get_column_letter(cg)}{hr})")
                roic_k = f"({roic_}+({ronic_}-{roic_})*{k}/{nfade})"
                ws.cell(hr, cf, f"={get_column_letter(crv)}{hr}*${lastx}${R['ebit_margin']}*(1-${lastx}${R['tax_rate']})"
                                f"*(1-{get_column_letter(cg)}{hr}/{roic_k})")
                pv_terms.append(f"{get_column_letter(cf)}{hr}/(1+{w_})^{nf + k - 0.5}")
            rev_last = f"{get_column_letter(H + 5 + 3 * (nfade - 1))}{hr}"
            nopat_last = f"{rev_last}*${lastx}${R['ebit_margin']}*(1-${lastx}${R['tax_rate']})"
            tv = f"{nopat_last}*(1+{g_})*(1-{g_}/{ronic_})/({w_}-{g_})/(1+{w_})^{nf + nfade - 0.5}"
            explicit = (f"SUMPRODUCT($B${R['fcff']}:${lastx}${R['fcff']},"
                        f"1/(1+{w_})^$B${R['period']}:${lastx}${R['period']})")
            vcol = H + 4 + 3 * nfade
            ws.cell(hr, vcol, f'=IF({w_}-{g_}<0.005,"n/a",({explicit}+{"+".join(pv_terms)}+{tv}+{bridge})/$C${R["shares"]})')
            cell = ws.cell(r, 2 + j, f"={get_column_letter(vcol)}{hr}")
            cell.number_format = '#,##0'
            if i == 2 and j == 2:
                cell.font = Font(bold=True)
                cell.fill = _fill(pal.light)
    for k in range(4 + 3 * nfade + 1):
        ws.column_dimensions[get_column_letter(H + k)].hidden = True


# --------------------------------------------------------------------------- #
# Peers
# --------------------------------------------------------------------------- #
PEER_COLS = [("name", "Company", None), ("symbol", "Symbol", None), ("price", "Price (₹)", F_EPS),
             ("market_cap", "Market cap (₹ Cr)", F_NUM), ("pe", "P/E (x)", F_X1), ("pb", "P/B (x)", F_X1),
             ("ev", "EV (₹ Cr)", F_NUM), ("ev_ebitda", "EV/EBITDA (x)", F_X1), ("roe", "ROE % (latest FY)", '0.0'),
             ("roce", "ROCE % (latest FY)", '0.0'), ("sales_growth_3y", "Sales CAGR 3y", F_PCT),
             ("ebitda_margin", "EBITDA margin", F_PCT), ("net_margin", "Net margin", F_PCT),
             ("div_yield", "Div. yield %", '0.00'), ("basis", "Basis", None)]


def write_peers(wb, ctx: Ctx):
    from .peers import EV_EBITDA_RANGE, subject_snapshot

    ws = wb.create_sheet(S_PEERS)
    pal = ctx.palette
    ps = ctx.peers
    src = ps.source if ps is not None else "No peer data (offline run without a cached peer snapshot)"
    _title(ws, f"{ctx.fin.name} - Peer Comparison",
           f"{src}. EV = market cap + borrowings - cash - investments; multiples on TTM figures where Screener shows "
           "them. Blue cells are Screener values.", pal)
    heads = [h for _, h, _ in PEER_COLS] + ["In EV/EBITDA median"]
    _head_row(ws, 4, heads, pal)
    ws.column_dimensions["A"].width = 30
    for j in range(2, len(heads) + 1):
        ws.column_dimensions[get_column_letter(j)].width = 13
    rows = [subject_snapshot(ctx.fin)] + (list(ps.peers) if ps is not None else [])
    r0 = 5
    evc = get_column_letter(1 + [k for k, _, _ in PEER_COLS].index("ev_ebitda"))
    helper_col = len(PEER_COLS) + 1
    lo, hi = EV_EBITDA_RANGE
    for i, p in enumerate(rows):
        r = r0 + i
        for j, (key, _, fmt) in enumerate(PEER_COLS, 1):
            v = getattr(p, key)
            if v is None:
                continue
            c = _put(ws, r, j, v, fmt)
            if i == 0:
                c.font = Font(bold=True, color=BLUE if not isinstance(v, str) else BLACK)
        if i > 0:
            ws.cell(r, helper_col, f'=IF(AND(ISNUMBER({evc}{r}),{evc}{r}>{lo},{evc}{r}<{hi}),{evc}{r},"")').number_format = F_X1
    last = r0 + len(rows) - 1
    rm = last + 2
    ws.cell(rm, 1, "Peer median (excl. subject)").font = Font(bold=True)
    hcol = get_column_letter(helper_col)
    for j, (key, _, fmt) in enumerate(PEER_COLS, 1):
        if fmt is None or key in ("price", "market_cap", "ev"):
            continue
        col = get_column_letter(j)
        if key == "ev_ebitda":
            f = (f'=IF(COUNT({hcol}{r0 + 1}:{hcol}{max(last, r0 + 1)})>=2,MEDIAN({hcol}{r0 + 1}:{hcol}{max(last, r0 + 1)}),"n/a")'
                 if len(rows) > 1 else '="n/a"')
        else:
            f = (f'=IF(COUNT({col}{r0 + 1}:{col}{last})>=1,MEDIAN({col}{r0 + 1}:{col}{last}),"n/a")'
                 if len(rows) > 1 else '="n/a"')
        c = ws.cell(rm, j, f)
        c.number_format = fmt
        c.font = Font(bold=True)
    ctx.refs["peer_ev_ebitda"] = f"{q(S_PEERS)}!${evc}${rm}"
    ws.cell(rm + 1, 1, f"EV/EBITDA median uses values between {lo:.0f}x and {hi:.0f}x and needs at least two peers.") \
        .font = Font(size=9, color=GREY)
    nxt = rm + 3
    if ps is not None and ps.warnings:
        for k, wtxt in enumerate(ps.warnings):
            ws.cell(rm + 3 + k, 1, wtxt).font = Font(size=9, color=GREY)
        nxt += len(ps.warnings)
    if ps is not None and ps.has_history:
        _write_peer_history(ws, ctx, rows, nxt + 1)
    ws.freeze_panes = "B5"
    ws.sheet_view.showGridLines = False


PEER_HISTORY_LABELS = {"sales_growth": "Revenue growth", "ebitda_margin": "EBITDA margin", "net_margin": "Net margin",
                       "roe": "ROE (net profit / average equity)",
                       "roce": "ROCE ((PBT + interest) / average (equity + borrowings))"}


def _write_peer_history(ws, ctx: Ctx, companies, r: int):
    """Five fiscal years of each ratio: reported values per company, live peer median, company gap to the median."""
    from .peers import HISTORY_METRICS

    pal, ps = ctx.palette, ctx.peers
    years = ps.history_years()
    if not years:
        return
    ny = len(years)
    width = 2 + ny + 1
    _section(ws, r, "Multi-year ratios vs peers (fiscal years; one set of definitions for every company)", pal, width)
    others = sorted({p.year_end for p in ps.peers if p.year_end and p.year_end != companies[0].year_end})
    note = ("Computed from each company's Screener annual P&L and balance sheet. Average balances use the opening and "
            "closing year. Peer median needs at least two peers; gap = company minus median.")
    if others:
        names = ", ".join(p.symbol for p in ps.peers if p.year_end in others)
        note += f" Different fiscal year-end: {names}."
    ws.cell(r + 1, 1, note).font = Font(size=9, color=GREY)
    r += 3
    ctx.refs["peer_history"] = {}
    for metric in HISTORY_METRICS:
        _head_row(ws, r, [PEER_HISTORY_LABELS[metric], "Symbol"] + years + ["Years above median"], pal)
        r0 = r + 1
        for i, p in enumerate(companies):
            rr = r0 + i
            ws.cell(rr, 1, p.name).font = Font(bold=i == 0)
            ws.cell(rr, 2, p.symbol).font = Font(bold=i == 0)
            hist = p.history.get(metric, {})
            for j, y in enumerate(years):
                v = hist.get(y)
                if v is not None:
                    _put(ws, rr, 3 + j, v, F_PCT, bold=i == 0)
        last = r0 + len(companies) - 1
        rm, rg = last + 1, last + 2
        ws.cell(rm, 1, "Peer median (excl. subject)").font = Font(bold=True)
        ws.cell(rg, 1, "Company minus median").font = Font(bold=True, color=pal.dark)
        refs = {}
        for j, y in enumerate(years):
            col = get_column_letter(3 + j)
            rng = f"{col}{r0 + 1}:{col}{last}"
            c = ws.cell(rm, 3 + j, f'=IF(COUNT({rng})>=2,MEDIAN({rng}),"n/a")' if last > r0 else '="n/a"')
            c.number_format, c.font = F_PCT, Font(bold=True)
            g = ws.cell(rg, 3 + j, f'=IF(AND(ISNUMBER({col}{r0}),ISNUMBER({col}{rm})),{col}{r0}-{col}{rm},"n/a")')
            g.number_format, g.font = '+0.0%;-0.0%;0.0%', Font(bold=True, color=pal.dark)
            refs[y] = (f"{col}{rm}", f"{col}{rg}")
        c1, c2 = get_column_letter(3), get_column_letter(2 + ny)
        ab = ws.cell(rg, 3 + ny, f'=COUNTIF({c1}{rg}:{c2}{rg},">0")&" of "&COUNT({c1}{rg}:{c2}{rg})')
        ab.font, ab.alignment = Font(bold=True), Alignment(horizontal="center")
        for k in range(r0, rg + 1):
            ws.cell(k, 1).border = Border(bottom=THIN)
        ctx.refs["peer_history"][metric] = refs
        r = rg + 2


# --------------------------------------------------------------------------- #
# Valuation summary (methods -> target -> rating)
# --------------------------------------------------------------------------- #
def _rating_formula(cell: str, cfg: dict) -> str:
    expr = f'"{cfg.get("rating_floor", "SELL")}"'
    for threshold, label in reversed(cfg["rating_thresholds"]):
        expr = f'IF({cell}>={threshold},"{label}",{expr})'
    return "=" + expr


def write_valuation_summary(wb, ctx: Ctx):
    ws = wb.create_sheet(S_VAL)
    pal, fin, v, n, D = ctx.palette, ctx.fin, ctx.val, ctx.n_act, ctx.drv
    _title(ws, f"{fin.name} - Valuation Summary (12-month target)",
           "Target = weighted blend of DCF (rolled forward 12 months), P/E vs own 5-year history and EV/EBITDA vs peer "
           "median, on next-twelve-months (NTM) earnings; weights renormalise over methods that are available. A sourced "
           "sum-of-the-parts replaces the blend.", pal)
    ws.column_dimensions["A"].width = 52
    for col in "BCDEFGH":
        ws.column_dimensions[col].width = 13
    ws.column_dimensions["I"].width = 70
    is_r = ctx.rows[S_IS]
    R: Dict[str, int] = {}
    r = 4
    _section(ws, r, "1. P/E vs own history (fiscal-year-end, actual traded price, that year's share count)", pal, 8)
    r += 1
    hist = list(v.pe_history.items())
    _head_row(ws, r, ["", *[k for k, _ in hist]], pal)
    R["pe_head"] = r
    labels_idx = {f"FY{y % 100:02d}": k for k, y in enumerate(fin.years)}
    md = ctx.md
    prices = (md.month_end_prices_unadj or md.month_end_prices) if md is not None else {}
    from .valuation import _MONTH

    rows = {"price": "Month-end price at fiscal year end (₹)", "np": "Net profit (₹ Cr)", "shares": "Share count (Cr)",
            "pe": "P/E (x)"}
    for j, key in enumerate(rows):
        R[f"pe_{key}"] = r + 1 + j
        ws.cell(r + 1 + j, 1, rows[key])
    for j, (lab, pe) in enumerate(hist, 2):
        k = labels_idx[lab]
        mon = _MONTH.get(fin.period_labels[k][:3])
        price = prices.get(f"{fin.years[k]}-{mon:02d}")
        npf = fin.lines["net_profit"][k]
        shares = npf * pe / price if price else None  # share count the engine used for that year
        cl = get_column_letter(j)
        _put(ws, R["pe_price"], j, price, F_EPS)
        _put(ws, R["pe_np"], j, f"={q(S_IS)}!{ctx.col(k)}{is_r['net_profit']}", F_NUM)
        _put(ws, R["pe_shares"], j, shares, '#,##0.00')
        _put(ws, R["pe_pe"], j, f"={cl}{R['pe_price']}/({cl}{R['pe_np']}/{cl}{R['pe_shares']})", F_X1, bold=True)
    r = R["pe_pe"] + 1
    rng = f"B{R['pe_pe']}:{get_column_letter(max(2, len(hist) + 1))}{R['pe_pe']}"
    items = [
        ("pe_fallback", "Current trailing P/E (Screener; used if fewer than 2 history points)", fin.pe, F_X1, None),
        ("pe_mean", "Mean P/E (x)", f'=IF(COUNT({rng})>=2,AVERAGE({rng}),IF(ISNUMBER(C{r}),C{r},"n/a"))', F_X1, None),
        ("pe_sd", "Standard deviation (x)", f'=IF(COUNT({rng})>=2,STDEV({rng}),"n/a")', F_X1, None),
    ]
    first_fc, second_fc = ctx.col(n), ctx.col(n + 1)
    shares = D["shares"]
    w_ = D["ntm_w"]
    items += [
        ("ntm_eps", f"NTM EPS (₹) = weight x {ctx.forecast.year_labels[0]} + (1 - weight) x {ctx.forecast.year_labels[1]}",
         f"=({w_}*{q(S_IS)}!{first_fc}{is_r['net_profit']}+(1-{w_})*{q(S_IS)}!{second_fc}{is_r['net_profit']})/{shares}",
         F_EPS, None),
        ("pe_value", "P/E value per share (₹) = mean P/E x NTM EPS", None, F_RS, None),
        ("pe_low", "P/E band low (mean - 1 sd)", None, F_RS, None),
        ("pe_high", "P/E band high (mean + 1 sd)", None, F_RS, None),
    ]
    for key, lab, f, fmt, _ in items:
        R[key] = r
        ws.cell(r, 1, lab)
        r += 1
    _put(ws, R["pe_fallback"], 3, fin.pe, F_X1)
    _put(ws, R["pe_mean"], 3, f'=IF(COUNT({rng})>=2,AVERAGE({rng}),IF(ISNUMBER(C{R["pe_fallback"]}),C{R["pe_fallback"]},"n/a"))',
         F_X1, bold=True)
    _put(ws, R["pe_sd"], 3, f'=IF(COUNT({rng})>=2,STDEV({rng}),"n/a")', F_X1)
    _put(ws, R["ntm_eps"], 3, items[3][2], F_EPS)
    pe_ok = f'AND(ISNUMBER(C{R["pe_mean"]}),C{R["ntm_eps"]}>0)'
    _put(ws, R["pe_value"], 3, f'=IF({pe_ok},C{R["pe_mean"]}*C{R["ntm_eps"]},"n/a")', F_RS, bold=True)
    _put(ws, R["pe_low"], 3, f'=IF(AND({pe_ok},ISNUMBER(C{R["pe_sd"]})),(C{R["pe_mean"]}-C{R["pe_sd"]})*C{R["ntm_eps"]},"n/a")', F_RS)
    _put(ws, R["pe_high"], 3, f'=IF(AND({pe_ok},ISNUMBER(C{R["pe_sd"]})),(C{R["pe_mean"]}+C{R["pe_sd"]})*C{R["ntm_eps"]},"n/a")', F_RS)
    ws.cell(R["pe_mean"], 9, v.pe_source).font = Font(size=9, color=GREY)

    r += 1
    _section(ws, r, "2. EV/EBITDA vs peer median", pal, 8)
    r += 1
    dcf = DCF_ROWS
    bridge = (f"{q(S_DCF)}!$C${dcf['cash']}+{q(S_DCF)}!$C${dcf['inv']}-{q(S_DCF)}!$C${dcf['debt']}"
              f"-{q(S_DCF)}!$C${dcf['minority']}")
    ev_rows = [
        ("peer_mult", "Peer median EV/EBITDA (x)", f"={ctx.refs['peer_ev_ebitda']}", F_X1),
        ("ntm_ebitda", "NTM EBITDA (₹ Cr)",
         f"={w_}*{q(S_IS)}!{first_fc}{is_r['ebitda']}+(1-{w_})*{q(S_IS)}!{second_fc}{is_r['ebitda']}", F_NUM),
        ("bridge", "Cash + investments - borrowings - minority (₹ Cr)", f"={bridge}", F_NUM),
        ("ev_value", "EV/EBITDA value per share (₹)", None, F_RS),
    ]
    for key, lab, f, fmt in ev_rows:
        R[key] = r
        ws.cell(r, 1, lab)
        if f:
            _put(ws, r, 3, f, fmt)
        r += 1
    _put(ws, R["ev_value"], 3, f'=IF(AND(ISNUMBER(C{R["peer_mult"]}),C{R["ntm_ebitda"]}>0),'
                               f'(C{R["peer_mult"]}*C{R["ntm_ebitda"]}+C{R["bridge"]})/{shares},"n/a")', F_RS, bold=True)
    ws.cell(R["peer_mult"], 9, v.ev_source).font = Font(size=9, color=GREY)

    r += 1
    _section(ws, r, "3. Target price", pal, 8)
    r += 1
    _head_row(ws, r, ["Method", "Value (₹)", "Weight (input)", "Weight used", "Contribution"], pal)
    r += 1
    methods = [("dcf", "DCF, rolled forward 12 months", f"={q(S_DCF)}!$C${dcf['dcf_12m']}", D["w_dcf"]),
               ("pe", "P/E vs own history", f"=C{R['pe_value']}", D["w_pe"]),
               ("ev", "EV/EBITDA vs peers", f"=C{R['ev_value']}", D["w_ev"])]
    mrows = {}
    for key, lab, f, wref in methods:
        mrows[key] = r
        ws.cell(r, 1, lab)
        _put(ws, r, 2, f, F_RS)
        _put(ws, r, 3, f"={wref}", '0%')
        r += 1
    avail = "+".join(f"IF(ISNUMBER(B{mrows[k]}),C{mrows[k]},0)" for k in mrows)
    for key in mrows:
        rr = mrows[key]
        _put(ws, rr, 4, f"=IF(ISNUMBER(B{rr}),C{rr}/({avail}),0)", '0%')
        _put(ws, rr, 5, f"=IF(ISNUMBER(B{rr}),B{rr}*D{rr},0)", F_RS)
    R["blend"] = r
    ws.cell(r, 1, "Blended value (₹)").font = Font(bold=True)
    _put(ws, r, 5, f"=SUM(E{mrows['dcf']}:E{mrows['ev']})", F_RS, bold=True)
    r += 2
    ai = ctx.analyst
    sotp_rows = v.sotp
    R["sotp_head"] = r
    _section(ws, r, "4. Sum of the parts (analyst input; replaces the blend when present)", pal, 8)
    r += 1
    if sotp_rows:
        _head_row(ws, r, ["Segment", "Metric (₹ Cr)", "Multiple (x)", "EV (₹ Cr)", "Method", "", "", "", "Source"], pal)
        r += 1
        s0 = r
        for s in sotp_rows:
            ws.cell(r, 1, s["segment"])
            if s.get("value_cr") is None:
                _put(ws, r, 2, float(s["metric_value_cr"]), F_NUM, color="C55A11")
                _put(ws, r, 3, float(s["multiple"]), F_X1, color="C55A11")
                _put(ws, r, 4, f"=B{r}*C{r}", F_NUM)
            else:
                _put(ws, r, 4, float(s["value_cr"]), F_NUM, color="C55A11")
            ws.cell(r, 5, s.get("method", ""))
            ws.cell(r, 9, s["source_text"]).font = Font(size=9, color=GREY)
            for cc in range(1, 5):
                ws.cell(r, cc).fill = _fill(OVERRIDE_FILL)
            r += 1
        R["sotp_ev"] = r
        ws.cell(r, 1, "Total segment EV (₹ Cr)").font = Font(bold=True)
        _put(ws, r, 4, f"=SUM(D{s0}:D{r - 1})", F_NUM, bold=True)
        r += 1
        R["sotp_adj"] = r
        ws.cell(r, 1, "Less: other adjustment (₹ Cr) " + (f"- {ai.sotp_adjustment_note}" if ai else ""))
        _put(ws, r, 4, ai.sotp_other_adjustment_cr if ai else 0.0, F_NUM, color="C55A11")
        r += 1
        R["sotp_value"] = r
        ws.cell(r, 1, "SOTP value per share (₹)").font = Font(bold=True)
        _put(ws, r, 4, f"=(D{R['sotp_ev']}+C{R['bridge']}-D{R['sotp_adj']})/{shares}", F_RS, bold=True)
        sotp_active, sotp_ref = 1, f"D{R['sotp_value']}"
    else:
        ws.cell(r, 1, "Not used (no SOTP in analyst_inputs.json).").font = Font(size=9, color=GREY)
        sotp_active, sotp_ref = 0, "0"
    r += 2
    _section(ws, r, "5. Target, return and rating", pal, 8)
    r += 1
    final = [
        ("sotp_on", "SOTP active (1 = yes)", sotp_active, '0'),
        ("target", "TARGET PRICE, 12 months (₹)", None, F_RS),
        ("price", "Current price (₹)", f"={D['price']}", F_RS),
        ("upside", "Upside / (downside)", None, '+0.0%;-0.0%;0.0%'),
        ("div_yield", "Dividend yield", f"={D['div_yield']}", '0.00%'),
        ("etr", "Expected total return (upside + dividend yield)", None, '+0.0%;-0.0%;0.0%'),
        ("rating", "Rating", None, None),
    ]
    for key, lab, f, fmt in final:
        R[key] = r
        ws.cell(r, 1, lab).font = Font(bold=key in ("target", "rating", "etr"))
        if f is not None:
            _put(ws, r, 3, f, fmt)
        r += 1
    _put(ws, R["target"], 3, f"=IF(C{R['sotp_on']}=1,{sotp_ref},E{R['blend']})", F_RS, bold=True)
    _put(ws, R["upside"], 3, f"=C{R['target']}/C{R['price']}-1", '+0.0%;-0.0%;0.0%')
    _put(ws, R["etr"], 3, f"=C{R['upside']}+C{R['div_yield']}", '+0.0%;-0.0%;0.0%', bold=True)
    _put(ws, R["rating"], 3, _rating_formula(f"C{R['etr']}", ctx.cfg), None, bold=True)
    for key in ("target", "rating"):
        ws.cell(R[key], 3).fill = _fill(pal.light)
        ws.cell(R[key], 1).fill = _fill(pal.light)
    th = ", ".join(f"{lab} >= {t:+.0%}" for t, lab in ctx.cfg["rating_thresholds"]) + f", else {ctx.cfg.get('rating_floor', 'SELL')}"
    ws.cell(R["rating"], 9, f"Rating on expected total return: {th}").font = Font(size=9, color=GREY)
    ctx.rows[S_VAL] = R
    ctx.refs.update({"target": f"{q(S_VAL)}!$C${R['target']}", "rating": f"{q(S_VAL)}!$C${R['rating']}",
                     "upside": f"{q(S_VAL)}!$C${R['upside']}", "etr": f"{q(S_VAL)}!$C${R['etr']}",
                     "blend": f"{q(S_VAL)}!$E${R['blend']}", "pe_value": f"{q(S_VAL)}!$C${R['pe_value']}",
                     "ev_value": f"{q(S_VAL)}!$C${R['ev_value']}"})
    ws.sheet_view.showGridLines = False


# --------------------------------------------------------------------------- #
# Scenarios
# --------------------------------------------------------------------------- #
SCEN_ROWS = ["sales_growth", "ebitda_margin", "capex_pct", "ronic", "dcf_value", "dcf_value_12m", "pe_value",
             "ev_ebitda_value", "target", "expected_total_return", "probability"]


def write_scenarios(wb, ctx: Ctx):
    ws = wb.create_sheet(S_SCEN)
    pal, v, D = ctx.palette, ctx.val, ctx.drv
    _title(ws, f"{ctx.fin.name} - Bull / Base / Bear",
           "Columns B-D are stored results of the Python engine for each scenario (values). Column E is live: it follows "
           "the scenario chosen in Drivers!C4. The check cell must be 0. Scenario targets use the model blend.", pal)
    ws.column_dimensions["A"].width = 46
    for col in "BCDEF":
        ws.column_dimensions[col].width = 14
    _head_row(ws, 4, ["", "Bull", "Base", "Bear", "Live (Drivers!C4)"], pal)
    labels = {"sales_growth": ("Revenue growth (driver)", F_PCT), "ebitda_margin": ("EBITDA margin (driver)", F_PCT),
              "capex_pct": ("Capex % of revenue (driver)", F_PCT), "ronic": ("Terminal RONIC", F_PCT),
              "dcf_value": ("DCF value today (₹)", F_RS), "dcf_value_12m": ("DCF value, 12 months (₹)", F_RS),
              "pe_value": ("P/E value (₹)", F_RS), "ev_ebitda_value": ("EV/EBITDA value (₹)", F_RS),
              "target": ("Scenario value, model blend (₹)", F_RS),
              "expected_total_return": ("Expected total return", '+0.0%;-0.0%'), "probability": ("Probability", '0%')}
    dcf = DCF_ROWS
    live = {"sales_growth": f"={D['sales_growth']}", "ebitda_margin": f"={D['ebitda_margin']}",
            "capex_pct": f"={D['capex_pct']}", "ronic": f"={q(S_DCF)}!$C${dcf['ronic']}",
            "dcf_value": f"={q(S_DCF)}!$C${dcf['dcf_value']}", "dcf_value_12m": f"={q(S_DCF)}!$C${dcf['dcf_12m']}",
            "pe_value": f"={ctx.refs['pe_value']}", "ev_ebitda_value": f"={ctx.refs['ev_value']}",
            "target": f"={ctx.refs['blend']}",
            "expected_total_return": f"={ctx.refs['blend']}/{D['price']}-1+{D['div_yield']}"}
    R = {}
    for i, key in enumerate(SCEN_ROWS):
        r = 5 + i
        R[key] = r
        lab, fmt = labels[key]
        ws.cell(r, 1, lab)
        for j, s in enumerate(("bull", "base", "bear"), 2):
            sv = v.scenarios.get(s)
            val = getattr(sv, key) if sv is not None else None
            if val is None:
                ws.cell(r, j, "n/a")
                continue
            _put(ws, r, j, val, fmt, color=GREY)
        if key in live:
            _put(ws, r, 5, live[key], fmt, bold=key == "target")
    r = 5 + len(SCEN_ROWS) + 1
    ws.cell(r, 1, "Probability-weighted value (₹)").font = Font(bold=True)
    _put(ws, r, 3, f"=SUMPRODUCT(B{R['probability']}:D{R['probability']},B{R['target']}:D{R['target']})", F_RS, bold=True)
    R["weighted"] = r
    r += 1
    ws.cell(r, 1, "Check: live value - stored value of the live scenario (must be 0)").font = Font(bold=True)
    _put(ws, r, 3, f"=ROUND(E{R['target']}-CHOOSE(Drivers!$C$4,C{R['target']},B{R['target']},D{R['target']}),2)",
         '0.00', bold=True)
    R["check"] = r
    ctx.rows[S_SCEN] = R
    ws.sheet_view.showGridLines = False


# --------------------------------------------------------------------------- #
# Analyst inputs
# --------------------------------------------------------------------------- #
def write_analyst_inputs(wb, ctx: Ctx):
    ws = wb.create_sheet(S_AI)
    pal, ai = ctx.palette, ctx.analyst
    _title(ws, f"{ctx.fin.name} - Analyst Inputs",
           "Sourced inputs from <SYMBOL>_analyst_inputs.json written by the host agent. Each changes the model only where "
           "shown (orange cells on the Drivers / Valuation Summary sheets).", pal)
    for col, w in zip("ABCDE", (34, 16, 16, 50, 90)):
        ws.column_dimensions[col].width = w
    _head_row(ws, 4, ["Item", "Analyst value", "Model value", "Rationale", "Source"], pal)
    if ai is None:
        ws.cell(5, 1, "None: the model uses only reported data and config/market.json.").font = Font(color=GREY)
        return
    if ai.prepared_by or ai.as_of:
        ws.cell(3, 1, f"Prepared by {ai.prepared_by or 'n/a'} on {ai.as_of or 'n/a'} ({ai.path})").font = Font(size=9, color=GREY)
    for i, e in enumerate(ai.entries()):
        r = 5 + i
        ws.cell(r, 1, e["item"])
        c = ws.cell(r, 2, e["value"])
        if isinstance(e["value"], float) and abs(e["value"]) < 5:
            c.number_format = '0.0%'
        elif isinstance(e["value"], float):
            c.number_format = F_NUM
        c.fill = _fill(OVERRIDE_FILL)
        if e["model"] is not None:
            m = ws.cell(r, 3, e["model"])
            m.number_format = '0.0%' if abs(e["model"]) < 5 else '0.0'
        ws.cell(r, 4, e["rationale"]).alignment = Alignment(wrap_text=True)
        ws.cell(r, 5, e["source"]).alignment = Alignment(wrap_text=True)
    ws.sheet_view.showGridLines = False


# --------------------------------------------------------------------------- #
# Cover page
# --------------------------------------------------------------------------- #
def write_cover(wb, ctx: Ctx):
    ws = wb.active
    ws.title = S_COVER
    pal, fin, val = ctx.palette, ctx.fin, ctx.val
    ws.sheet_view.showGridLines = False
    ws.column_dimensions["A"].width = 3
    ws.column_dimensions["B"].width = 34
    ws.column_dimensions["C"].width = 60
    for r in range(1, 5):
        for col in range(1, 5):
            ws.cell(r, col).fill = _fill(pal.dark)
    ws["B2"] = fin.name
    ws["B2"].font = Font(size=22, bold=True, color=pal.text_on_dark)
    ws["B3"] = f"NSE: {fin.symbol}" + (f"   BSE: {fin.bse_code}" if fin.bse_code else "") + f"   |   {fin.sector} - {fin.industry}"
    ws["B3"].font = Font(size=11, color=pal.text_on_dark)
    view = "consolidated" if fin.consolidated else "standalone"
    info = [
        ("Model date", date.today().isoformat()),
        ("Data source", f"Screener.in ({view}), fetched {fin.fetched_at} - {fin.source_url}"),
        ("Price data", "yfinance (beta, price history) - " + (val.beta_source or "")),
        ("Units", "₹ crore unless stated; per-share values in ₹"),
        ("Fiscal years", f"FY{fin.years[0] % 100:02d}A - FY{fin.years[-1] % 100:02d}A reported, "
                         f"{ctx.forecast.year_labels[0]} - {ctx.forecast.year_labels[-1]} forecast, then "
                         f"{int(ctx.cfg.get('fade_years', 5))} fade years"),
        ("Current price (₹)", f"{fin.price:,.2f}"),
        ("Market cap (₹ Cr)", f"{fin.market_cap:,.0f}"),
        ("12-month target price (₹, base case)", f"{val.target_price:,.2f}"),
        ("Upside / (downside)", f"{val.upside:+.1%}"),
        ("Expected total return", f"{val.expected_total_return:+.1%}"),
        ("Rating", val.rating),
        ("Analyst inputs", ctx.analyst.path if ctx.analyst is not None else "none"),
    ]
    r = 6
    for k, v in info:
        ws.cell(r, 2, k).font = Font(bold=True)
        ws.cell(r, 3, v)
        r += 1
    r += 1
    ws.cell(r, 2, "Contents").font = Font(size=13, bold=True, color=pal.dark)
    for name in (S_DASH, S_DRV, S_AI, S_CB, S_IS, S_PPE, S_WC, S_CF, S_BS, S_WACC, S_DCF, S_VAL, S_PEERS, S_SCEN):
        r += 1
        c = ws.cell(r, 2, f'=HYPERLINK("#\'{name}\'!A1","{name}")')
        c.font = Font(color="0563C1", underline="single")
    r += 2
    ws.cell(r, 2, "How to read this model").font = Font(size=13, bold=True, color=pal.dark)
    notes = [
        "Blue numbers are reported figures or inputs; black numbers are formulas; orange cells are sourced analyst inputs.",
        "Change the scenario in Drivers!C4 (1 Base, 2 Bull, 3 Bear). Only forecast years and the valuation change.",
        "Change the year on the Dashboard (cell E6) to update the KPI cards and charts.",
        "Liquidity ratios are approximations: Screener does not split current and non-current items.",
        "This is an educational model, not investment advice.",
    ]
    for t in notes:
        r += 1
        ws.cell(r, 2, "- " + t)
    if ctx.flags:
        r += 2
        ws.cell(r, 2, "Model credibility flags").font = Font(size=13, bold=True, color="C00000")
        for fl in ctx.flags:
            r += 1
            ws.cell(r, 2, f"[{fl.severity.upper()}] {fl.code}").font = Font(bold=fl.severity == "high",
                                                                            color="C00000" if fl.severity == "high" else BLACK)
            ws.cell(r, 3, fl.message)
    warnings = list(fin.warnings) + list(ctx.drivers.notes) + list(val.warnings) + (list(ctx.md.warnings) if ctx.md else [])
    if warnings:
        r += 2
        ws.cell(r, 2, "Data notes & warnings").font = Font(size=13, bold=True, color="C00000")
        for t in warnings:
            r += 1
            ws.cell(r, 2, "- " + t).alignment = Alignment(wrap_text=False)


# --------------------------------------------------------------------------- #
# entry point
# --------------------------------------------------------------------------- #
def build_workbook(fin: CompanyFinancials, drivers: Drivers, forecast: Forecast, val: Valuation, palette: Palette,
                   cfg: dict, md: Optional[MarketData] = None, peers=None, analyst=None, flags=None) -> Workbook:
    from .config import with_defaults
    from .dashboard import build_dashboard  # local import: dashboard needs the row layouts

    ctx = Ctx(fin=fin, drivers=drivers, forecast=forecast, val=val, palette=palette, cfg=with_defaults(cfg), md=md,
              peers=peers, analyst=analyst, flags=list(flags or []))
    build_layouts(ctx)
    ctx.rows[S_DCF] = DCF_ROWS
    wb = Workbook()
    write_cover(wb, ctx)
    wb.create_sheet(S_DASH)  # placeholder position; filled by build_dashboard
    write_drivers(wb, ctx)
    write_analyst_inputs(wb, ctx)
    write_cost_breakdown(wb, ctx)
    write_income_statement(wb, ctx)
    write_ppe(wb, ctx)
    write_working_capital(wb, ctx)
    write_cash_flow(wb, ctx)
    write_balance_sheet(wb, ctx)
    write_wacc(wb, ctx)
    write_dcf(wb, ctx)
    write_peers(wb, ctx)
    write_valuation_summary(wb, ctx)
    write_scenarios(wb, ctx)
    # sheet order: Valuation Summary before Peers
    wb.move_sheet(S_VAL, offset=-1)
    build_dashboard(wb, ctx)
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    wb.active = wb.sheetnames.index(S_DASH)
    return wb


def write_model(path, *args, **kwargs) -> str:
    wb = build_workbook(*args, **kwargs)
    wb.save(path)
    return str(path)
