import shutil
import zipfile

import openpyxl
import pytest

import recalc
from equity_agent.dashboard import CHART_ANCHORS
from equity_agent.excel_writer import DCF_ROWS

EXPECTED_SHEETS = ["Cover Page", "Dashboard", "Drivers", "Analyst Inputs", "Cost Breakdown", "Income Statement",
                   "PP&E Schedule", "Working Capital", "Cash Flow", "Balance Sheet", "CAPM & WACC", "DCF Valuation",
                   "Valuation Summary", "Peers", "Scenarios", "Dashboard_Calc"]


@pytest.fixture(scope="module")
def wb_tata(built_tata):
    return openpyxl.load_workbook(built_tata.xlsx_path)


def _labels(ws):
    return {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value}


def test_sheet_order_and_hidden_calc(wb_tata):
    assert wb_tata.sheetnames == EXPECTED_SHEETS
    assert wb_tata["Dashboard_Calc"].sheet_state == "hidden"
    assert wb_tata.calculation.fullCalcOnLoad is True


def test_statement_row_labels(wb_tata):
    is_labels = _labels(wb_tata["Income Statement"])
    for lab in ("Revenue (Sales)", "Gross profit", "EBITDA (Operating profit)", "Profit before tax", "Net profit (reported)"):
        assert lab in is_labels
    bs_labels = _labels(wb_tata["Balance Sheet"])
    for lab in ("TOTAL ASSETS", "TOTAL EQUITY & LIABILITIES", "Balance check (assets - liabilities) = 0", "Cash & equivalents"):
        assert lab in bs_labels
    assert not any("Plug" in str(k) for k in bs_labels)


def test_actual_cells_are_reported_values(wb_tata, fin_tata):
    ws = wb_tata["Income Statement"]
    rows = _labels(ws)
    assert [ws.cell(3, 2 + i).value for i in range(fin_tata.n)] == fin_tata.year_labels
    assert [ws.cell(4, 2 + i).value for i in range(fin_tata.n)] == fin_tata.years
    for i in range(fin_tata.n):
        assert ws.cell(rows["Revenue (Sales)"], 2 + i).value == fin_tata.lines["sales"][i]
        assert ws.cell(rows["Net profit (reported)"], 2 + i).value == fin_tata.lines["net_profit"][i]
    first_fc = ws.cell(rows["Revenue (Sales)"], 2 + fin_tata.n).value
    assert isinstance(first_fc, str) and first_fc.startswith("=") and "Drivers!" in first_fc
    bs = wb_tata["Balance Sheet"]
    brows = _labels(bs)
    assert bs.cell(brows["Cash & equivalents"], 1 + fin_tata.n).value == 715


def test_dashboard_layout(wb_tata, fin_tata):
    ws = wb_tata["Dashboard"]
    assert ws.sheet_view.showGridLines is False
    assert ws.sheet_view.showRowColHeaders is False
    assert ws.sheet_view.zoomScale == 90
    merged = {str(m) for m in ws.merged_cells.ranges}
    for rng in ("B6:C7", "E6:E7", "T2:V6", "F1:G2", "R3:S4"):
        assert rng in merged
    assert ws["B6"].value == "Financial Year"
    assert ws["E6"].value == fin_tata.years[-1]
    assert ws.column_dimensions["A"].width == 1.5 and ws.column_dimensions["E"].width == 12.8
    assert ws.column_dimensions["T"].width == 19
    dvs = [dv for dv in ws.data_validations.dataValidation if "E6" in str(dv.sqref)]
    assert len(dvs) == 1
    lst = dvs[0].formula1.strip('"')
    assert " " not in lst
    assert lst.split(",") == [str(y) for y in fin_tata.years + [fin_tata.years[-1] + k for k in range(1, 6)]]
    titles = [ws.cell(1, c).value for c in range(6, 20, 2)]
    assert titles == ["Revenue (₹ Cr)", "COGS (₹ Cr)", "OPEX (₹ Cr)", "Gross Profit (₹ Cr)", "Net Profit (₹ Cr)", "ROA", "ROE"]
    assert ws["F3"].value == "='Dashboard_Calc'!$B$5"
    assert ws["F6"].value == "VS Pre Year"
    assert ws["X6"].value.endswith(f"$B${ws['X6'].value.split('$B$')[1]}")
    assert ws["W7"].value == "STATUS"


def test_dashboard_has_eight_anchored_charts(wb_tata, built_tata):
    with zipfile.ZipFile(built_tata.xlsx_path) as z:
        charts = sorted(n for n in z.namelist() if n.startswith("xl/charts/chart"))
        drawing = [n for n in z.namelist() if n.startswith("xl/drawings/drawing")]
        assert len(charts) == 8
        xml = "".join(z.read(n).decode("utf-8") for n in charts)
        dr = "".join(z.read(n).decode("utf-8") for n in drawing)
    pal = built_tata.palette
    assert pal.primary in xml and pal.dark in xml
    assert "Dashboard_Calc" in xml
    assert "70AD47" not in xml or pal.primary == "70AD47"
    for c1, r1, c2, r2 in CHART_ANCHORS.values():
        assert f"<xdr:col>{c1}</xdr:col>" in dr or f"<col>{c1}</col>" in dr
    for title in ("Revenue Growth Y-o-Y", "Expense Breakup", "Where Revenue Goes", "Margins", "P&amp;L Waterfall",
                  "Trade Cycle", "Liquidity Ratios", "Free Cash Flow"):
        assert title in xml
    assert "<c:layout><c:layout>" not in xml


def _all_text(path):
    wb = openpyxl.load_workbook(path)
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    out.append(str(c.value))
    return " ".join(out)


def test_no_hardcoded_company_strings(built_asian):
    text = _all_text(built_asian.xlsx_path)
    for bad in ("Tata", "TATATECH", "950.02", "805.10", "Automotive Eng.", "Aerospace", "Rs. 32,691"):
        assert bad not in text
    assert "Asian Paints" in text


# ---------------------------------------------------------------- recalculated
def test_balance_check_zero_every_year(recalc_tata, fin_tata, wb_tata):
    row = _labels(wb_tata["Balance Sheet"])["Balance check (assets - liabilities) = 0"]
    n_all = fin_tata.n + 5
    for i in range(n_all):
        col = openpyxl.utils.get_column_letter(2 + i)
        assert recalc_tata[("BALANCE SHEET", f"{col}{row}")] == pytest.approx(0, abs=0.005)


def test_excel_dcf_matches_python(recalc_tata, res_tata):
    v = res_tata.valuation
    get = lambda key: recalc_tata[("DCF VALUATION", f"C{DCF_ROWS[key]}")]
    assert get("wacc") == pytest.approx(v.wacc, abs=1e-9)
    assert get("roic") == pytest.approx(v.roic, abs=1e-9)
    assert get("ronic") == pytest.approx(v.ronic, abs=1e-9)
    assert get("tv") == pytest.approx(v.terminal_value, rel=1e-9)
    assert get("tv_share") == pytest.approx(v.tv_share, abs=1e-9)
    assert get("dcf_value") == pytest.approx(v.dcf_value, abs=0.01)
    assert get("dcf_12m") == pytest.approx(v.dcf_value_12m, abs=0.01)
    assert get("exit_value") == pytest.approx(v.exit_value, abs=0.01)
    for j in range(10):  # explicit and fade years
        col = openpyxl.utils.get_column_letter(2 + j)
        assert recalc_tata[("DCF VALUATION", f"{col}{DCF_ROWS['fcff']}")] == pytest.approx(v.fcff[j], abs=0.01)
    # every cell of the WACC x g grid re-runs the fade and terminal stages
    h = DCF_ROWS["sens_head"]
    for i in range(5):
        for j in range(5):
            col = openpyxl.utils.get_column_letter(2 + j)
            assert recalc_tata[("DCF VALUATION", f"{col}{h + 1 + i}")] == pytest.approx(v.sensitivity_grid[i][j], abs=0.01)


def test_excel_target_and_rating_match_python(recalc_tata, res_tata, wb_tata):
    v = res_tata.valuation
    labels = _labels(wb_tata["Valuation Summary"])
    get = lambda lab: recalc_tata[("VALUATION SUMMARY", f"C{labels[lab]}")]
    assert get("Mean P/E (x)") == pytest.approx(v.pe_mean, rel=1e-9)
    assert get("P/E value per share (₹) = mean P/E x NTM EPS") == pytest.approx(v.pe_value, abs=0.01)
    assert get("Peer median EV/EBITDA (x)") == pytest.approx(v.peer_ev_ebitda, rel=1e-9)
    assert get("EV/EBITDA value per share (₹)") == pytest.approx(v.ev_ebitda_value, abs=0.01)
    assert get("TARGET PRICE, 12 months (₹)") == pytest.approx(v.target_price, abs=0.01)
    assert get("Expected total return (upside + dividend yield)") == pytest.approx(v.expected_total_return, abs=1e-6)
    assert get("Rating") == v.rating


def test_scenarios_sheet_check_cell(recalc_tata, wb_tata, res_tata):
    labels = _labels(wb_tata["Scenarios"])
    r = labels["Check: live value - stored value of the live scenario (must be 0)"]
    assert recalc_tata[("SCENARIOS", f"C{r}")] == pytest.approx(0, abs=0.01)
    rw = labels["Probability-weighted value (₹)"]
    assert recalc_tata[("SCENARIOS", f"C{rw}")] == pytest.approx(res_tata.valuation.prob_weighted_value, abs=0.01)


def test_scenario_switch_check_cell_bull(built_tata, tmp_path):
    vals = _recalc_modified(built_tata.xlsx_path, tmp_path, {("Drivers", "C4"): 2})
    wb = openpyxl.load_workbook(built_tata.xlsx_path)
    r = _labels(wb["Scenarios"])["Check: live value - stored value of the live scenario (must be 0)"]
    assert vals[("SCENARIOS", f"C{r}")] == pytest.approx(0, abs=0.01)


def test_peers_sheet(wb_tata, res_tata):
    ws = wb_tata["Peers"]
    names = [ws.cell(r, 2).value for r in range(5, 5 + 1 + len(res_tata.peers.peers))]
    assert names[0] == "TATATECH"
    assert names[1:] == [p.symbol for p in res_tata.peers.peers]


def test_peer_history_in_excel_matches_python(recalc_tata, wb_tata, res_tata):
    from equity_agent.excel_writer import PEER_HISTORY_LABELS
    from equity_agent.peers import HISTORY_METRICS, subject_history

    ws, ps = wb_tata["Peers"], res_tata.peers
    years = ps.history_years()
    assert ps.has_history and len(years) == 5
    col_a = {ws.cell(r, 1).value: r for r in range(1, ws.max_row + 1) if ws.cell(r, 1).value}
    subj = subject_history(res_tata.fin)
    for metric in HISTORY_METRICS:
        h = col_a[PEER_HISTORY_LABELS[metric]]
        assert [ws.cell(h, 3 + j).value for j in range(5)] == years
        med_row = next(r for r in range(h + 1, ws.max_row + 1) if ws.cell(r, 1).value == "Peer median (excl. subject)")
        for j, y in enumerate(years):
            col = chr(ord("C") + j)
            med = ps.history_median(metric, y)
            got = recalc_tata[("PEERS", f"{col}{med_row}")]
            if med is None:
                assert got == "n/a"
                continue
            assert got == pytest.approx(med, abs=1e-9)
            if subj[metric].get(y) is not None:
                assert recalc_tata[("PEERS", f"{col}{med_row + 1}")] == pytest.approx(subj[metric][y] - med, abs=1e-9)


def test_dashboard_cards_latest_year(recalc_tata, fin_tata):
    L = fin_tata.lines
    assert recalc_tata[("DASHBOARD", "F3")] == pytest.approx(L["sales"][-1])
    assert recalc_tata[("DASHBOARD", "N3")] == pytest.approx(L["net_profit"][-1])
    assert recalc_tata[("DASHBOARD", "F5")] == pytest.approx(L["sales"][-1] / L["sales"][-2] - 1)
    assert recalc_tata[("DASHBOARD", "X5")] == pytest.approx(fin_tata.price)


def test_no_unexpected_errors(recalc_tata):
    bad = [(k, v) for k, v in recalc_tata.items() if isinstance(v, str) and v.startswith("#")
           and k[0] not in ("COVER PAGE", "DASHBOARD_CALC")]
    assert bad == []


def _recalc_modified(path, tmp_path, edits):
    if not recalc.available():
        pytest.skip("No recalculation engine installed")
    dst = tmp_path / "edited.xlsx"
    shutil.copy(path, dst)
    wb = openpyxl.load_workbook(dst)
    for (sheet, addr), value in edits.items():
        wb[sheet][addr] = value
    wb.save(dst)
    return recalc.recalc(dst)


def test_year_selector_changes_cards(built_tata, fin_tata, tmp_path):
    vals = _recalc_modified(built_tata.xlsx_path, tmp_path, {("Dashboard", "E6"): 2023})
    k = fin_tata.years.index(2023)
    L = fin_tata.lines
    assert vals[("DASHBOARD", "F3")] == pytest.approx(L["sales"][k])
    assert vals[("DASHBOARD", "H3")] == pytest.approx(L["cogs"][k])
    assert vals[("DASHBOARD", "N3")] == pytest.approx(L["net_profit"][k])
    assert vals[("DASHBOARD", "F5")] == pytest.approx(L["sales"][k] / L["sales"][k - 1] - 1)


def test_scenario_switch_changes_forecast_only(built_tata, res_tata, tmp_path):
    from equity_agent.model_engine import run_forecast

    vals = _recalc_modified(built_tata.xlsx_path, tmp_path, {("Drivers", "C4"): 2})
    f = res_tata.fin
    bull = run_forecast(f, res_tata.drivers, "bull")
    wb = openpyxl.load_workbook(built_tata.xlsx_path)
    row = _labels(wb["Income Statement"])["Revenue (Sales)"]
    last_a = openpyxl.utils.get_column_letter(1 + f.n)
    first_f = openpyxl.utils.get_column_letter(2 + f.n)
    assert vals[("INCOME STATEMENT", f"{last_a}{row}")] == pytest.approx(f.lines["sales"][-1])
    assert vals[("INCOME STATEMENT", f"{first_f}{row}")] == pytest.approx(bull.lines["sales"][0])


def test_working_capital_glide_in_excel(recalc_tata, res_tata, wb_tata):
    rows = _labels(wb_tata["Working Capital"])
    n = res_tata.fin.n
    for j in range(5):
        col = openpyxl.utils.get_column_letter(2 + n + j)
        assert recalc_tata[("WORKING CAPITAL", f"{col}{rows['Trade receivables']}")] == pytest.approx(
            res_tata.forecast.lines["receivables"][j], abs=0.01)
        assert recalc_tata[("WORKING CAPITAL", f"{col}{rows['Operating net working capital']}")] == pytest.approx(
            res_tata.forecast.lines["nwc"][j], abs=0.01)
