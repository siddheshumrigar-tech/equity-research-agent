"""The PDF report and the Excel model show the same value for every shared ratio, in every year.

Both read their definitions from ``equity_agent.ratios``: the PDF calls it directly, the Excel model writes the same
formulas. These tests recalculate the workbook and compare it with the Python series, for a service company
(Tata Technologies) and a manufacturer (Asian Paints).
"""
import openpyxl
import pytest
from openpyxl.utils import get_column_letter

from equity_agent.ratios import annual

# (sheet, start of the row label in column A) -> key in equity_agent.ratios.annual
ROWS = {
    ("Income Statement", "Gross margin %"): "gross_margin",
    ("Balance Sheet", "Return on equity"): "roe",
    ("Balance Sheet", "Return on assets"): "roa",
    ("Balance Sheet", "ROCE, pre-tax"): "roce",
    ("Working Capital", "Receivable days (DSO)"): "dso",
    ("Working Capital", "Inventory days (DIO)"): "dio",
    ("Working Capital", "Payable days (DPO)"): "dpo",
    ("Working Capital", "Cash conversion cycle"): "ccc",
}


@pytest.fixture(scope="module")
def recalc_asian(built_asian):
    import recalc

    if not recalc.available():
        pytest.skip("No recalculation engine (LibreOffice or the 'formulas' package) installed")
    return recalc.recalc(built_asian.xlsx_path)


def _row(ws, prefix):
    rows = [r for r in range(1, ws.max_row + 1)
            if isinstance(ws.cell(r, 1).value, str) and ws.cell(r, 1).value.strip().startswith(prefix)]
    assert len(rows) == 1, f"{ws.title}: expected one row starting '{prefix}', found {rows}"
    return rows[0]


def _check_ratios(built, values):
    wb = openpyxl.load_workbook(built.xlsx_path)
    expected = annual(built)
    n_all = built.fin.n + len(built.forecast.years)
    for (sheet, prefix), key in ROWS.items():
        r = _row(wb[sheet], prefix)
        for i in range(n_all):
            got = values.get((sheet.upper(), f"{get_column_letter(2 + i)}{r}"))
            want = expected[key][i]
            where = f"{sheet} '{prefix}' {expected['labels'][i]}"
            if want is None:
                assert got in (None, "n/a"), f"{where}: Excel {got!r}, PDF n/a"
            else:
                assert got == pytest.approx(want, rel=1e-9, abs=1e-9), f"{where}: Excel {got!r}, PDF {want!r}"


def test_excel_ratios_equal_pdf_ratios_services(built_tata, recalc_tata):
    assert built_tata.fin.cost_basis == "services"
    _check_ratios(built_tata, recalc_tata)


def test_excel_ratios_equal_pdf_ratios_goods(built_asian, recalc_asian):
    assert built_asian.fin.cost_basis == "goods"
    _check_ratios(built_asian, recalc_asian)


def test_dashboard_cards_use_the_same_returns(built_tata, recalc_tata):
    rt = annual(built_tata)
    k = built_tata.fin.n - 1  # the dashboard opens on the latest reported year
    assert recalc_tata[("DASHBOARD", "P3")] == pytest.approx(rt["roa"][k], rel=1e-9)
    assert recalc_tata[("DASHBOARD", "R3")] == pytest.approx(rt["roe"][k], rel=1e-9)
    L = built_tata.fin.lines
    assert recalc_tata[("DASHBOARD", "L3")] == pytest.approx(L["gross_profit"][k])


def test_pdf_tables_use_the_shared_ratios(res_tata):
    from equity_agent.report.metrics import key_metrics

    m, rt = key_metrics(res_tata), annual(res_tata)
    assert m["roe"] == rt["roe"] and m["roce"] == rt["roce"]


def test_peer_tables_and_tokens_use_the_shared_ratios(res_tata):
    from equity_agent.narrative import build_tokens
    from equity_agent.peers import subject_history, subject_snapshot

    f = res_tata.fin
    rt = annual(res_tata)
    k = f.n - 1
    snap = subject_snapshot(f)
    assert snap.roe == pytest.approx(rt["roe"][k] * 100, rel=1e-12)
    assert snap.roce == pytest.approx(rt["roce"][k] * 100, rel=1e-12)
    hist = subject_history(f)
    for j, y in enumerate(hist["roe"]):  # last 5 reported years, oldest first
        i = f.n - len(hist["roe"]) + j
        assert hist["roe"][y] == pytest.approx(rt["roe"][i], rel=1e-12)
        assert hist["roce"][y] == pytest.approx(rt["roce"][i], rel=1e-12)
    tok = build_tokens(res_tata)
    assert tok["roe"] == f"{rt['roe'][k] * 100:.1f}%"
    assert tok["roce"] == f"{rt['roce'][k] * 100:.1f}%"
    for p in res_tata.peers.peers:  # every peer on the same definition: latest year of its own history
        if p.history and p.history.get("roe"):
            latest = list(p.history["roe"].values())[-1]
            assert p.roe == (None if latest is None else pytest.approx(latest * 100))
