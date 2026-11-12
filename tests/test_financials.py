import copy

import pytest

from equity_agent.errors import DataValidationError
from equity_agent.financials import build_financials, classify_expense


def test_tata_years_and_reported_values(fin_tata):
    f = fin_tata
    assert f.years == list(range(2019, 2027))
    assert f.year_labels[-1] == "FY26A"
    i25 = f.years.index(2025)
    L = f.lines
    assert L["sales"][i25] == 5168
    assert L["ebitda"][-1] == 853
    assert L["net_profit"][-1] == 547
    assert L["receivables"][-1] == 1187
    assert L["cash"][-1] == 715
    assert L["capex"][-1] == 34
    assert L["trade_payables"][-1] == 581
    assert L["lease_liabilities"][-1] == 272
    assert f.gross_block[-1] == 2858 and f.acc_dep[-1] == 803
    assert f.price == 676 and f.market_cap == 27470
    assert f.warnings == []


def test_cogs_plus_opex_equals_expenses(fin_tata, fin_asian):
    for f in (fin_tata, fin_asian):
        for c, o, e in zip(f.lines["cogs"], f.lines["opex"], f.lines["expenses"]):
            assert c + o == pytest.approx(e)


def test_components_add_to_reported_totals(fin_asian):
    L = fin_asian.lines
    for k in range(fin_asian.n):
        assets = (L["fixed_assets"][k] + L["cwip"][k] + L["investments"][k] + L["inventory"][k] + L["receivables"][k]
                  + L["cash"][k] + L["loans_advances"][k] + L["other_asset_items"][k])
        liabs = (L["equity_capital"][k] + L["reserves"][k] + L["borrowings"][k] + L["trade_payables"][k]
                 + L["advance_from_customers"][k] + L["minority_interest"][k] + L["other_liability_items"][k])
        assert assets == pytest.approx(L["total_assets"][k])
        assert liabs == pytest.approx(L["total_liabilities"][k])


def test_asian_paints_manufacturer_lines(fin_asian):
    L = fin_asian.lines
    assert fin_asian.years[0] == 2015 and fin_asian.years[-1] == 2026
    assert L["inventory"][-1] == 5744
    assert L["minority_interest"][-1] == 643
    assert L["cogs"][-1] == 17746
    assert fin_asian.shares_cr == pytest.approx(96.0, rel=0.01)


def test_cost_of_revenue_follows_the_cost_mix(fin_tata, fin_asian):
    # Tata Technologies: material is ~21% of expenses -> service cost structure, employee cost is cost of revenue
    assert fin_tata.cost_basis == "services"
    kind = {it.name.lower(): it.kind for it in fin_tata.expense_items}
    assert [kind[k] for k in sorted(kind)] == ["COGS", "COGS", "OPEX"]  # employee, material, other
    assert fin_tata.lines["cogs"][-1] == pytest.approx(967 + 2737)
    # Asian Paints: manufacturer, employee cost stays an operating expense
    assert fin_asian.cost_basis == "goods"
    assert all(it.kind == "OPEX" for it in fin_asian.expense_items if "employee" in it.name.lower())


def test_tax_and_minority_reconcile_to_reported_net_profit(fin_tata):
    L = fin_tata.lines
    for p, t, m, npf in zip(L["pbt"], L["tax"], L["minority_other"], L["net_profit"]):
        assert p - t - m == pytest.approx(npf, abs=0.01)


def _set(raw, section, label, header, value):
    t = raw.tables[section]
    t.rows[label][t.headers.index(header)] = value


def test_large_balance_mismatch_stops_run(raw_tata):
    raw = copy.deepcopy(raw_tata)
    _set(raw, "balance-sheet", "Total Assets", "Mar 2024", 9999)
    with pytest.raises(DataValidationError):
        build_financials(raw)


def test_small_rounding_gap_is_a_warning(raw_tata):
    raw = copy.deepcopy(raw_tata)
    _set(raw, "profit-loss", "Operating Profit", "Mar 2024", 941 + 40)  # 0.8% of sales: above tolerance, below hard limit
    f = build_financials(raw)
    assert any("Operating Profit" in w for w in f.warnings)


def test_missing_other_assets_schedule_stops_run(raw_tata):
    raw = copy.deepcopy(raw_tata)
    raw.schedules["balance-sheet"].pop("Other Assets")
    with pytest.raises(DataValidationError):
        build_financials(raw)


def test_missing_expense_schedule_warns(raw_tata):
    raw = copy.deepcopy(raw_tata)
    raw.schedules["profit-loss"].pop("Expenses")
    f = build_financials(raw)
    assert all(c == 0 for c in f.lines["cogs"])
    assert any("no expense breakdown" in w for w in f.warnings)


def test_missing_price_stops_run(raw_tata):
    raw = copy.deepcopy(raw_tata)
    raw.top_ratios.pop("Current Price")
    with pytest.raises(DataValidationError):
        build_financials(raw)


@pytest.mark.parametrize("name,kind", [("Material Cost", "COGS"), ("Power and Fuel", "COGS"), ("Other Mfr. Exp", "COGS"),
                                       ("Employee Cost", "OPEX"), ("Selling and admin", "OPEX"), ("Other Expenses", "OPEX")])
def test_expense_classification(name, kind):
    assert classify_expense(name) == kind
