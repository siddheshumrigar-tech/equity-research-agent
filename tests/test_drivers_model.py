import statistics

import pytest

from equity_agent.drivers import GROWTH_OFFSET, MARGIN_OFFSET, build_drivers
from equity_agent.model_engine import actual_balance_check, run_forecast


def test_drivers_are_history_averages(fin_tata):
    d = build_drivers(fin_tata)
    L = fin_tata.lines
    assert d.value("sales_growth") == pytest.approx((5506 / 4414) ** (1 / 3) - 1)
    exp_dso = statistics.mean(L["receivables"][k] / L["sales"][k] * 365 for k in (-3, -2, -1))
    assert d.value("dso") == pytest.approx(exp_dso)
    exp_margin = statistics.mean(L["ebitda"][k] / L["sales"][k] for k in (-3, -2, -1))
    assert d.value("ebitda_margin") == pytest.approx(exp_margin)
    assert d.value("dio") == 0  # services company: no inventory
    assert d.value("dividend_payout") == pytest.approx(statistics.mean([0.60, 0.70, 0.87]))
    assert d.window_years == [2024, 2025, 2026]


def test_payable_days_basis(fin_tata, fin_asian):
    # Tata Technologies: material cost is ~21% of expenses -> service cost structure, so COGS includes employee
    # cost and payables are measured on that COGS
    d = build_drivers(fin_tata)
    L = fin_tata.lines
    assert fin_tata.cost_basis == "services"
    assert d.dpo_basis == "cogs"
    assert d.value("dpo") == pytest.approx(statistics.mean(L["trade_payables"][k] / L["cogs"][k] * 365 for k in (-3, -2, -1)))
    assert d.value("dpo") < 120
    # Asian Paints: manufacturer -> COGS basis (Screener reports ~79 days)
    da = build_drivers(fin_asian)
    assert da.dpo_basis == "cogs"
    assert 60 < da.value("dpo") < 100


def test_scenario_offsets(fin_asian):
    d = build_drivers(fin_asian)
    g = d.items["sales_growth"]
    m = d.items["ebitda_margin"]
    assert g.bull == pytest.approx(g.base + GROWTH_OFFSET)
    assert g.bear == pytest.approx(g.base - GROWTH_OFFSET)
    assert m.bull == pytest.approx(m.base + MARGIN_OFFSET)
    cx = d.items["capex_pct"]
    assert cx.bull == pytest.approx(cx.base * 0.9) and cx.bear == pytest.approx(cx.base * 1.1)
    assert d.items["ronic_share"].bull > d.items["ronic_share"].base > d.items["ronic_share"].bear


def test_minority_share_forecast(fin_asian):
    d = build_drivers(fin_asian)
    L = fin_asian.lines
    exp = statistics.mean(L["minority_other"][k] / (L["pbt"][k] - L["tax"][k]) for k in (-3, -2, -1))
    assert d.value("minority_pct") == pytest.approx(exp)
    fc = run_forecast(fin_asian, d, "base")
    for p, t, m, n in zip(fc.lines["pbt"], fc.lines["tax"], fc.lines["minority_other"], fc.lines["net_profit"]):
        assert m == pytest.approx((p - t) * d.value("minority_pct"))
        assert n == pytest.approx(p - t - m)


def test_loans_flat_and_outside_nwc(fin_asian):
    fc = run_forecast(fin_asian, build_drivers(fin_asian), "base")
    assert all(x == fin_asian.lines["loans_advances"][-1] for x in fc.lines["loans_advances"])
    F = fc.lines
    for k in range(5):
        assert F["nwc"][k] == pytest.approx(F["receivables"][k] + F["inventory"][k] + F["other_ca"][k]
                                            - F["trade_payables"][k] - F["other_liab"][k])


def test_cwip_capitalised(fin_asian):
    fc = run_forecast(fin_asian, build_drivers(fin_asian), "base", cwip_years=2)
    c0 = fin_asian.lines["cwip"][-1]
    assert fc.lines["cwip_transfer"][0] == pytest.approx(c0 / 2)
    assert fc.lines["cwip"][1] == pytest.approx(0.0, abs=1e-9)
    assert fc.lines["cwip"][-1] == pytest.approx(0.0, abs=1e-9)


def test_working_capital_glide(fin_tata):
    d = build_drivers(fin_tata)
    L = fin_tata.lines
    base_dso = L["receivables"][-1] / L["sales"][-1] * 365
    assert d.value("dso", "base", 0) == pytest.approx(base_dso + (d.value("dso") - base_dso) * 0.5)
    assert d.value("dso", "base", 1) == pytest.approx(d.value("dso"))
    fc = run_forecast(fin_tata, d, "base")
    assert fc.lines["receivables"][0] == pytest.approx(fc.lines["sales"][0] * d.value("dso", "base", 0) / 365)


def test_override_keeps_scenario_offset(fin_tata):
    from equity_agent.drivers import Override

    d = build_drivers(fin_tata)
    d.overrides["sales_growth"] = {0: Override(0.20, "test")}
    assert d.value("sales_growth", "base", 0) == pytest.approx(0.20)
    assert d.value("sales_growth", "bull", 0) == pytest.approx(0.20 + GROWTH_OFFSET)
    assert d.value("sales_growth", "base", 1) == pytest.approx(d.items["sales_growth"].base)


@pytest.mark.parametrize("scenario", ["base", "bull", "bear"])
def test_balance_sheet_balances_without_plug(fin_tata, fin_asian, scenario):
    for f in (fin_tata, fin_asian):
        fc = run_forecast(f, build_drivers(f), scenario)
        assert all(abs(x) < 1e-6 for x in fc.lines["balance_check"])
    assert all(abs(x) < 1e-6 for x in actual_balance_check(fin_tata))


def test_scenarios_change_forecast_only(fin_asian):
    d = build_drivers(fin_asian)
    actual_before = list(fin_asian.lines["sales"])
    base = run_forecast(fin_asian, d, "base")
    bull = run_forecast(fin_asian, d, "bull")
    assert fin_asian.lines["sales"] == actual_before
    assert bull.lines["sales"][0] > base.lines["sales"][0]
    assert bull.lines["sales"][0] == pytest.approx(fin_asian.lines["sales"][-1] * (1 + d.items["sales_growth"].bull))


def test_cash_comes_from_cash_flow(fin_tata):
    fc = run_forecast(fin_tata, build_drivers(fin_tata), "base")
    prev = fin_tata.lines["cash"][-1]
    for cash, ncf in zip(fc.lines["cash"], fc.lines["net_cash_flow"]):
        assert cash == pytest.approx(prev + ncf)
        prev = cash
