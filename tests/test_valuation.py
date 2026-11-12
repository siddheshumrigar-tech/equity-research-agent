import copy
import math
from datetime import date

import pytest

from equity_agent.config import load_market_config
from equity_agent.drivers import build_drivers
from equity_agent.market_data import MarketData, blume, compute_beta
from equity_agent.model_engine import run_forecast
from equity_agent.valuation import (DCFInputs, blend, capm, compute_valuation, dcf_inputs, dcf_per_share, gordon_tv,
                                    historical_pe, mid_year_factor, ntm_weight, rating_for, reverse_dcf_growth,
                                    ronic_rule, terminal_value, three_stage, wacc_formula)


def test_known_answers():
    assert capm(0.0535, 1.0, 0.0708) == pytest.approx(0.1243)
    assert wacc_formula(0.12, 0.06, 80, 20) == pytest.approx(0.108)
    assert mid_year_factor(0.10, 1) == pytest.approx(1 / math.sqrt(1.10))
    assert gordon_tv(100.0, 0.045, 0.115) == pytest.approx(1492.857, rel=1e-6)
    # value-driver terminal value: 100 x 1.05 x (1 - 0.05/0.20) / (0.12 - 0.05)
    assert terminal_value(100.0, 0.05, 0.12, 0.20) == pytest.approx(1125.0)
    assert blume(1.3) == pytest.approx(1.201)


def test_beta_regression():
    idx = [100, 101, 99, 102, 104, 103, 106]
    stock = [100 * (1 + 2 * (idx[i] / idx[0] - 1)) for i in range(len(idx))]
    assert compute_beta(idx, idx) == pytest.approx(1.0)
    assert compute_beta(stock, idx) > 1.5


def test_blume_applied(res_tata):
    md = res_tata.market
    assert md.beta_raw is not None
    assert md.beta == pytest.approx(min(max(0.67 * md.beta_raw + 0.33, 0.5), 2.0))


@pytest.mark.parametrize("etr,rating", [(0.25, "BUY"), (0.15, "BUY"), (0.149, "ADD"), (0.05, "ADD"), (0.0, "REDUCE"),
                                        (-0.05, "REDUCE"), (-0.051, "SELL"), (-0.4, "SELL")])
def test_rating_thresholds(etr, rating):
    assert rating_for(etr, load_market_config()) == rating


def test_ronic_rule():
    assert ronic_rule(0.30, 0.12, 0.5) == pytest.approx(0.21)
    assert ronic_rule(0.08, 0.12, 0.5) == pytest.approx(0.12)  # floor at WACC
    assert ronic_rule(0.30, 0.12, 1.0) == pytest.approx(0.30)  # cap at ROIC


def _inp(**kw):
    base = dict(fcff=[100.0] * 5, sales_last=1000.0, growth_last=0.0, margin_last=0.1, tax_last=0.0, nopat_last=100.0,
                ebitda_margin_last=0.12, roic_raw=0.12, roic_cap=1.0, n_fade=5)
    base.update(kw)
    return DCFInputs(**base)


def test_zero_growth_ronic_wacc_is_perpetuity():
    """g = 0 and RONIC = WACC: no reinvestment, value = NOPAT / WACC from year 1."""
    w = 0.10
    r = three_stage(_inp(), w, 0.0, 0.5, ronic_override=w)
    assert all(x == pytest.approx(100.0) for x in r.fcff)
    assert r.tv == pytest.approx(100.0 / w)
    # all cash flows 100, mid-year discounting: sum = 100/w x (1+w)^0.5
    assert r.ev == pytest.approx(100 / w * (1 + w) ** 0.5, rel=1e-9)


def test_fade_reinvestment_math():
    inp = _inp(growth_last=0.10, margin_last=0.2, tax_last=0.25, roic_raw=0.30)
    r = three_stage(inp, 0.12, 0.05, 0.5)
    assert r.roic == pytest.approx(0.30)
    assert r.ronic == pytest.approx(0.21)
    assert r.growth_fade[0] == pytest.approx(0.10 + (0.05 - 0.10) / 5)
    assert r.growth_fade[-1] == pytest.approx(0.05)
    assert r.roic_fade[-1] == pytest.approx(0.21)
    nopat1 = 1000 * (1 + r.growth_fade[0]) * 0.2 * 0.75
    assert r.nopat_fade[0] == pytest.approx(nopat1)
    assert r.reinvestment_fade[0] == pytest.approx(nopat1 * r.growth_fade[0] / r.roic_fade[0])
    assert r.fcff[5] == pytest.approx(nopat1 - r.reinvestment_fade[0])
    assert r.tv == pytest.approx(r.nopat_fade[-1] * 1.05 * (1 - 0.05 / 0.21) / 0.07)


def test_ntm_weight():
    class F:
        years = [2026]
        period_labels = ["Mar 2026"]
    assert ntm_weight(F, date(2026, 3, 31)) == pytest.approx(1.0)
    assert ntm_weight(F, date(2026, 9, 30)) == pytest.approx(182 / 365)
    assert ntm_weight(F, date(2027, 6, 1)) == 0.0


def test_blend_renormalises():
    v, used = blend({"dcf": 100.0, "pe": None, "ev_ebitda": 200.0}, {"dcf": 0.5, "pe": 0.25, "ev_ebitda": 0.25})
    assert used == pytest.approx({"dcf": 2 / 3, "ev_ebitda": 1 / 3})
    assert v == pytest.approx(100 * 2 / 3 + 200 / 3)


def test_valuation_consistency(res_tata):
    v = res_tata.valuation
    assert v.ke == pytest.approx(v.rf + v.beta * v.erp)
    assert v.wacc == pytest.approx(v.ke * v.weight_equity + v.kd_post * v.weight_debt)
    assert len(v.fcff) == 10 and v.stage[:5] == ["explicit"] * 5
    assert v.enterprise_value == pytest.approx(v.sum_pv_explicit + v.sum_pv_fade + v.pv_terminal_value)
    assert v.equity_value == pytest.approx(v.enterprise_value + v.cash + v.investments - v.borrowings - v.minority_interest)
    assert v.dcf_value == pytest.approx(v.equity_value / v.shares_cr)
    assert v.dcf_value_12m == pytest.approx(v.dcf_value * (1 + v.ke - v.dividend_yield))
    assert v.sensitivity_grid[2][2] == pytest.approx(v.dcf_value)
    assert v.sens_gm_grid[2][2] == pytest.approx(v.dcf_value)
    assert v.tv_share == pytest.approx(v.pv_terminal_value / v.enterprise_value)
    expected, _ = blend(v.method_values, load_market_config()["target_weights"])
    assert v.target_price == pytest.approx(expected)
    assert v.upside == pytest.approx(v.target_price / v.price - 1)
    assert v.expected_total_return == pytest.approx(v.upside + v.dividend_yield)
    assert v.rating == rating_for(v.expected_total_return, load_market_config())
    assert v.verdict == v.rating


def test_relative_methods(res_tata):
    v, fc = res_tata.valuation, res_tata.forecast
    w = v.ntm_weight
    eps = (w * fc.lines["net_profit"][0] + (1 - w) * fc.lines["net_profit"][1]) / v.shares_cr
    assert v.ntm_eps == pytest.approx(eps)
    assert v.pe_value == pytest.approx(v.pe_mean * eps)
    assert v.pe_low < v.pe_value < v.pe_high
    assert v.peer_ev_ebitda == pytest.approx(res_tata.peers.median("ev_ebitda"))
    bridge = v.cash + v.investments - v.borrowings - v.minority_interest
    assert v.ev_ebitda_value == pytest.approx((v.peer_ev_ebitda * v.ntm_ebitda + bridge) / v.shares_cr)


def test_historical_pe_uses_traded_prices_and_year_share_count(fin_tata, res_tata):
    md = res_tata.market
    assert md.month_end_prices_unadj
    pe = historical_pe(fin_tata, md)
    k = fin_tata.n - 1
    price = md.month_end_prices_unadj[f"{fin_tata.years[k]}-03"]
    eps = fin_tata.eps[k]
    assert pe[f"FY{fin_tata.years[k] % 100:02d}"] == pytest.approx(price / eps, rel=1e-6)


def test_three_stage_lowers_tv_share(res_tata, res_asian):
    for r in (res_tata, res_asian):
        assert r.valuation.tv_share < 0.75


def test_scenarios_ordered(res_tata, res_asian):
    for r in (res_tata, res_asian):
        s = r.valuation.scenarios
        assert s["bear"].target <= s["base"].target <= s["bull"].target
        assert s["base"].target == pytest.approx(r.valuation.target_price)
        assert r.valuation.prob_weighted_value == pytest.approx(
            0.25 * s["bull"].target + 0.5 * s["base"].target + 0.25 * s["bear"].target)


def _value_at_growth(res, growth):
    v, f = res.valuation, res.fin
    d2 = copy.deepcopy(res.drivers)
    for attr in ("base", "bull", "bear"):
        setattr(d2.items["sales_growth"], attr, growth)
    fc = run_forecast(f, d2, "base")
    bridge = v.cash + v.investments - v.borrowings - v.minority_interest
    cfg = load_market_config()
    return dcf_per_share(dcf_inputs(fc, d2, cfg), v.wacc, v.terminal_growth, d2.value("ronic_share"), bridge, v.shares_cr)


def test_reverse_dcf_solves_reachable_price(res_asian):
    v, f = res_asian.valuation, res_asian.fin
    price = _value_at_growth(res_asian, 0.10)
    g = reverse_dcf_growth(f, res_asian.drivers, "base", v.wacc, v.terminal_growth, price)
    assert g == pytest.approx(0.10, abs=1e-4)


def test_reverse_dcf_matches_market(res_tata):
    v = res_tata.valuation
    if v.reverse_dcf_growth is not None:
        assert _value_at_growth(res_tata, v.reverse_dcf_growth) == pytest.approx(res_tata.fin.price, rel=1e-3)


def test_negative_value_is_not_floored(fin_tata):
    d = build_drivers(fin_tata)
    fc = run_forecast(fin_tata, d, "base")
    fc.lines["fcff"] = [-5000.0] * len(fc.years)
    fc.lines["ebit_operating"] = [-5000.0] * len(fc.years)
    fc.lines["nopat"] = [-5000.0] * len(fc.years)
    v = compute_valuation(fin_tata, d, fc, MarketData(beta=1.0, beta_source="test"), compute_reverse=False,
                          compute_extras=False)
    assert v.dcf_value < 0
    assert any("negative" in w for w in v.warnings)


def test_terminal_growth_from_config(res_tata):
    assert res_tata.valuation.terminal_growth == load_market_config()["terminal_growth"]
