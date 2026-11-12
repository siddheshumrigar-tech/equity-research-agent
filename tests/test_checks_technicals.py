import copy
from datetime import date

import pytest

from equity_agent.checks import run_checks
from equity_agent.config import load_market_config
from equity_agent.technicals import compute_technicals, period_return, rsi, sma


def _codes(flags):
    return {f.code: f.severity for f in flags}


def _check(res, **changes):
    v = copy.deepcopy(res.valuation)
    for k, x in changes.items():
        setattr(v, k, x)
    return _codes(run_checks(res.fin, res.drivers, v, res.market, res.cfg, valuation_date=date(2026, 11, 11)))


def test_tv_share_flag(res_tata):
    assert _check(res_tata, tv_share=0.80)["TV_SHARE_HIGH"] == "high"
    assert "TV_SHARE_HIGH" not in _check(res_tata, tv_share=0.50)


def test_divergence_flag(res_tata):
    c = _check(res_tata, dcf_value_12m=100.0, pe_value=300.0, ev_ebitda_value=300.0)
    assert c["DCF_RELATIVE_DIVERGENCE"] == "high"
    assert "DCF_RELATIVE_DIVERGENCE" not in _check(res_tata, dcf_value_12m=250.0, pe_value=300.0, ev_ebitda_value=300.0)


def test_extreme_upside_and_negative_fcff(res_tata):
    assert _check(res_tata, upside=-0.58)["EXTREME_UPSIDE"] == "high"
    fcff = list(res_tata.valuation.fcff)
    fcff[-1] = -1.0
    assert _check(res_tata, fcff=fcff)["NEGATIVE_FCFF"] == "high"


def test_reverse_dcf_and_staleness(res_tata):
    assert _check(res_tata, reverse_dcf_growth=None)["REVERSE_DCF_UNREACHABLE"] == "medium"
    cfg = dict(load_market_config(), as_of="2026-01-01")
    flags = run_checks(res_tata.fin, res_tata.drivers, res_tata.valuation, res_tata.market, cfg,
                       valuation_date=date(2026, 11, 11))
    assert _codes(flags)["STALE_MARKET_INPUTS"] == "medium"


def test_clamped_driver_flag(res_tata):
    d = copy.deepcopy(res_tata.drivers)
    it = d.items["interest_rate"]
    it.clamped, it.raw = True, 1.5655
    flags = run_checks(res_tata.fin, d, res_tata.valuation, res_tata.market, res_tata.cfg)
    assert any(f.code == "DRIVER_CLAMPED" and "1.5655" in f.message for f in flags)


def test_ytd_growth(fin_tata):
    n, g = fin_tata.ytd_growth("sales")
    assert n >= 1 and g is not None
    labels = fin_tata.quarter_labels
    fy1 = fin_tata.years[-1] + 1
    cur = [lab for lab in labels if fin_tata.quarter_fy(lab) == fy1]
    assert len(cur) == n


def test_shareholding_parsed(fin_tata, fin_asian):
    for f in (fin_tata, fin_asian):
        assert "Promoters" in f.shareholding and f.shareholding_labels
        last = [f.shareholding[k][-1] for k in f.shareholding if f.shareholding[k][-1] is not None]
        assert sum(last) == pytest.approx(100, abs=0.5)


def test_technical_indicators():
    s = [float(x) for x in range(1, 301)]
    assert sma(s, 50)[-1] == pytest.approx(sum(s[-50:]) / 50)
    assert sma(s, 50)[48] is None
    assert rsi(s)[-1] == pytest.approx(100.0)
    assert period_return(s, 21) == pytest.approx(300 / 279 - 1)
    zig = [100 + (i % 20 if (i // 20) % 2 == 0 else 20 - i % 20) for i in range(300)]
    t = compute_technicals([f"d{i}" for i in range(300)], zig, [100.0] * 300)
    assert t is not None and t.high_52w == max(zig[-252:]) and 0 <= t.rsi14 <= 100
    assert all(x > t.last_close for x in t.resistances) and all(x < t.last_close for x in t.supports)


def test_technicals_on_fixture(res_tata):
    t = res_tata.technicals
    assert t is not None and t.sma200 is not None
    assert t.returns["12M"] is not None and "12M" in t.relative
