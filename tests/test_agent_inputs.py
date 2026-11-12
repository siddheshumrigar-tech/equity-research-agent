import copy
import json
import sys
from pathlib import Path

import pytest

from equity_agent.analyst_inputs import parse_analyst_inputs, validate_analyst_inputs
from equity_agent.errors import InputValidationError
from equity_agent.narrative import build_tokens, lint_numbers, validate_narrative

ROOT = Path(__file__).resolve().parents[1]
AGENT = ROOT / "tests" / "fixtures" / "agent"
SRC = {"type": "concall", "title": "Q2 call", "url": "https://x", "date": "2026-10-20"}


def _ai():
    return json.loads((AGENT / "TATATECH_analyst_inputs.json").read_text(encoding="utf-8"))


def _nar():
    return json.loads((AGENT / "TATATECH_narrative.json").read_text(encoding="utf-8"))


def test_analyst_input_without_source_rejected():
    d = _ai()
    del d["driver_overrides"][0]["source"]
    errs = validate_analyst_inputs(d, "TATATECH")
    assert any("source" in e for e in errs)
    with pytest.raises(InputValidationError):
        parse_analyst_inputs(d, "TATATECH")


def test_analyst_input_range_and_symbol():
    d = _ai()
    d["driver_overrides"][0]["years"]["FY27"] = 12  # percent typed as a whole number
    assert any("outside the allowed range" in e for e in validate_analyst_inputs(d, "TATATECH"))
    assert any("does not match" in e for e in validate_analyst_inputs(_ai(), "TCS"))


@pytest.fixture(scope="module")
def res_tata_ai(fixtures_dir):
    from equity_agent.pipeline import analyse

    return analyse("TATATECH", cache_dir=fixtures_dir, offline=True, analyst_inputs_path=AGENT / "TATATECH_analyst_inputs.json")


def test_overrides_applied(res_tata_ai, res_tata):
    d, fc = res_tata_ai.drivers, res_tata_ai.forecast
    assert d.value("sales_growth", "base", 0) == pytest.approx(0.12)
    assert d.value("sales_growth", "base", 1) == pytest.approx(0.11)
    assert d.value("sales_growth", "base", 2) == pytest.approx(d.items["sales_growth"].base)
    assert fc.lines["sales"][0] == pytest.approx(res_tata.fin.lines["sales"][-1] * 1.12)
    assert res_tata_ai.valuation.ronic == pytest.approx(0.18)
    assert res_tata_ai.valuation.ronic_source == "Analyst input"
    items = [e["item"] for e in res_tata_ai.analyst.entries()]
    assert "Driver: sales_growth FY27E" in items and "Terminal RONIC" in items


def test_sotp_replaces_blend(fixtures_dir):
    from equity_agent.analyst_inputs import parse_analyst_inputs
    from equity_agent.pipeline import analyse

    res = analyse("TATATECH", cache_dir=fixtures_dir, offline=True)
    d = {"symbol": "TATATECH", "sotp": {"segments": [
        {"segment": "A", "method": "EV/EBITDA", "metric_value_cr": 1000, "multiple": 20, "source": SRC},
        {"segment": "B", "method": "Stake value", "value_cr": 5000, "source": SRC}]}}
    ai = parse_analyst_inputs(d, "TATATECH")
    from equity_agent.valuation import compute_valuation

    v = compute_valuation(res.fin, res.drivers, res.forecast, res.market, res.cfg, peers=res.peers, analyst=ai,
                          compute_reverse=False, compute_extras=False)
    bridge = v.cash + v.investments - v.borrowings - v.minority_interest
    assert v.target_method == "sotp"
    assert v.target_price == pytest.approx((25000 + bridge) / v.shares_cr)


def test_excel_marks_overrides(res_tata_ai, tmp_path):
    import openpyxl

    from equity_agent.excel_writer import write_model

    r = res_tata_ai
    p = write_model(tmp_path / "ai.xlsx", r.fin, r.drivers, r.forecast, r.valuation, r.palette, r.cfg, r.market,
                    peers=r.peers, analyst=r.analyst, flags=r.flags)
    wb = openpyxl.load_workbook(p)
    ws = wb["Drivers"]
    cell = ws["G7"]  # sales growth, first forecast year
    assert cell.comment is not None and "Source:" in cell.comment.text
    assert cell.fill.start_color.rgb.endswith("FCE4D6")
    assert ws["I7"].comment is None
    ai_ws = wb["Analyst Inputs"]
    assert any("SAMPLE Q2 FY27 earnings call" in str(ai_ws.cell(rr, 5).value) for rr in range(5, 15))


def test_narrative_validation(res_tata):
    tokens = build_tokens(res_tata)
    data = _nar()
    errs, lint = validate_narrative(data, tokens, ["DCF_RELATIVE_DIVERGENCE"], "TATATECH")
    assert errs == []
    assert any("450" in w for w in lint)
    bad = copy.deepcopy(data)
    bad["thesis_points"][0]["text"] += " {{no_such_token}} [9]"
    del bad["flag_responses"]["DCF_RELATIVE_DIVERGENCE"]
    errs, _ = validate_narrative(bad, tokens, ["DCF_RELATIVE_DIVERGENCE"], "TATATECH")
    joined = " ".join(errs)
    assert "no_such_token" in joined and "[9]" in joined and "flag_responses.DCF_RELATIVE_DIVERGENCE" in joined


def test_lint_ignores_cited_and_years():
    assert lint_numbers("x", "Market share rose to 23% in 2025 [1].") == []
    assert lint_numbers("x", "Revenue for FY26 grew.") == []
    assert lint_numbers("x", "Volumes rose 23% last year.")


def _cli(args):
    sys.path.insert(0, str(ROOT))
    import generate_equity_report

    return generate_equity_report.main(args)


def test_two_step_render(tmp_path, fixtures_dir):
    from pypdf import PdfReader

    base = ["--ticker", "TATATECH", "--offline", "--cache-dir", str(fixtures_dir), "--output-dir", str(tmp_path)]
    assert _cli(base) == 0
    pack = json.loads((tmp_path / "TATATECH_research_pack.json").read_text(encoding="utf-8"))
    assert "narrative_schema" in pack and "tokens" in pack and "flags_to_answer" in pack
    s = json.loads((tmp_path / "TATATECH_summary.json").read_text(encoding="utf-8"))
    assert s["report_stage"] == "draft"
    # render without a narrative fails with exit code 6
    assert _cli(base + ["--render"]) == 6
    (tmp_path / "TATATECH_narrative.json").write_text((AGENT / "TATATECH_narrative.json").read_text(encoding="utf-8"),
                                                      encoding="utf-8")
    (tmp_path / "TATATECH_analyst_inputs.json").write_text(
        (AGENT / "TATATECH_analyst_inputs.json").read_text(encoding="utf-8"), encoding="utf-8")
    assert _cli(base + ["--render"]) == 0
    s = json.loads((tmp_path / "TATATECH_summary.json").read_text(encoding="utf-8"))
    assert s["report_stage"] == "final" and s["high_flags_unanswered"] == []
    assert s["analyst_inputs"].endswith("TATATECH_analyst_inputs.json")
    text = "\n".join(p.extract_text() or "" for p in PdfReader(tmp_path / "TATATECH_Equity_Research_Report.pdf").pages)
    assert "Test fixture: growth recovery priced in" in text
    assert "DRAFT" not in text
    # narrative sections appear in the final report, in the spec order
    order = ["Investment thesis", "Company overview", "Industry and competitive position", "Porter's five forces",
             "SWOT analysis", "PESTLE and ESG", "Quarterly results", "Valuation football field", "Technical view",
             "Risks and catalysts", "Catalyst timeline", "Corporate actions", "Financial statements"]
    pos = [text.find(s) for s in order]
    assert all(p >= 0 for p in pos), [s for s, p in zip(order, pos) if p < 0]
    assert pos == sorted(pos)
    assert "Segment share of revenue" in text and "Sample summary one for page 1" in text
    assert any("Narrative lint" in w for w in s["warnings"])


def test_catalyst_dates():
    from datetime import datetime

    from equity_agent.report.sections.research import catalyst_date

    assert catalyst_date("2027-01-20") == datetime(2027, 1, 20)
    assert catalyst_date("2027-02") == datetime(2027, 2, 1)
    assert catalyst_date("FY28") == datetime(2028, 3, 31)
    assert catalyst_date("Q1 FY28") == datetime(2027, 6, 30)
    assert catalyst_date("Q4 FY27") == datetime(2027, 3, 31)
    assert catalyst_date("soon") is None
