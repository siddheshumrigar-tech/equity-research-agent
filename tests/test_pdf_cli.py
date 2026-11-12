import json
import sys
from pathlib import Path

import pytest
from pypdf import PdfReader

ROOT = Path(__file__).resolve().parents[1]


def _pdf_text(path):
    r = PdfReader(path)
    return len(r.pages), "\n".join(p.extract_text() or "" for p in r.pages)


def test_pdf_uses_engine_numbers(built_asian):
    pages, text = _pdf_text(built_asian.pdf_path)
    v, f = built_asian.valuation, built_asian.fin
    assert 5 <= pages <= 14
    assert f"{v.target_price:,.0f}" in text
    assert v.rating in text
    assert "DRAFT" in text  # no narrative yet
    assert f"{f.latest('sales'):,.0f}" in text
    assert f.name in text
    for bad in ("Tata", "Pillar 1", "Porter", "Automotive Eng", "25000", "Hermes"):
        assert bad not in text


def test_pdf_tatatech(built_tata):
    pages, text = _pdf_text(built_tata.pdf_path)
    assert f"{built_tata.valuation.dcf_value:,.2f}" in text
    assert "Screener.in" in text


def test_front_page_fits_and_matches_summary(built_tata, built_asian):
    for b in (built_tata, built_asian):
        r = PdfReader(b.pdf_path)
        p1 = r.pages[0].extract_text()
        p2 = r.pages[1].extract_text()
        v = b.valuation
        for s in ("Key financials and valuation", "Stock data", "Expected total", "Risk-reward", v.rating,
                  f"{v.target_price:,.0f}"):
            assert s in p1, s
        assert "Key financials and valuation" not in p2
        data = json.loads(Path(b.summary_path).read_text(encoding="utf-8"))
        assert data["valuation"]["rating"] == v.rating
        assert f"{data['valuation']['target_price']:,.0f}" in p1


def test_exhibits_numbered_in_order(built_tata):
    _, text = _pdf_text(built_tata.pdf_path)
    import re

    nums = [int(n) for n in re.findall(r"Exhibit (\d+):", text)]
    assert nums == list(range(1, len(nums) + 1)) and len(nums) >= 10
    assert len(built_tata.report_exhibits) == len(nums)


def test_contrast_helper():
    from equity_agent.report.styles import contrast, readable_on_white

    for hex_ in ("FFD700", "D8B15F", "70AD47", "003366"):
        assert contrast(readable_on_white(hex_)) >= 4.5
    assert readable_on_white("003366") == "003366"


def test_summary_json(built_tata):
    data = json.loads(Path(built_tata.summary_path).read_text(encoding="utf-8"))
    assert data["status"] == "ok"
    assert data["symbol"] == "TATATECH"
    assert data["valuation"]["verdict"] == built_tata.valuation.verdict == data["valuation"]["rating"]
    assert data["valuation"]["expected_total_return"] == pytest.approx(built_tata.valuation.expected_total_return, abs=1e-4)
    assert {f["code"] for f in data["flags"]} == {f.code for f in built_tata.flags}
    assert data["shareholding_latest"]["Promoters"] is not None
    assert data["price_performance"]["absolute"]["12M"] is not None
    assert data["technicals"]["sma200"] is not None
    assert len(data["peers"]) >= 2
    assert Path(data["files"]["excel"]).exists() and Path(data["files"]["pdf"]).exists()
    assert Path(data["files"]["research_pack"]).exists()


def _cli(args):
    sys.path.insert(0, str(ROOT))
    import generate_equity_report

    return generate_equity_report.main(args)


def test_cli_offline_success(tmp_path, fixtures_dir):
    code = _cli(["--ticker", "ASIANPAINT", "--offline", "--cache-dir", str(fixtures_dir), "--output-dir", str(tmp_path),
                 "--no-pdf"])
    assert code == 0
    assert (tmp_path / "ASIANPAINT_Financial_Model.xlsx").exists()
    s = json.loads((tmp_path / "ASIANPAINT_summary.json").read_text(encoding="utf-8"))
    assert s["status"] == "ok" and s["files"]["pdf"] is None


def test_cli_bank_exit_code(tmp_path, fixtures_dir):
    code = _cli(["--ticker", "HDFCBANK", "--offline", "--cache-dir", str(fixtures_dir), "--output-dir", str(tmp_path)])
    assert code == 3
    s = json.loads((tmp_path / "HDFCBANK_summary.json").read_text(encoding="utf-8"))
    assert s["status"] == "error" and s["error_type"] == "BankNotSupported"


def test_cli_offline_cache_miss_exit_code(tmp_path, fixtures_dir):
    code = _cli(["--ticker", "NOSUCHCO", "--offline", "--cache-dir", str(fixtures_dir), "--output-dir", str(tmp_path)])
    assert code == 5


def test_cli_symbol_not_found(tmp_path, monkeypatch):
    import requests

    class R:
        status_code = 404
        text = ""
        content = b""

    monkeypatch.setattr(requests.Session, "get", lambda self, url, params=None, timeout=None: R())
    code = _cli(["--ticker", "NOSUCHCO", "--cache-dir", str(tmp_path / "c"), "--output-dir", str(tmp_path)])
    assert code == 2
