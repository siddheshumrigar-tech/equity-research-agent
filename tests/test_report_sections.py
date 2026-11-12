"""Task 13: data-driven report sections (story in charts, quarters, peers, forecasts vs history, statements)."""
from pypdf import PdfReader


def _text(path):
    return "\n".join(p.extract_text() or "" for p in PdfReader(path).pages)


def test_new_sections_present(built_asian):
    text = _text(built_asian.pdf_path)
    for title in ("Story in charts", "Quarterly results", "Peer comparison", "Ratios vs peers", "Forecasts vs history",
                  "Financial statements", "Income statement", "Balance sheet", "Cash flow statement", "Key ratios"):
        assert title in text, title
    # sections follow the order in the design spec
    order = [text.index(s) for s in ("Story in charts", "Quarterly results", "Peer comparison", "Forecasts vs history",
                                     "10. Valuation", "Financial statements")]
    assert order == sorted(order)


def test_quarter_growth_matches_reported(fin_asian):
    from equity_agent.report.sections.quarterly import last_quarters

    q = last_quarters(fin_asian, 8)
    assert len(q["labels"]) == 8
    labs = fin_asian.quarter_labels
    i = labs.index(q["labels"][-1])
    mon, yr = q["labels"][-1].split()
    j = labs.index(f"{mon} {int(yr) - 1}")
    s = fin_asian.quarters["sales"]
    assert abs(q["sales_yoy"][-1] - (s[i] / s[j] - 1)) < 1e-12
    assert abs(q["sales_qoq"][-1] - (s[i] / s[i - 1] - 1)) < 1e-12


def test_peer_comment():
    from equity_agent.report.sections.peers import ratio_comment

    assert ratio_comment("roe", "pctpt", "quality", 21.8, [17.3, 8.9, 19.1]) == (
        "4.5 pp above the peer median of 17.3%; the highest of 4.")
    assert ratio_comment("pe", "x", "multiple", 30.0, [20.0, 40.0, 25.0]) == (
        "A 20% premium to the peer median of 25.0x; the 2nd most expensive of 4.")
    assert "Not meaningful" in ratio_comment("pe", "x", "multiple", 500.0, [20.0, 30.0])


def test_statements_balance(res_asian):
    from equity_agent.report.sections import statements  # noqa: F401  (module imports cleanly)

    fc = res_asian.forecast
    for a, l in zip(fc.lines["total_assets"], fc.lines["total_liabilities"]):
        assert abs(a - l) < 1e-6


def test_peer_history_definitions(fin_asian):
    from equity_agent.peers import subject_history

    h = subject_history(fin_asian)
    L = fin_asian.lines
    y = f"FY{fin_asian.years[-1] % 100:02d}"
    eq = (L["total_equity"][-2] + L["total_equity"][-1]) / 2
    cap = (L["total_equity"][-2] + L["borrowings"][-2] + L["total_equity"][-1] + L["borrowings"][-1]) / 2
    assert abs(h["roe"][y] - L["net_profit"][-1] / eq) < 1e-12
    assert abs(h["roce"][y] - (L["pbt"][-1] + L["interest"][-1]) / cap) < 1e-12
    assert abs(h["ebitda_margin"][y] - L["ebitda"][-1] / L["sales"][-1]) < 1e-12
    assert abs(h["sales_growth"][y] - (L["sales"][-1] / L["sales"][-2] - 1)) < 1e-12
    assert len(h["roe"]) == 5


def test_peer_history_in_fixture_and_pdf(res_asian, built_asian):
    ps = res_asian.peers
    assert ps.has_history and len(ps.history_years()) == 5
    y = ps.history_years()[-1]
    vals = sorted(p.history["ebitda_margin"][y] for p in ps.peers if p.history["ebitda_margin"].get(y) is not None)
    assert ps.history_median("ebitda_margin", y) == vals[len(vals) // 2] if len(vals) % 2 else True
    text = _text(built_asian.pdf_path)
    # header cells wrap, so check phrases that stay on one line
    assert "5-year ratios vs the peer median" in text and "Ratio history: company, peer median" in text


def test_years_above():
    from equity_agent.report.sections.peers import years_above

    assert years_above([0.2, 0.1, None, 0.3], [0.1, 0.2, 0.1, 0.25]) == (2, 3)


def test_peer_names_escaped():
    from equity_agent.report.sections.peers import _short

    assert _short("L&T Technology Services Ltd") == "L&amp;T Technology Services"


def test_technical_section_in_pdf(built_asian):
    text = _text(built_asian.pdf_path)
    assert "Technical view" in text and "Technical levels" in text
    assert text.index("10. Valuation") < text.index("Technical view") < text.index("Financial statements")


def test_technical_commentary():
    from equity_agent.report.sections.technical import rsi_zone, technical_commentary
    from equity_agent.technicals import compute_technicals

    assert rsi_zone(75).startswith("overbought") and rsi_zone(25).startswith("oversold")
    assert rsi_zone(50).startswith("neutral")
    up = [100 + i * 0.5 + (3 if i % 15 < 7 else -3) for i in range(300)]
    dates = [f"2025-{1 + i // 28 % 12:02d}-{1 + i % 28:02d}" for i in range(300)]
    t = compute_technicals(dates, up, [100.0] * 300)
    lines = technical_commentary(t)
    assert lines[0].startswith("<b>Trend.</b> Uptrend")
    assert f"{t.last_close:,.0f}" in lines[0]
    assert any(s.startswith("<b>Momentum.</b>") and f"{t.rsi14:.0f}" in s for s in lines)
    rel = [s for s in lines if s.startswith("<b>Relative")][0]
    assert f"{abs(t.relative['12M']) * 100:.1f} pp" in rel
