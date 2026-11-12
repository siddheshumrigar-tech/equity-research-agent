# Handoff: institutional-grade equity research agent (state at end of 12-Nov-2026, second session)

Read this first in a new session. The full design is in
[2026-11-11-institutional-upgrade-design.md](2026-11-11-institutional-upgrade-design.md) (its last section has the
before/after result) and the numbers are in [baseline.md](baseline.md).

**All 16 tasks are done.** 136 tests pass. Repository: `d:\Antigravity 2.0\Equity Research Agent\equity-research-agent`,
branch `feature/real-data-dashboard`. **Nothing is committed yet.**

## Decisions agreed with the user

1. Fix the analysis and redesign the PDF in one plan, with the analysis first.
2. The tool makes no LLM calls. The host agent writes `<SYM>_narrative.json` and optionally
   `<SYM>_analyst_inputs.json`. Every input must cite a source and is labelled "analyst input" in the outputs.
3. Layout: Indian broker front page (main column plus data sidebar), plus a risk-reward box and a bull/base/bear page.
4. Brand colour throughout, darkened for WCAG 4.5:1 contrast. White background.
5. Ratings BUY / ADD / REDUCE / SELL on 12-month expected total return (BUY > 15%, ADD 5-15%, REDUCE -5% to 5%,
   SELL < -5%).
6. Every new valuation item in Excel is a live formula, with Excel-vs-Python tests (Scenarios sheet exception).
7. Style reference: `..\HUL_Equity_Research_Report_Page_01..25.png`.
8. Defaults: terminal RONIC = WACC + 50% of (ROIC - WACC); target weights 50% DCF / 25% P/E / 25% EV/EBITDA.
9. Peer comparison is multi-year, in the PDF and (since this session) in the Excel Peers sheet.
10. `memory/learnings.json` was deleted; `EQUITY_AGENT_NO_HISTORY` stops the CLI recreating it.
11. (This session) The user approved the technical page and asked for all remaining tasks to be completed, taking
    the recommended option at each decision.

## Done this session

* **Technical view** (Task 14): `report/sections/technical.py`, `Charts.technical`. Page after valuation.
* **Excel Peers history**: `excel_writer._write_peer_history` writes 5 fiscal years of revenue growth, EBITDA and net
  margin, ROE, ROCE per company, a live `MEDIAN` peer row, a company-minus-median row and "years above median"
  (`COUNTIF`). Test: `test_peer_history_in_excel_matches_python`.
* **Narrative sections** (Task 15): `report/sections/research.py` with thesis (optional chart per point via the
  `exhibit` keyword), company overview (+ segment donuts from analyst inputs), industry (market share table and
  chart, Porter's table, SWOT grid), PESTLE and ESG, risks (impact on target, chart) and catalyst timeline, corporate
  actions. Each returns `[]` without content; `builder.maybe()` numbers only the sections present. Drafts show none
  (the draft test still asserts no "Porter"). Page 1 adds "Key risks to the call" (downside risks only) when the
  narrative has risks. New charts: `impact_bars`, `timeline`. Citation markers use a non-breaking space.
* **Task 16**: disclosures with rating definitions, preparation, interests and limits (`appendix.disclosures`);
  left-aligned data-source bullets; README rewritten for the two-step workflow, new sheets and method; AGENT.md notes
  `summary` / `exhibit`; before/after tables in the spec and `baseline.md`.
* **Host-agent files for all 5 tickers** in `output/agent/` (narratives with web sources dated Feb-Nov 2026, analyst
  inputs). FY27 revenue-growth overrides: ASIANPAINT 12.6%, TATATECH 19.8%, MARUTI 19.2% (first-half results plus
  second-half guidance or volume trend). Reliance has segments and catalysts but no sum-of-the-parts (reason in its
  flag responses). Both steps run cleanly for all 5 (draft 13 pages; final 18-19 pages).
* Test fixture `tests/fixtures/agent/TATATECH_narrative.json` now exercises every narrative section;
  `test_two_step_render` checks section order; `test_catalyst_dates`, `test_technical_commentary`,
  `test_technical_section_in_pdf` added.

## Final calls (live data, with analyst inputs)

| Stock | Target | Upside | Rating | TV share of EV |
|---|---|---|---|---|
| TCS | 3,107 | +25.7% | BUY | 54% |
| MARUTI | 17,151 | +30.1% | BUY | 54% |
| TATATECH | 563 | -16.7% | SELL | 48% |
| RELIANCE | 794 | -40.5% | SELL | 61% |
| ASIANPAINT | 1,553 | -39.1% | SELL | 45% |

## Possible next steps (not started; ask the user)

* A sourced sum-of-the-parts for Reliance (needs full-year segment EBITDA, multiples, and a bridge adjustment so the
  listed Jio stake does not double count Jio debt and minorities).
* Commit the work.
* `test_pdf_cli` still allows 5-14 pages for drafts (they are 13); final reports are not page-limited by a test.

## Practical notes

* **Shell:** `execute_pwsh` runs in a shared terminal; a new command can interrupt one still running, and output is
  often garbled. Start anything longer than a few seconds detached:
  `Start-Process -FilePath python -ArgumentList '"<script>"','ARG' -WorkingDirectory '<repo>' -WindowStyle Hidden`,
  then poll files. `control_pwsh_process` processes exit immediately here.
* **`read_file` may return a cached copy**; pass `offset: 0` with a different `limit` to force a fresh read.
* **Scripts in `output/checkpoint/` (temporary):** `_render.py SYMS...` (drafts + PNGs + tests), `_final.py SYM
  [fixture]` (both CLI steps into `output/checkpoint/final/`, PNGs in `final/pages/<SYM>/`), `_final_all.py` (all 5
  then the test suite into `final/_pytest.log`), `_t.py [pytest args]` (tests into `_t.log`), `_final_table.py`.
* **Excel parity:** recalculation uses the `formulas` package (no LibreOffice installed).
* **Plan conventions:** do not add unrequested features; keep tests green; show the user rendered pages for layout work.
