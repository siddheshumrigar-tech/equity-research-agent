# Equity Research Agent - instructions for AI agents (Hermes, Claude, Codex, Cursor)

Use this repository when a user asks for an equity research report, financial model,
valuation or "dashboard like the Tesla model" for an **Indian listed company (NSE/BSE)**.
The tool builds every number from reported data. **You** (the host agent) write the
qualitative research and may adjust drivers, but only with sources.

## 1. Identify the NSE symbol

The tool needs the exact symbol used in the Screener.in URL
(`https://www.screener.in/company/<SYMBOL>/`). Examples: Tata Technologies `TATATECH`,
Asian Paints `ASIANPAINT`, Reliance Industries `RELIANCE`, Infosys `INFY`. If you are not
sure, look it up on nseindia.com or ask the user. US and other non-Indian stocks are not
supported; banks and NBFCs are rejected (exit code 3).

## 2. Step 1 - draft run

```bash
pip install -r requirements.txt          # first time only
python generate_equity_report.py --ticker <SYMBOL>
```

Writes to `output/`: the Excel model, a **draft** PDF (watermarked), `<SYMBOL>_summary.json`
and `<SYMBOL>_research_pack.json`.

## 3. Step 2 - research and write two files

Read `<SYMBOL>_research_pack.json`. It contains the model numbers, the **flags you must
answer** (`flags_to_answer`), peers, quarterly results, shareholding, technicals, the
`tokens` you may use in text, and the schemas (`narrative_schema`, `analyst_inputs_schema`).

Research the company from primary sources first (annual report, earnings-call transcripts,
investor presentations, exchange filings), then industry and news sources. Then write:

* `output/<SYMBOL>_narrative.json` (required): headline, at least three thesis points,
  company overview, business model, management, industry, market share, Porter's five
  forces (rating + evidence each), SWOT, PESTLE, ESG, risks (impact + mitigant),
  corporate actions, `flag_responses` for every high flag, and `sources`. Each thesis point
  may carry a one-sentence `summary` (used on page 1; otherwise `text` is used) and an
  `exhibit` keyword (`rev_ebitda`, `margins`, `fcf`, `roe`, `wc_days`, `peers_pe`, `price`)
  that puts a chart next to it. Sections you leave out are left out of the PDF.
* `output/<SYMBOL>_analyst_inputs.json` (optional): per-year driver overrides (e.g.
  management growth guidance), terminal growth / RONIC, a peer list, sum-of-the-parts
  segment values, segment revenue shares and catalysts.

Writing rules (enforced or linted by `--render`):

1. Be specific to this company. Generic sector statements are not acceptable on their own.
2. Cite every fact with `[n]` matching `sources[].id` (each source: type, title, url or
   document, date). Unknown citations are errors.
3. Never type model numbers. Use tokens like `{{target_price}}`, `{{wacc}}`,
   `{{revenue_latest}}`. Unknown tokens are errors; uncited numbers produce lint warnings.
4. Change numbers only in `analyst_inputs.json`, each entry with a `source`. Entries
   without a valid source are rejected. Percentages are decimals (0.12, not 12).
5. Answer every high-severity flag in `flag_responses`: explain why the call stands, or
   fix the inputs and render again.

## 4. Step 3 - final render

```bash
python generate_equity_report.py --ticker <SYMBOL> --render
```

Re-uses the cached data (numbers do not move between draft and final), applies the analyst
inputs, validates the narrative and writes the final Excel, PDF and summary.

## 5. Read the result

Read `output/<SYMBOL>_summary.json`:

* `status`, `report_stage` (`draft` / `final`)
* `valuation.target_price`, `valuation.upside`, `valuation.expected_total_return`,
  `valuation.rating` (BUY / ADD / REDUCE / SELL; `verdict` is the same value),
  `valuation.target_method`, `valuation.weights`, `valuation.scenarios`
* `flags` and `high_flags_unanswered`
* `files.excel`, `files.pdf`, `files.research_pack`
* `warnings`: data notes and narrative lint you must pass on to the user

| Exit code | Meaning | What to do |
|---|---|---|
| 0 | Success | Share files, rating, target, return, flags and warnings |
| 2 | Symbol not found | Ask for the correct NSE symbol |
| 3 | Bank / NBFC / financial company | Not supported yet |
| 4 | Screener data failed validation | Report the message; do not build numbers yourself |
| 5 | Network failure | Retry later |
| 6 | narrative.json / analyst_inputs.json invalid or missing | Fix the listed errors and run `--render` again |
| 1 | Unexpected error | Report the traceback |

## 6. Method in one paragraph

Drivers are 3-year historical averages (working-capital days glide from the last year);
bull / bear move growth, margin and capex. Valuation: 3-stage FCFF DCF (5 explicit years,
5 fade years toward a terminal return on new capital, value-driver terminal value) rolled
forward 12 months, P/E vs the company's own 5-year history and EV/EBITDA vs the Screener
peer median, both on next-twelve-months earnings; weights 50 / 25 / 25 (config). Rating on
expected total return: BUY >= 15%, ADD >= 5%, REDUCE >= -5%, else SELL. Credibility checks
flag a terminal value above 75% of EV, DCF vs multiples more than 2x apart, |upside| above
40%, negative FCFF, an unreachable reverse DCF, clamped drivers, stale market inputs and
quarterly growth far from the forecast. Market inputs live in `config/market.json`; update
them when they go stale. The report is not investment advice.
