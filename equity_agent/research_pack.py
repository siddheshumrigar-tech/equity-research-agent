"""``<SYMBOL>_research_pack.json``: everything the host agent needs to write the narrative.

Contains the model numbers, the credibility flags it must answer, peers, quarterly
results, shareholding, technicals, the tokens it can use in text, and the schemas and
writing rules for ``narrative.json`` and ``analyst_inputs.json``.
"""
from __future__ import annotations

import json
from pathlib import Path

from .analyst_inputs import SCHEMA_EXAMPLE as ANALYST_SCHEMA
from .checks import high_flags
from .narrative import SCHEMA_EXAMPLE as NARRATIVE_SCHEMA
from .narrative import build_tokens

WRITING_RULES = [
    "Write about THIS company. Every claim needs company-specific evidence; generic statements that would fit any "
    "company in the sector (e.g. 'faces intense competition', 'benefits from India's growth') are not acceptable on "
    "their own.",
    "Cite every fact with a marker like [3] that refers to sources[].id. Use primary sources first: annual report, "
    "earnings-call transcript, investor presentation, exchange filings; then reputable industry and news sources. "
    "Give each source a url or document reference and a date.",
    "Never type a model number yourself. Use tokens such as {{target_price}}, {{wacc}}, {{revenue_latest}}; the list "
    "is in 'tokens'. Numbers from your sources (market share, volumes, guidance) must carry a citation marker.",
    "Answer every entry in 'flags_to_answer' in flag_responses: explain why the call stands, or change the inputs in "
    "analyst_inputs.json (with sources) and run --render again.",
    "Change numbers only through analyst_inputs.json, and only with a source (management guidance, disclosed segment "
    "data, regulatory changes). Each override is shown to the reader as an analyst input next to the historical value.",
    "Porter's five forces: rate each force low / moderate / high and give the evidence. SWOT: 2-4 points per quadrant. "
    "Risks: state the impact on earnings or the target (impact_on_tp_pct) and any mitigant.",
    "Keep sentences short and specific. No hype words; state the mechanism and the number.",
]


def _round(x, d=4):
    if isinstance(x, float):
        return round(x, d)
    if isinstance(x, list):
        return [_round(v, d) for v in x]
    if isinstance(x, dict):
        return {k: _round(v, d) for k, v in x.items()}
    return x


def build_research_pack(res) -> dict:
    from .pipeline import summary_dict

    f, v, fc = res.fin, res.valuation, res.forecast
    L = f.lines
    hist_keys = ("sales", "ebitda", "net_profit", "cfo", "capex", "fcf", "borrowings", "cash", "total_equity",
                 "receivables", "inventory", "trade_payables")
    fc_keys = ("sales", "ebitda", "net_profit", "fcff", "capex", "nwc")
    summ = summary_dict(res)
    pack = {
        "how_to_use": [
            f"1. Read this file. 2. Research {f.name} and write {f.symbol}_narrative.json (schema below) in the output "
            f"folder. 3. Optionally write {f.symbol}_analyst_inputs.json with sourced overrides. 4. Run: python "
            f"generate_equity_report.py --ticker {f.symbol} --render",
        ],
        "writing_rules": WRITING_RULES,
        "company": {"symbol": f.symbol, "name": f.name, "sector": f.sector, "industry": f.industry,
                    "website": f.website, "bse_code": f.bse_code, "consolidated": f.consolidated,
                    "screener_url": f.source_url},
        "summary": summ,
        "flags_to_answer": [fl.to_dict() for fl in high_flags(res.flags)],
        "all_flags": [fl.to_dict() for fl in res.flags],
        "history": {"years": f.year_labels, **{k: L[k] for k in hist_keys}},
        "expense_items": [{"name": it.name, "kind": it.kind, "latest_cr": it.values[-1]} for it in f.expense_items],
        "forecast": {"years": fc.year_labels, **{k: fc.lines[k] for k in fc_keys}},
        "drivers": [{"key": d.key, "label": d.label, "historical": d.historical, "base": d.base, "bull": d.bull,
                     "bear": d.bear, "basis": d.basis, "clamped": d.clamped} for d in res.drivers],
        "valuation_detail": {
            "wacc": v.wacc, "ke": v.ke, "beta": v.beta, "beta_raw": v.beta_raw, "g": v.terminal_growth,
            "roic": v.roic, "ronic": v.ronic, "tv_share": v.tv_share, "dcf_value": v.dcf_value,
            "dcf_value_12m": v.dcf_value_12m, "pe_history": v.pe_history, "pe_mean": v.pe_mean, "pe_sd": v.pe_sd,
            "pe_value": v.pe_value, "peer_ev_ebitda": v.peer_ev_ebitda, "ev_ebitda_value": v.ev_ebitda_value,
            "weights": v.weights, "target": v.target_price, "rating": v.rating,
            "reverse_dcf_growth": v.reverse_dcf_growth,
            "scenarios": {s: sv.__dict__ for s, sv in v.scenarios.items()},
        },
        "quarters": {"labels": f.quarter_labels, **f.quarters},
        "shareholding": {"labels": f.shareholding_labels, **f.shareholding},
        "peers": res.peers.to_dict() if res.peers else None,
        "technicals": res.technicals.to_dict() if res.technicals else None,
        "tokens": build_tokens(res),
        "narrative_schema": NARRATIVE_SCHEMA,
        "analyst_inputs_schema": ANALYST_SCHEMA,
        "analyst_inputs_applied": res.analyst.entries() if res.analyst else [],
    }
    return _round(pack)


def write_research_pack(res, out: Path) -> Path:
    path = Path(out) / f"{res.fin.symbol}_research_pack.json"
    path.write_text(json.dumps(build_research_pack(res), indent=1, ensure_ascii=False, default=str), encoding="utf-8")
    return path
