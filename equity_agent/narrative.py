"""Narrative written by the host agent: ``<SYMBOL>_narrative.json``.

The tool never writes opinions. The host agent (Claude, Hermes, ...) reads
``<SYMBOL>_research_pack.json``, researches the company (annual report, earnings calls,
exchange filings, industry sources) and writes the narrative. Rules enforced here:

* every factual claim cites a source with a marker like ``[2]`` that refers to
  ``sources[].id``; unknown markers are errors
* model numbers are written as tokens, e.g. ``{{target_price}}``, and filled in by the
  renderer, so text and model cannot disagree; unknown tokens are errors
* any other number in a sentence without a citation marker produces a lint warning
* ``flag_responses`` must answer every high-severity model flag
* ``headline`` and at least three ``thesis_points`` are required; all other sections are
  optional and are left out of the PDF when missing
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional

from .analyst_inputs import SOURCE_TYPES
from .errors import InputValidationError
from .ratios import latest_actual

TOKEN_RE = re.compile(r"\{\{\s*([a-z0-9_.]+)\s*\}\}")
CITE_RE = re.compile(r"\[(\d+(?:\s*,\s*\d+)*)\]")
NUM_RE = re.compile(r"(?<![A-Za-z])(\d[\d,]*(?:\.\d+)?)\s*(%|x|bn|cr|crore|mn|lakh|bps)?", re.I)
PORTER_FORCES = ("new_entrants", "supplier_power", "buyer_power", "substitutes", "rivalry")
PORTER_LABELS = {"new_entrants": "Threat of new entrants", "supplier_power": "Bargaining power of suppliers",
                 "buyer_power": "Bargaining power of buyers", "substitutes": "Threat of substitutes",
                 "rivalry": "Competitive rivalry"}
RATINGS = ("low", "moderate", "high")
PESTLE = ("political", "economic", "social", "technological", "legal", "environmental")
SWOT = ("strengths", "weaknesses", "opportunities", "threats")
TEXT_SECTIONS = ("company_overview", "business_model", "management", "industry", "esg")


@dataclass
class Narrative:
    data: dict
    tokens: Dict[str, str]
    lint: List[str] = field(default_factory=list)
    path: str = ""

    @property
    def flag_responses(self) -> Dict[str, str]:
        return self.data.get("flag_responses", {}) or {}

    def has(self, key: str) -> bool:
        v = self.data.get(key)
        return bool(v) and (not isinstance(v, str) or v.strip() != "")

    def text(self, value) -> str:
        """Fill tokens in a string."""
        if not isinstance(value, str):
            return value
        return TOKEN_RE.sub(lambda m: self.tokens.get(m.group(1), m.group(0)), value)

    def get(self, key, default=None):
        return self.data.get(key, default)

    @property
    def sources(self) -> List[dict]:
        return list(self.data.get("sources", []) or [])


# --------------------------------------------------------------------------- #
# tokens
# --------------------------------------------------------------------------- #
def _rs(x) -> str:
    return "n/a" if x is None else f"₹{x:,.0f}" if abs(x) >= 100 else f"₹{x:,.2f}"


def _pct(x, d=1) -> str:
    return "n/a" if x is None else f"{x * 100:.{d}f}%"


def _signed(x) -> str:
    return "n/a" if x is None else f"{x * 100:+.1f}%"


def _x(x) -> str:
    return "n/a" if x is None else f"{x:.1f}x"


def _cr(x) -> str:
    return "n/a" if x is None else f"₹{x:,.0f} Cr"


def build_tokens(res) -> Dict[str, str]:
    f, v, fc, d = res.fin, res.valuation, res.forecast, res.drivers
    L = f.lines
    y = f"FY{f.years[-1] % 100:02d}"
    t = res.technicals
    tok = {
        "company": f.name, "symbol": f.symbol, "price": _rs(f.price), "market_cap": _cr(f.market_cap),
        "target_price": _rs(v.target_price), "upside": _signed(v.upside), "rating": v.rating,
        "expected_total_return": _signed(v.expected_total_return), "dividend_yield": _pct(v.dividend_yield),
        "dcf_value": _rs(v.dcf_value), "dcf_value_12m": _rs(v.dcf_value_12m), "pe_value": _rs(v.pe_value),
        "ev_ebitda_value": _rs(v.ev_ebitda_value), "sotp_value": _rs(v.sotp_value), "wacc": _pct(v.wacc, 2),
        "cost_of_equity": _pct(v.ke, 2), "beta": f"{v.beta:.2f}", "terminal_growth": _pct(v.terminal_growth),
        "roic": _pct(v.roic), "terminal_ronic": _pct(v.ronic), "tv_share": _pct(v.tv_share, 0),
        "mean_pe": _x(v.pe_mean), "peer_ev_ebitda": _x(v.peer_ev_ebitda), "current_ev_ebitda": _x(v.current_ev_ebitda),
        "trailing_pe": _x(f.pe), "reverse_dcf_growth": _pct(v.reverse_dcf_growth),
        "prob_weighted_value": _rs(v.prob_weighted_value),
        "latest_fy": y, "revenue_latest": _cr(L["sales"][-1]), "ebitda_latest": _cr(L["ebitda"][-1]),
        "net_profit_latest": _cr(L["net_profit"][-1]),
        "ebitda_margin_latest": _pct(L["ebitda"][-1] / L["sales"][-1] if L["sales"][-1] else None),
        "net_margin_latest": _pct(L["net_profit"][-1] / L["sales"][-1] if L["sales"][-1] else None),
        # latest reported year on the model's definitions (average balances), as in the tables and the Excel model
        "roe": _pct(latest_actual(res, "roe")), "roce": _pct(latest_actual(res, "roce")),
        "growth_forecast": _pct(d.value("sales_growth")), "margin_forecast": _pct(d.value("ebitda_margin")),
        "fy1": fc.year_labels[0], "fy2": fc.year_labels[1], "fy_last": fc.year_labels[-1],
        "revenue_fy1": _cr(fc.lines["sales"][0]), "revenue_fy2": _cr(fc.lines["sales"][1]),
        "eps_ntm": _rs(v.ntm_eps) if v.ntm_eps else "n/a",
    }
    n = f.n
    for k, yrs in (("revenue_cagr_3y", 3), ("revenue_cagr_5y", 5)):
        if n > yrs and L["sales"][-1 - yrs] > 0 and L["sales"][-1] > 0:
            tok[k] = _pct((L["sales"][-1] / L["sales"][-1 - yrs]) ** (1 / yrs) - 1)
    for s, sv in v.scenarios.items():
        tok[f"{s}_value"] = _rs(sv.target)
        tok[f"{s}_return"] = _signed(sv.expected_total_return)
    if t is not None:
        tok.update({"return_1m": _signed(t.returns.get("1M")), "return_3m": _signed(t.returns.get("3M")),
                    "return_12m": _signed(t.returns.get("12M")), "relative_12m": _signed(t.relative.get("12M")),
                    "rsi": f"{t.rsi14:.0f}" if t.rsi14 is not None else "n/a", "sma200": _rs(t.sma200),
                    "high_52w": _rs(t.high_52w), "low_52w": _rs(t.low_52w)})
    if f.shareholding:
        for k, lab in (("promoter_holding", "Promoters"), ("fii_holding", "FIIs"), ("dii_holding", "DIIs")):
            vals = f.shareholding.get(lab)
            if vals and vals[-1] is not None:
                tok[k] = f"{vals[-1]:.1f}%"
    return tok


# --------------------------------------------------------------------------- #
# validation
# --------------------------------------------------------------------------- #
def _strings(data, path="") -> List[tuple]:
    """(path, text) for every string in the narrative except sources and ids."""
    out = []
    if isinstance(data, dict):
        for k, v in data.items():
            if k in ("sources", "source_ids", "rating", "symbol", "as_of", "prepared_by", "date", "impact_level"):
                continue
            out += _strings(v, f"{path}.{k}" if path else k)
    elif isinstance(data, list):
        for i, v in enumerate(data):
            out += _strings(v, f"{path}[{i}]")
    elif isinstance(data, str):
        out.append((path, data))
    return out


def lint_numbers(path: str, text: str) -> List[str]:
    warns = []
    clean = TOKEN_RE.sub("", text)
    for sent in re.split(r"(?<=[.!?;])\s+", clean):
        if CITE_RE.search(sent):
            continue
        s2 = re.sub(r"\bFY\d{2,4}[AE]?\b|\bQ[1-4]\b|\bH[12]\b|\b(19|20)\d{2}\b", "", sent)
        for m in NUM_RE.finditer(s2):
            raw = m.group(1).replace(",", "")
            try:
                num = float(raw)
            except ValueError:
                continue
            if num <= 10 and not m.group(2):
                continue  # small counts ("three", "5 brands") are allowed
            warns.append(f"{path}: number '{m.group(0).strip()}' has no citation marker or token: \"{sent.strip()[:90]}\"")
            break
    return warns


def validate_narrative(data: dict, tokens: Dict[str, str], high_flag_codes: List[str],
                       symbol: Optional[str] = None) -> (List[str], List[str]):
    errs: List[str] = []
    if not isinstance(data, dict):
        return ["narrative must be a JSON object."], []
    if symbol and str(data.get("symbol", "")).upper() != symbol.upper():
        errs.append(f"narrative.symbol '{data.get('symbol')}' does not match '{symbol}'.")
    if not str(data.get("headline", "")).strip():
        errs.append("headline is required (one sentence stating the call and why).")
    elif len(data["headline"]) > 140:
        errs.append("headline must be at most 140 characters.")
    tp = data.get("thesis_points") or []
    if len(tp) < 3:
        errs.append("thesis_points needs at least three entries.")
    for i, p in enumerate(tp):
        if not isinstance(p, dict) or not str(p.get("title", "")).strip() or not str(p.get("text", "")).strip():
            errs.append(f"thesis_points[{i}] needs 'title' and 'text'.")
    srcs = data.get("sources") or []
    ids = set()
    for i, s in enumerate(srcs):
        if not isinstance(s, dict) or not isinstance(s.get("id"), int):
            errs.append(f"sources[{i}] needs an integer 'id'.")
            continue
        ids.add(s["id"])
        if s.get("type") not in SOURCE_TYPES:
            errs.append(f"sources[{i}].type must be one of {', '.join(SOURCE_TYPES)}.")
        if not str(s.get("title", "")).strip():
            errs.append(f"sources[{i}].title is required.")
        if not (str(s.get("url", "")).strip() or str(s.get("document", "")).strip()):
            errs.append(f"sources[{i}] needs a url or document reference.")
        if not re.match(r"^\d{4}-\d{2}-\d{2}$", str(s.get("date", ""))):
            errs.append(f"sources[{i}].date must be YYYY-MM-DD.")
    if not srcs:
        errs.append("sources must list at least one source.")
    porter = data.get("porter") or {}
    for k, v in porter.items():
        if k not in PORTER_FORCES:
            errs.append(f"porter.{k}: unknown force (use {', '.join(PORTER_FORCES)}).")
            continue
        if not isinstance(v, dict) or v.get("rating") not in RATINGS or not str(v.get("evidence", "")).strip():
            errs.append(f"porter.{k} needs rating (low / moderate / high) and evidence.")
    for k in (data.get("pestle") or {}):
        if k not in PESTLE:
            errs.append(f"pestle.{k}: unknown factor (use {', '.join(PESTLE)}).")
    for k in (data.get("swot") or {}):
        if k not in SWOT:
            errs.append(f"swot.{k}: unknown quadrant (use {', '.join(SWOT)}).")
    for i, r in enumerate(data.get("risks") or []):
        if not isinstance(r, dict) or not str(r.get("risk", "")).strip() or not str(r.get("impact", "")).strip():
            errs.append(f"risks[{i}] needs 'risk' and 'impact' (and ideally 'mitigant', 'impact_on_tp_pct').")
    fr = data.get("flag_responses") or {}
    for code in high_flag_codes:
        if not str(fr.get(code, "")).strip():
            errs.append(f"flag_responses.{code} is required: the model raised this high-severity flag.")
    for path, text in _strings(data):
        for m in TOKEN_RE.finditer(text):
            if m.group(1) not in tokens:
                errs.append(f"{path}: unknown token {{{{{m.group(1)}}}}}.")
        for m in CITE_RE.finditer(text):
            for n in re.split(r"\s*,\s*", m.group(1)):
                if int(n) not in ids:
                    errs.append(f"{path}: citation [{n}] has no matching sources[].id.")
    lint: List[str] = []
    for path, text in _strings(data):
        lint += lint_numbers(path, text)
    return errs, lint


def load_narrative(path, res, high_flag_codes: List[str]) -> Narrative:
    p = Path(path)
    if not p.exists():
        raise InputValidationError(f"{p.name} not found. Write it from the research pack, then run --render again.")
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except ValueError as exc:
        raise InputValidationError(f"{p.name} is not valid JSON: {exc}") from exc
    tokens = build_tokens(res)
    errs, lint = validate_narrative(data, tokens, high_flag_codes, res.fin.symbol)
    if errs:
        raise InputValidationError(f"{p.name} failed validation:\n  - " + "\n  - ".join(errs))
    return Narrative(data=data, tokens=tokens, lint=lint, path=str(p))


SCHEMA_EXAMPLE = {
    "symbol": "<SYMBOL>",
    "prepared_by": "<agent name>",
    "as_of": "YYYY-MM-DD",
    "headline": "One sentence with the call and the reason, e.g. 'Share gains priced in; initiate at {{rating}}'",
    "thesis_points": [
        {"title": "Short claim", "text": "Two to four sentences of evidence specific to this company [1]. Model "
                                         "numbers as tokens: revenue {{revenue_latest}} in {{latest_fy}}.",
         "exhibit": "optional: rev_ebitda | margins | fcf | roe | wc_days | peers_pe | price"}],
    "company_overview": "What the company does, segments, history [n].",
    "business_model": "How it makes money, pricing power, cost structure [n].",
    "management": "Leadership, track record, capital allocation, governance [n].",
    "industry": "Market size, growth, structure, regulation [n].",
    "market_share": [{"company": "Peer or self", "share": 0.12, "metric": "FY26 revenue share of organised market",
                      "source_ids": [1]}],
    "porter": {k: {"rating": "low | moderate | high", "evidence": "company-specific evidence [n]"} for k in PORTER_FORCES},
    "swot": {k: ["point with evidence [n]"] for k in SWOT},
    "pestle": {k: "factor and how it affects this company [n]" for k in PESTLE},
    "esg": "Material ESG issues for this business with data [n].",
    "risks": [{"risk": "Name", "impact": "What happens [n]", "impact_on_tp_pct": -0.10, "mitigant": "..."}],
    "corporate_actions": [{"date": "YYYY-MM-DD", "event": "Acquisition / merger / buyback [n]"}],
    "flag_responses": {"<FLAG_CODE>": "Why the call stands despite the flag, or what was changed [n]"},
    "sources": [{"id": 1, "type": "annual_report", "title": "Annual Report FY26", "url": "https://...",
                 "date": "YYYY-MM-DD"}],
}
