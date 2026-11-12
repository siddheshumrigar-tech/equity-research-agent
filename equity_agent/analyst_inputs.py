"""Sourced analyst inputs written by the host agent: ``<SYMBOL>_analyst_inputs.json``.

Everything here changes numbers, so every entry must carry a ``source``:

    {"type": "annual_report" | "concall" | "investor_presentation" | "exchange_filing" | "news" | "website" | "other",
     "title": "Q2 FY27 earnings call", "url": "https://..."  (or "document": "AR FY26 p.42"),
     "date": "2026-10-15"}

Entries without a valid source are rejected (exit code 6). Applied overrides are
listed on the Excel ``Analyst Inputs`` sheet, highlighted in the driver table, and
labelled "analyst input" in the PDF.

Schema (all keys optional except ``symbol``)::

    {
      "symbol": "RELIANCE",
      "prepared_by": "Claude (host agent)",
      "as_of": "2026-11-12",
      "driver_overrides": [
        {"driver": "sales_growth", "years": {"FY27": 0.11, "FY28": 0.10}, "rationale": "...", "source": {...}},
        {"driver": "ebitda_margin", "all_years": 0.20, "rationale": "...", "source": {...}}
      ],
      "terminal": {"growth": 0.05, "ronic": 0.14, "rationale": "...", "source": {...}},
      "peers": {"symbols": ["INFY", "HCLTECH"], "reason": "...", "source": {...}},
      "sotp": {"segments": [
                 {"segment": "O2C", "method": "EV/EBITDA", "metric": "FY28E EBITDA", "metric_value_cr": 62000,
                  "multiple": 7.5, "source": {...}},
                 {"segment": "Jio (66% stake)", "method": "Stake value", "value_cr": 450000, "source": {...}}],
               "other_adjustment_cr": 0, "adjustment_note": "holding-company discount", "source": {...}},
      "segments": [{"name": "Home care", "revenue_share": 0.33, "ebit_share": 0.29, "period": "FY26",
                    "source": {...}}],
      "catalysts": [{"date": "2027-01-15", "event": "Q3 FY27 results", "impact": "positive|negative|neutral",
                     "source": {...}}]
    }
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from .drivers import YEAR_DRIVERS, Drivers, Override
from .errors import InputValidationError

SOURCE_TYPES = ("annual_report", "concall", "investor_presentation", "exchange_filing", "news", "website",
                "regulator", "industry_report", "other")
_DATE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
DRIVER_RANGES = {"sales_growth": (-0.5, 1.0), "cogs_pct": (0, 0.95), "ebitda_margin": (-0.5, 0.8),
                 "other_income_pct": (0, 0.3), "interest_rate": (0, 0.3), "dep_rate": (0, 0.6), "tax_rate": (0, 0.5),
                 "minority_pct": (-0.6, 0.6), "capex_pct": (0, 0.8), "dso": (0, 400), "dio": (0, 400),
                 "dpo": (0, 400), "dividend_payout": (0, 1.5), "other_ca_pct": (0, 3), "other_liab_pct": (0, 3)}


def source_errors(src, where: str) -> List[str]:
    if not isinstance(src, dict):
        return [f"{where}: missing 'source' object (type, title, url or document, date)."]
    errs = []
    if src.get("type") not in SOURCE_TYPES:
        errs.append(f"{where}: source.type must be one of {', '.join(SOURCE_TYPES)}.")
    if not str(src.get("title", "")).strip():
        errs.append(f"{where}: source.title is required.")
    if not (str(src.get("url", "")).strip() or str(src.get("document", "")).strip()):
        errs.append(f"{where}: source needs a url or a document reference.")
    if not _DATE.match(str(src.get("date", ""))):
        errs.append(f"{where}: source.date must be YYYY-MM-DD.")
    return errs


def source_text(src: dict) -> str:
    ref = src.get("url") or src.get("document") or ""
    return f"{src.get('title', '')} ({src.get('type', '')}, {src.get('date', '')}) {ref}".strip()


@dataclass
class AppliedOverride:
    driver: str
    year_label: str
    value: float
    historical: Optional[float]
    model_value: float
    rationale: str
    source: str


@dataclass
class AnalystInputs:
    symbol: str
    prepared_by: str = ""
    as_of: str = ""
    driver_overrides: List[dict] = field(default_factory=list)
    terminal_growth: Optional[float] = None
    terminal_ronic: Optional[float] = None
    terminal_note: str = ""
    peer_symbols: Optional[List[str]] = None
    peer_reason: str = ""
    sotp: List[dict] = field(default_factory=list)
    sotp_other_adjustment_cr: float = 0.0
    sotp_adjustment_note: str = ""
    segments: List[dict] = field(default_factory=list)
    catalysts: List[dict] = field(default_factory=list)
    applied: List[AppliedOverride] = field(default_factory=list)
    path: str = ""

    def apply(self, drivers: Drivers, forecast_labels: List[str]) -> List[AppliedOverride]:
        """Write per-year overrides into ``drivers``. Labels like 'FY27' or 'FY27E' map to forecast years."""
        idx = {lab.rstrip("E"): t for t, lab in enumerate(forecast_labels)}
        self.applied = []
        for ov in self.driver_overrides:
            key = ov["driver"]
            years = ov.get("years") or {}
            if "all_years" in ov:
                years = {lab.rstrip("E"): ov["all_years"] for lab in forecast_labels}
            for lab, val in years.items():
                t = idx.get(str(lab).upper().rstrip("E"))
                if t is None:
                    raise InputValidationError(f"analyst_inputs: driver '{key}' year '{lab}' is not a forecast year "
                                               f"({', '.join(forecast_labels)}).")
                src = source_text(ov["source"])
                model_value = drivers.value(key, "base", t)
                drivers.overrides.setdefault(key, {})[t] = Override(float(val), src)
                self.applied.append(AppliedOverride(key, forecast_labels[t], float(val), drivers.items[key].historical,
                                                    model_value, ov.get("rationale", ""), src))
        return self.applied

    def sotp_rows(self) -> Tuple[List[dict], float]:
        rows, total = [], 0.0
        for s in self.sotp:
            v = s.get("value_cr")
            if v is None:
                v = float(s["metric_value_cr"]) * float(s["multiple"])
            rows.append({**s, "ev_cr": float(v), "source_text": source_text(s["source"])})
            total += float(v)
        return rows, total

    def entries(self) -> List[dict]:
        """Flat list for the Excel 'Analyst Inputs' sheet and the PDF."""
        out = []
        for a in self.applied:
            out.append({"item": f"Driver: {a.driver} {a.year_label}", "value": a.value, "model": a.model_value,
                        "rationale": a.rationale, "source": a.source})
        if self.terminal_growth is not None:
            out.append({"item": "Terminal growth", "value": self.terminal_growth, "model": None,
                        "rationale": self.terminal_note, "source": self._terminal_src})
        if self.terminal_ronic is not None:
            out.append({"item": "Terminal RONIC", "value": self.terminal_ronic, "model": None,
                        "rationale": self.terminal_note, "source": self._terminal_src})
        if self.peer_symbols:
            out.append({"item": "Peer set", "value": ", ".join(self.peer_symbols), "model": None,
                        "rationale": self.peer_reason, "source": self._peer_src})
        for r in self.sotp_rows()[0]:
            out.append({"item": f"SOTP: {r['segment']}", "value": r["ev_cr"], "model": None,
                        "rationale": r.get("method", ""), "source": r["source_text"]})
        for s in self.segments:
            out.append({"item": f"Segment: {s.get('name')}", "value": s.get("revenue_share"), "model": None,
                        "rationale": f"{s.get('period', '')} revenue share", "source": source_text(s["source"])})
        for c in self.catalysts:
            out.append({"item": f"Catalyst {c.get('date')}", "value": c.get("event"), "model": None,
                        "rationale": c.get("impact", ""), "source": source_text(c["source"])})
        return out

    _terminal_src: str = ""
    _peer_src: str = ""


def validate_analyst_inputs(data: dict, symbol: Optional[str] = None) -> List[str]:
    errs: List[str] = []
    if not isinstance(data, dict):
        return ["analyst_inputs must be a JSON object."]
    if symbol and str(data.get("symbol", "")).upper() != symbol.upper():
        errs.append(f"analyst_inputs.symbol '{data.get('symbol')}' does not match '{symbol}'.")
    for i, ov in enumerate(data.get("driver_overrides", []) or []):
        where = f"driver_overrides[{i}]"
        key = ov.get("driver")
        if key not in YEAR_DRIVERS:
            errs.append(f"{where}: driver must be one of {', '.join(YEAR_DRIVERS)}.")
            continue
        vals = list((ov.get("years") or {}).values()) + ([ov["all_years"]] if "all_years" in ov else [])
        if not vals:
            errs.append(f"{where}: give 'years' {{'FY27': value}} or 'all_years'.")
        lo, hi = DRIVER_RANGES[key]
        for v in vals:
            if not isinstance(v, (int, float)) or not lo <= v <= hi:
                errs.append(f"{where}: value {v!r} outside the allowed range {lo}..{hi} for {key} "
                            "(percentages are decimals, e.g. 0.12).")
        errs += source_errors(ov.get("source"), where)
    term = data.get("terminal")
    if term:
        for k, (lo, hi) in (("growth", (0.0, 0.09)), ("ronic", (0.03, 1.5))):
            if k in term and not (isinstance(term[k], (int, float)) and lo <= term[k] <= hi):
                errs.append(f"terminal.{k}: {term.get(k)!r} outside {lo}..{hi}.")
        errs += source_errors(term.get("source"), "terminal")
    peers = data.get("peers")
    if peers:
        syms = peers.get("symbols")
        if not syms or not all(isinstance(s, str) and re.fullmatch(r"[A-Za-z0-9&\-_.]{1,20}", s) for s in syms):
            errs.append("peers.symbols must be a non-empty list of NSE symbols.")
        if not str(peers.get("reason", "")).strip():
            errs.append("peers.reason is required.")
        errs += source_errors(peers.get("source"), "peers")
    sotp = data.get("sotp")
    if sotp:
        segs = sotp.get("segments") or []
        if not segs:
            errs.append("sotp.segments must list at least one segment.")
        for i, s in enumerate(segs):
            where = f"sotp.segments[{i}]"
            if not str(s.get("segment", "")).strip():
                errs.append(f"{where}: segment name is required.")
            if s.get("value_cr") is None and (s.get("metric_value_cr") is None or s.get("multiple") is None):
                errs.append(f"{where}: give value_cr, or metric_value_cr and multiple.")
            errs += source_errors(s.get("source"), where)
        if sotp.get("other_adjustment_cr"):
            errs += source_errors(sotp.get("source"), "sotp (other_adjustment_cr)")
    for i, s in enumerate(data.get("segments", []) or []):
        where = f"segments[{i}]"
        if not str(s.get("name", "")).strip():
            errs.append(f"{where}: name is required.")
        for k in ("revenue_share", "ebit_share"):
            if k in s and s[k] is not None and not (isinstance(s[k], (int, float)) and -1 <= s[k] <= 1.5):
                errs.append(f"{where}: {k} must be a decimal share (e.g. 0.33).")
        errs += source_errors(s.get("source"), where)
    for i, c in enumerate(data.get("catalysts", []) or []):
        where = f"catalysts[{i}]"
        if not _DATE.match(str(c.get("date", ""))) and not re.match(r"^(Q[1-4] )?FY\d{2}$|^\d{4}-\d{2}$",
                                                                     str(c.get("date", ""))):
            errs.append(f"{where}: date must be YYYY-MM-DD, YYYY-MM, FYxx or 'Qn FYxx'.")
        if not str(c.get("event", "")).strip():
            errs.append(f"{where}: event is required.")
        errs += source_errors(c.get("source"), where)
    return errs


def parse_analyst_inputs(data: dict, symbol: Optional[str] = None, path: str = "") -> AnalystInputs:
    errs = validate_analyst_inputs(data, symbol)
    if errs:
        raise InputValidationError("analyst_inputs.json failed validation:\n  - " + "\n  - ".join(errs))
    ai = AnalystInputs(symbol=str(data.get("symbol", symbol or "")).upper(), prepared_by=data.get("prepared_by", ""),
                       as_of=data.get("as_of", ""), driver_overrides=list(data.get("driver_overrides", []) or []),
                       segments=list(data.get("segments", []) or []), catalysts=list(data.get("catalysts", []) or []),
                       path=path)
    term = data.get("terminal") or {}
    if term:
        ai.terminal_growth = term.get("growth")
        ai.terminal_ronic = term.get("ronic")
        ai.terminal_note = term.get("rationale", "")
        ai._terminal_src = source_text(term["source"])
    peers = data.get("peers") or {}
    if peers:
        ai.peer_symbols = [s.upper() for s in peers["symbols"]]
        ai.peer_reason = peers.get("reason", "")
        ai._peer_src = source_text(peers["source"])
    sotp = data.get("sotp") or {}
    if sotp:
        ai.sotp = list(sotp.get("segments", []))
        ai.sotp_other_adjustment_cr = float(sotp.get("other_adjustment_cr", 0) or 0)
        ai.sotp_adjustment_note = sotp.get("adjustment_note", "")
    return ai


def load_analyst_inputs(path, symbol: Optional[str] = None) -> Optional[AnalystInputs]:
    p = Path(path)
    if not p.exists():
        return None
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except ValueError as exc:
        raise InputValidationError(f"{p.name} is not valid JSON: {exc}") from exc
    return parse_analyst_inputs(data, symbol, str(p))


SCHEMA_EXAMPLE: Dict = {
    "symbol": "<SYMBOL>",
    "prepared_by": "<agent name>",
    "as_of": "YYYY-MM-DD",
    "driver_overrides": [{"driver": "sales_growth", "years": {"FY27": 0.11}, "rationale": "Management guidance",
                          "source": {"type": "concall", "title": "Q2 FY27 earnings call", "url": "https://...",
                                     "date": "YYYY-MM-DD"}}],
    "terminal": {"growth": 0.05, "ronic": 0.15, "rationale": "...",
                 "source": {"type": "other", "title": "...", "document": "...", "date": "YYYY-MM-DD"}},
    "peers": {"symbols": ["PEER1", "PEER2"], "reason": "closest business mix",
              "source": {"type": "annual_report", "title": "...", "url": "https://...", "date": "YYYY-MM-DD"}},
    "sotp": {"segments": [{"segment": "Segment A", "method": "EV/EBITDA", "metric": "FY28E EBITDA",
                           "metric_value_cr": 1000, "multiple": 10,
                           "source": {"type": "investor_presentation", "title": "...", "url": "https://...",
                                      "date": "YYYY-MM-DD"}}],
             "other_adjustment_cr": 0},
    "segments": [{"name": "Segment A", "revenue_share": 0.6, "ebit_share": 0.7, "period": "FY26",
                  "source": {"type": "annual_report", "title": "...", "document": "AR FY26 p.12", "date": "YYYY-MM-DD"}}],
    "catalysts": [{"date": "YYYY-MM-DD", "event": "Q3 results", "impact": "neutral",
                   "source": {"type": "exchange_filing", "title": "...", "url": "https://...", "date": "YYYY-MM-DD"}}],
}
