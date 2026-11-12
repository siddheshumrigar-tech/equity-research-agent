"""End-to-end run.

Run 1 (draft):  Screener -> financials -> drivers -> forecast -> valuation -> checks ->
                Excel + draft PDF + summary.json + research_pack.json
Run 2 (render): same data from cache + the host agent's analyst_inputs.json and
                narrative.json -> final Excel, PDF and summary.json
"""
from __future__ import annotations

import json
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import List, Optional

from .analyst_inputs import AnalystInputs, load_analyst_inputs
from .brand import Palette, get_brand
from .checks import Flag, high_flags, run_checks
from .config import with_defaults
from .drivers import Drivers, build_drivers
from .errors import NetworkError
from .financials import CompanyFinancials, build_financials
from .market_data import MarketData, fetch_market_data
from .model_engine import Forecast, run_forecast
from .peers import PeerSet, get_peers
from .screener_client import ScreenerClient
from .technicals import Technicals, compute_technicals
from .valuation import Valuation, compute_valuation, load_market_config  # noqa: F401


@dataclass
class RunResult:
    fin: CompanyFinancials
    drivers: Drivers
    forecast: Forecast
    valuation: Valuation
    palette: Palette
    market: MarketData
    cfg: dict
    peers: Optional[PeerSet] = None
    technicals: Optional[Technicals] = None
    flags: List[Flag] = field(default_factory=list)
    analyst: Optional[AnalystInputs] = None
    narrative: object = None
    xlsx_path: Optional[str] = None
    pdf_path: Optional[str] = None
    summary_path: Optional[str] = None
    pack_path: Optional[str] = None
    warnings: List[str] = field(default_factory=list)
    report_exhibits: List[str] = field(default_factory=list)

    @property
    def is_final(self) -> bool:
        return self.narrative is not None


def analyse(symbol: str, cache_dir: str | Path = "cache", offline: bool = False, refresh: bool = False,
            cfg: Optional[dict] = None, client: Optional[ScreenerClient] = None,
            analyst_inputs_path: Optional[str | Path] = None, frozen_cache: bool = False) -> RunResult:
    """``frozen_cache`` reuses cached responses regardless of age (used by --render so numbers do not move
    between the draft and the final report); missing items are still downloaded."""
    cfg = with_defaults(cfg)
    ttl = None if frozen_cache else 24 * 3600
    client = client or ScreenerClient(cache_dir=cache_dir, offline=offline, refresh=refresh, ttl_seconds=ttl)
    raw = client.fetch_company(symbol)
    fin = build_financials(raw)
    analyst = load_analyst_inputs(analyst_inputs_path, fin.symbol) if analyst_inputs_path else None
    drivers = build_drivers(fin, cfg)
    n_years = int(cfg.get("forecast_years", 5))
    labels = [f"FY{(fin.years[-1] + i) % 100:02d}E" for i in range(1, n_years + 1)]
    if analyst is not None:
        analyst.apply(drivers, labels)
    forecast = run_forecast(fin, drivers, "base", n_years, cfg.get("cwip_years"))
    md = fetch_market_data(fin.symbol, cfg, cache_dir=client.cache_dir, offline=offline, refresh=refresh,
                           ttl_seconds=ttl)
    peer_warn: List[str] = []
    try:
        peers = get_peers(client, raw, cfg, symbols=analyst.peer_symbols if analyst else None,
                          reason=analyst.peer_reason if analyst else "")
    except NetworkError as exc:
        peers = None
        peer_warn.append(f"Peer data not downloaded ({exc}); EV/EBITDA method left out.")
    if peers is None and not peer_warn:
        peer_warn.append("No peer snapshot available; EV/EBITDA method left out.")
    val = compute_valuation(fin, drivers, forecast, md, cfg, peers=peers, analyst=analyst)
    tech = compute_technicals(md.price_dates, md.price_close, md.index_close) if md.price_close else None
    flags = run_checks(fin, drivers, val, md, cfg, peers=peers)
    palette = get_brand(fin.symbol, fin.logo_url, fin.website, client)
    warnings = (list(fin.warnings) + list(drivers.notes) + list(md.warnings) + list(val.warnings) + peer_warn
                + (list(peers.warnings) if peers else []))
    return RunResult(fin=fin, drivers=drivers, forecast=forecast, valuation=val, palette=palette, market=md, cfg=cfg,
                     peers=peers, technicals=tech, flags=flags, analyst=analyst, warnings=warnings)


def write_outputs(res: RunResult, output_dir: str | Path = "output", make_pdf: bool = True) -> RunResult:
    from .excel_writer import write_model

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    sym = res.fin.symbol
    xlsx = out / f"{sym}_Financial_Model.xlsx"
    write_model(xlsx, res.fin, res.drivers, res.forecast, res.valuation, res.palette, res.cfg, res.market,
                peers=res.peers, analyst=res.analyst, flags=res.flags)
    res.xlsx_path = str(xlsx)
    if make_pdf:
        from .pdf_report import write_pdf

        pdf = out / f"{sym}_Equity_Research_Report.pdf"
        write_pdf(pdf, res)
        res.pdf_path = str(pdf)
    from .research_pack import write_research_pack

    res.pack_path = str(write_research_pack(res, out))
    res.summary_path = str(write_summary(res, out))
    return res


def _r(x, d=4):
    return round(x, d) if isinstance(x, (int, float)) else x


def summary_dict(res: RunResult) -> dict:
    f, v = res.fin, res.valuation
    t = res.technicals
    shp = {k: vals[-1] for k, vals in f.shareholding.items() if vals}
    return {
        "status": "ok",
        "report_stage": "final" if res.is_final else "draft",
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "symbol": f.symbol,
        "name": f.name,
        "sector": f.sector,
        "industry": f.industry,
        "data_source": f.source_url,
        "data_fetched_at": f.fetched_at,
        "consolidated": f.consolidated,
        "fiscal_years_reported": [f.years[0], f.years[-1]],
        "forecast_years": res.forecast.years,
        "price": f.price,
        "market_cap_cr": f.market_cap,
        "latest_year": {
            "fiscal_year": f.years[-1],
            "revenue_cr": f.latest("sales"),
            "ebitda_cr": f.latest("ebitda"),
            "net_profit_cr": f.latest("net_profit"),
        },
        "valuation": {
            "valuation_date": v.valuation_date,
            "wacc": round(v.wacc, 6),
            "cost_of_equity": round(v.ke, 6),
            "beta": round(v.beta, 4),
            "beta_raw": _r(v.beta_raw),
            "terminal_growth": v.terminal_growth,
            "roic": _r(v.roic),
            "terminal_ronic": _r(v.ronic),
            "tv_share_of_ev": _r(v.tv_share),
            "dcf_value": round(v.dcf_value, 2),
            "dcf_value_12m": round(v.dcf_value_12m, 2),
            "pe_value": _r(v.pe_value, 2),
            "mean_pe": _r(v.pe_mean, 2),
            "median_pe": _r(v.pe_mean, 2),  # kept for older callers (now the mean)
            "ev_ebitda_value": _r(v.ev_ebitda_value, 2),
            "peer_median_ev_ebitda": _r(v.peer_ev_ebitda, 2),
            "sotp_value": _r(v.sotp_value, 2),
            "target_method": v.target_method,
            "weights": {k: round(w, 4) for k, w in v.weights.items()},
            "target_price": round(v.target_price, 2),
            "upside": round(v.upside, 4),
            "dividend_yield": round(v.dividend_yield, 4),
            "expected_total_return": round(v.expected_total_return, 4),
            "rating": v.rating,
            "verdict": v.rating,
            "reverse_dcf_growth": _r(v.reverse_dcf_growth),
            "scenarios": {s: {"target": round(sv.target, 2), "expected_total_return": round(sv.expected_total_return, 4),
                              "probability": sv.probability} for s, sv in v.scenarios.items()},
            "probability_weighted_value": _r(v.prob_weighted_value, 2),
        },
        "flags": [fl.to_dict() for fl in res.flags],
        "high_flags_unanswered": [fl.code for fl in high_flags(res.flags)
                                  if not (res.narrative and fl.code in getattr(res.narrative, "flag_responses", {}))],
        "shareholding_latest": {"quarter": f.shareholding_labels[-1] if f.shareholding_labels else None, **shp},
        "price_performance": ({"absolute": {k: _r(x) for k, x in t.returns.items()},
                               "relative_to_nifty": {k: _r(x) for k, x in t.relative.items()}} if t else None),
        "technicals": ({"sma50": _r(t.sma50, 2), "sma200": _r(t.sma200, 2), "rsi14": _r(t.rsi14, 1),
                        "supports": [round(x, 2) for x in t.supports], "resistances": [round(x, 2) for x in t.resistances],
                        "trend": t.trend} if t else None),
        "peers": ([p.symbol for p in res.peers.peers] if res.peers else []),
        "analyst_inputs": res.analyst.path if res.analyst else None,
        "brand_colour": "#" + res.palette.primary,
        "files": {"excel": res.xlsx_path, "pdf": res.pdf_path, "research_pack": res.pack_path},
        "warnings": res.warnings,
    }


def write_summary(res: RunResult, out: Path) -> Path:
    path = out / f"{res.fin.symbol}_summary.json"
    path.write_text(json.dumps(summary_dict(res), indent=2, ensure_ascii=False), encoding="utf-8")
    return path
