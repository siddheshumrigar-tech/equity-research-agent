"""Market and model configuration (``config/market.json``) with defaults for keys added later."""
from __future__ import annotations

import copy
import json
from pathlib import Path
from typing import Optional

CONFIG_PATH = Path(__file__).resolve().parents[1] / "config" / "market.json"

DEFAULTS = {
    "beta_min": 0.5,
    "beta_max": 2.0,
    "beta_default": 1.0,
    "beta_blume": True,
    "forecast_years": 5,
    "fade_years": 5,
    "roic_cap": 1.0,
    "ronic_premium_share": {"base": 0.5, "bull": 0.75, "bear": 0.25},
    "cwip_years": 2,
    "wc_glide_years": 2,
    "target_weights": {"dcf": 0.5, "pe": 0.25, "ev_ebitda": 0.25},
    "rating_thresholds": [[0.15, "BUY"], [0.05, "ADD"], [-0.05, "REDUCE"]],
    "rating_floor": "SELL",
    "scenarios": {"growth_offset": 0.03, "margin_offset": 0.015, "capex_mult": {"bull": 0.9, "bear": 1.1},
                  "probabilities": {"bull": 0.25, "base": 0.5, "bear": 0.25}},
    "checks": {"tv_share_max": 0.75, "dcf_relative_ratio_max": 2.0, "upside_abs_max": 0.40, "stale_days": 90,
               "ytd_gap_pp": 0.05},
    "max_peers": 6,
}


def _merge(base: dict, over: dict) -> dict:
    out = copy.deepcopy(base)
    for k, v in over.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _merge(out[k], v)
        else:
            out[k] = v
    return out


def load_market_config(path: Optional[str | Path] = None) -> dict:
    raw = json.loads(Path(path or CONFIG_PATH).read_text(encoding="utf-8"))
    return _merge(DEFAULTS, raw)


def with_defaults(cfg: Optional[dict]) -> dict:
    """Fill keys missing from a caller-supplied config dict."""
    if cfg is None:
        return load_market_config()
    return _merge(DEFAULTS, cfg)
