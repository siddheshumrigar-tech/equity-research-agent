"""Credibility checks on the model output.

Each check returns a ``Flag``. High-severity flags are printed on the PDF valuation
page and in the research pack, and ``narrative.json`` must answer every one of them
in ``flag_responses`` before ``--render`` produces a final report.
"""
from __future__ import annotations

from dataclasses import asdict, dataclass
from datetime import date
from typing import List, Optional

from .config import with_defaults

HIGH, MEDIUM, LOW = "high", "medium", "low"


@dataclass
class Flag:
    code: str
    severity: str
    message: str
    metric: Optional[float] = None

    def to_dict(self) -> dict:
        return asdict(self)


def relative_value(val) -> Optional[float]:
    vals = [v for v in (val.pe_value, val.ev_ebitda_value) if v is not None]
    return sum(vals) / len(vals) if vals else None


def run_checks(fin, drivers, val, md, cfg: Optional[dict] = None, valuation_date: Optional[date] = None,
               peers=None) -> List[Flag]:
    cfg = with_defaults(cfg)
    c = cfg["checks"]
    flags: List[Flag] = []
    vdate = valuation_date or date.fromisoformat(val.valuation_date)

    if val.tv_share > c["tv_share_max"]:
        flags.append(Flag("TV_SHARE_HIGH", HIGH,
                          f"Terminal value is {val.tv_share:.0%} of enterprise value (limit {c['tv_share_max']:.0%}); "
                          "the DCF rests on long-run assumptions (g, RONIC).", val.tv_share))
    rel = relative_value(val)
    dcf = val.dcf_value_12m
    if rel is not None:
        if dcf <= 0 or rel <= 0:
            flags.append(Flag("DCF_RELATIVE_DIVERGENCE", HIGH,
                              f"DCF value ({dcf:,.0f}) and relative value ({rel:,.0f}) have different signs.", None))
        else:
            ratio = max(dcf, rel) / min(dcf, rel)
            if ratio > c["dcf_relative_ratio_max"]:
                flags.append(Flag("DCF_RELATIVE_DIVERGENCE", HIGH,
                                  f"DCF value Rs {dcf:,.0f} and relative value Rs {rel:,.0f} differ {ratio:.1f}x "
                                  f"(limit {c['dcf_relative_ratio_max']:.1f}x); the methods disagree on what the "
                                  "business is worth.", ratio))
    if abs(val.upside) > c["upside_abs_max"]:
        flags.append(Flag("EXTREME_UPSIDE", HIGH,
                          f"Target implies {val.upside:+.0%} from the current price (limit +/-{c['upside_abs_max']:.0%}); "
                          "check what the market is pricing in before publishing.", val.upside))
    if val.fcff and (val.fcff[-1] <= 0 or val.fcff[len(val.fcff) // 2 - 1] <= 0):
        flags.append(Flag("NEGATIVE_FCFF", HIGH, "Final explicit or fade-year free cash flow is negative; the DCF is "
                          "not meaningful without a view on when investment pays back.", val.fcff[-1]))
    if val.reverse_dcf_growth is None:
        flags.append(Flag("REVERSE_DCF_UNREACHABLE", MEDIUM,
                          "No 5-year revenue growth between -30% and +80% (other drivers unchanged) makes the DCF equal "
                          "the market price; the market prices margins, returns or a discount rate outside the model.",
                          None))
    for d in drivers:
        if d.clamped:
            flags.append(Flag("DRIVER_CLAMPED", MEDIUM,
                              f"{d.label}: historical value {d.raw:.4f} was outside the allowed range and clamped to "
                              f"{d.base:.4f}.", d.raw))
    as_of = cfg.get("as_of")
    if as_of:
        age = (vdate - date.fromisoformat(as_of)).days
        if age > c["stale_days"]:
            flags.append(Flag("STALE_MARKET_INPUTS", MEDIUM,
                              f"config/market.json (risk-free rate, ERP) is {age} days old (limit {c['stale_days']}).",
                              float(age)))
    n_q, ytd = fin.ytd_growth("sales")
    if ytd is not None:
        g1 = drivers.value("sales_growth", "base", 0)
        gap = ytd - g1
        if abs(gap) > c["ytd_gap_pp"]:
            flags.append(Flag("YTD_GROWTH_GAP", MEDIUM,
                              f"Revenue grew {ytd:+.1%} in the {n_q} reported quarter(s) of FY{(fin.years[-1] + 1) % 100:02d} "
                              f"vs {g1:+.1%} forecast for the year ({gap * 100:+.1f} pp).", gap))
    if md is None or md.beta_raw is None:
        flags.append(Flag("BETA_DEFAULT", LOW, "No price history for a regression beta; default beta used.", None))
    if val.peer_ev_ebitda is None:
        flags.append(Flag("PEERS_MISSING", LOW, "No peer EV/EBITDA available; target uses the other methods.", None))
    return flags


def high_flags(flags: List[Flag]) -> List[Flag]:
    return [f for f in flags if f.severity == HIGH]
