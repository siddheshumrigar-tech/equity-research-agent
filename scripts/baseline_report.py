"""Print a valuation baseline table (used to compare calls before and after model changes).

Usage:
    python scripts/baseline_report.py TCS RELIANCE ASIANPAINT TATATECH MARUTI
    python scripts/baseline_report.py --offline --cache-dir tests/fixtures/screener TATATECH ASIANPAINT
    python scripts/baseline_report.py --markdown docs/superpowers/specs/baseline.md TCS RELIANCE

Columns: DCF value, P/E value, target, upside, verdict, PV of terminal value as a share
of EV, and DCF / P/E value. Works with any version of the valuation engine (new fields
are read with getattr).
"""
from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from equity_agent.errors import EquityAgentError  # noqa: E402
from equity_agent.pipeline import analyse  # noqa: E402

COLS = ["Symbol", "CMP", "DCF", "P/E value", "Target", "Upside", "Rating", "TV % EV", "DCF / P/E"]


def row_for(symbol: str, cache_dir: str, offline: bool) -> list:
    try:
        res = analyse(symbol, cache_dir=cache_dir, offline=offline)
    except EquityAgentError as exc:
        return [symbol, "error", exc.__class__.__name__] + [""] * (len(COLS) - 3)
    v = res.valuation
    tv_share = v.pv_terminal_value / v.enterprise_value if v.enterprise_value else None
    ratio = v.dcf_value / v.pe_value if v.pe_value else None
    rating = getattr(v, "rating", None) or v.verdict
    return [symbol, f"{v.price:,.0f}", f"{v.dcf_value:,.0f}", f"{v.pe_value:,.0f}" if v.pe_value else "n/a",
            f"{v.target_price:,.0f}", f"{v.upside:+.1%}", rating,
            f"{tv_share:.0%}" if tv_share is not None else "n/a", f"{ratio:.2f}" if ratio is not None else "n/a"]


def to_markdown(rows) -> str:
    out = ["| " + " | ".join(COLS) + " |", "|" + "---|" * len(COLS)]
    out += ["| " + " | ".join(str(c) for c in r) + " |" for r in rows]
    return "\n".join(out)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    ap.add_argument("symbols", nargs="*", default=["TCS", "RELIANCE", "ASIANPAINT", "TATATECH", "MARUTI"])
    ap.add_argument("--cache-dir", default=str(ROOT / "cache"))
    ap.add_argument("--offline", action="store_true")
    ap.add_argument("--markdown", help="Also write the table to this markdown file")
    ap.add_argument("--title", default="Valuation baseline")
    a = ap.parse_args(argv)
    rows = [row_for(s.upper(), a.cache_dir, a.offline) for s in a.symbols]
    table = to_markdown(rows)
    print(table)
    if a.markdown:
        p = Path(a.markdown)
        p.parent.mkdir(parents=True, exist_ok=True)
        prev = p.read_text(encoding="utf-8") if p.exists() else ""
        p.write_text(prev + f"\n## {a.title}\n\n{table}\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    sys.exit(main())
