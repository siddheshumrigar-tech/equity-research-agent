#!/usr/bin/env python3
"""Equity research agent CLI (NSE/BSE stocks, Screener.in data).

    python generate_equity_report.py --ticker TATATECH

Writes to ./output:
    {SYMBOL}_Financial_Model.xlsx      Tesla-style dashboard + 3-statement model
    {SYMBOL}_Equity_Research_Report.pdf
    {SYMBOL}_summary.json              machine-readable result for the calling agent

    {SYMBOL}_research_pack.json        numbers, flags and schemas for the host agent's narrative

Two-step agent workflow: run once (draft), write {SYMBOL}_narrative.json (and optionally
{SYMBOL}_analyst_inputs.json) from the research pack, then run again with --render.

Exit codes: 0 ok, 2 symbol not found, 3 bank/financial not supported,
4 data validation failed, 5 network failure, 6 narrative / analyst inputs invalid, 1 other error.
On failure a summary JSON with "status": "error" is still written.
"""
from __future__ import annotations

import argparse
import json
import os
import smtplib
import sys
import traceback
from datetime import datetime
from email.message import EmailMessage
from pathlib import Path

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

from equity_agent.errors import EquityAgentError  # noqa: E402
from equity_agent.pipeline import analyse, summary_dict, write_outputs  # noqa: E402


def _log_run(res) -> None:
    if os.environ.get("EQUITY_AGENT_NO_HISTORY"):
        return
    try:
        from memory.memory_manager import AgentMemoryManager

        v = res.valuation
        AgentMemoryManager().log_research_run(res.fin.symbol, res.fin.name, res.fin.price, v.target_price, v.verdict,
                                              f"{v.upside:+.1%}")
    except Exception as exc:  # history logging must never break a run
        print(f"note: research history not logged ({exc})", file=sys.stderr)


def send_email(to_addr: str, files, subject: str, body: str) -> str:
    """Send outputs via SMTP using SMTP_HOST, SMTP_PORT, SMTP_USER, SMTP_PASSWORD, SMTP_FROM env vars."""
    host = os.environ.get("SMTP_HOST")
    if not host:
        return "skipped: SMTP_HOST not set"
    port = int(os.environ.get("SMTP_PORT", "587"))
    user, pwd = os.environ.get("SMTP_USER"), os.environ.get("SMTP_PASSWORD")
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = os.environ.get("SMTP_FROM", user or "equity-agent@localhost")
    msg["To"] = to_addr
    msg.set_content(body)
    for f in files:
        if f and Path(f).exists():
            data = Path(f).read_bytes()
            sub = "pdf" if f.endswith(".pdf") else "vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            maintype = "application"
            msg.add_attachment(data, maintype=maintype, subtype=sub, filename=Path(f).name)
    with smtplib.SMTP(host, port, timeout=60) as s:
        s.starttls()
        if user and pwd:
            s.login(user, pwd)
        s.send_message(msg)
    return f"sent to {to_addr}"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Generate an Excel model + PDF report for an NSE/BSE stock from Screener.in data.")
    ap.add_argument("--ticker", required=True, help="NSE symbol as used on Screener.in, e.g. TATATECH, ASIANPAINT")
    ap.add_argument("--output-dir", "--output_dir", dest="output_dir", default=str(ROOT / "output"))
    ap.add_argument("--cache-dir", default=str(ROOT / "cache"), help="Screener/yfinance response cache (24 h)")
    ap.add_argument("--refresh", action="store_true", help="Ignore the cache and download fresh data")
    ap.add_argument("--offline", action="store_true", help="Use cached data only (no network)")
    ap.add_argument("--no-pdf", action="store_true", help="Skip the PDF report")
    ap.add_argument("--email", help="Optional: email the outputs (needs SMTP_* environment variables)")
    ap.add_argument("--render", action="store_true",
                    help="Step 2: rebuild from cached data with <SYMBOL>_narrative.json (required) and "
                         "<SYMBOL>_analyst_inputs.json (optional) from the output folder")
    ap.add_argument("--narrative", help="Path to narrative.json (default: <output-dir>/<SYMBOL>_narrative.json)")
    ap.add_argument("--analyst-inputs", help="Path to analyst_inputs.json (default: <output-dir>/<SYMBOL>_analyst_inputs.json)")
    args = ap.parse_args(argv)

    out = Path(args.output_dir)
    out.mkdir(parents=True, exist_ok=True)
    symbol = args.ticker.strip().upper()
    for suffix in (".NS", ".BO"):
        if symbol.endswith(suffix):
            symbol = symbol[: -len(suffix)]
    ai_path = Path(args.analyst_inputs) if args.analyst_inputs else out / f"{symbol}_analyst_inputs.json"
    try:
        res = analyse(symbol, cache_dir=args.cache_dir, offline=args.offline, refresh=args.refresh and not args.render,
                      analyst_inputs_path=ai_path if ai_path.exists() else None, frozen_cache=args.render)
        if args.render:
            from equity_agent.checks import high_flags
            from equity_agent.narrative import load_narrative

            nar_path = Path(args.narrative) if args.narrative else out / f"{symbol}_narrative.json"
            res.narrative = load_narrative(nar_path, res, [f.code for f in high_flags(res.flags)])
            res.warnings += [f"Narrative lint: {w}" for w in res.narrative.lint]
        write_outputs(res, out, make_pdf=not args.no_pdf)
    except EquityAgentError as exc:
        _write_error(out, symbol, exc.__class__.__name__, str(exc), exc.exit_code)
        print(f"ERROR ({exc.__class__.__name__}): {exc}", file=sys.stderr)
        return exc.exit_code
    except Exception as exc:
        _write_error(out, symbol, exc.__class__.__name__, str(exc), 1)
        traceback.print_exc()
        return 1

    summary = summary_dict(res)
    if args.email:
        try:
            summary["email"] = send_email(args.email, [res.xlsx_path, res.pdf_path],
                                          f"Equity research: {res.fin.name} ({res.fin.symbol})",
                                          _digest(res))
        except Exception as exc:
            summary["email"] = f"failed: {exc}"
        Path(res.summary_path).write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    _log_run(res)
    print(_digest(res))
    return 0


def _digest(res) -> str:
    f, v = res.fin, res.valuation
    lines = [
        f"{f.name} ({f.symbol}) - {f.sector} / {f.industry}",
        f"Data: Screener.in {'consolidated' if f.consolidated else 'standalone'}, FY{f.years[0] % 100:02d}-FY{f.years[-1] % 100:02d} reported",
        f"Price Rs {f.price:,.2f} | Market cap Rs {f.market_cap:,.0f} Cr",
        f"FY{f.years[-1] % 100:02d}: revenue Rs {f.latest('sales'):,.0f} Cr, EBITDA Rs {f.latest('ebitda'):,.0f} Cr, "
        f"net profit Rs {f.latest('net_profit'):,.0f} Cr",
        f"WACC {v.wacc:.2%} | DCF (12m) Rs {v.dcf_value_12m:,.2f} | P/E value "
        + (f"Rs {v.pe_value:,.2f}" if v.pe_value is not None else "n/a") + " | EV/EBITDA value "
        + (f"Rs {v.ev_ebitda_value:,.2f}" if v.ev_ebitda_value is not None else "n/a"),
        f"12-month target Rs {v.target_price:,.2f} ({v.upside:+.1%}; expected total return "
        f"{v.expected_total_return:+.1%}) -> {v.rating}",
        f"Report stage: {'final' if res.is_final else 'DRAFT (write the narrative, then run --render)'}",
        f"Excel: {res.xlsx_path}",
        f"PDF:   {res.pdf_path or 'not generated'}",
        f"JSON:  {res.summary_path}",
        f"Pack:  {res.pack_path}",
    ]
    hi = [fl for fl in res.flags if fl.severity == "high"]
    if hi:
        lines.append("High-severity model flags (answer in narrative.json flag_responses):")
        lines += [f"  - {fl.code}: {fl.message}" for fl in hi]
    if res.warnings:
        lines.append("Warnings:")
        lines += [f"  - {w}" for w in res.warnings]
    return "\n".join(lines)


def _write_error(out: Path, symbol: str, kind: str, message: str, code: int):
    data = {"status": "error", "symbol": symbol, "error_type": kind, "message": message, "exit_code": code,
            "generated_at": datetime.now().isoformat(timespec="seconds")}
    (out / f"{symbol}_summary.json").write_text(json.dumps(data, indent=2), encoding="utf-8")


if __name__ == "__main__":
    sys.exit(main())
