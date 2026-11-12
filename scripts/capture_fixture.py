"""Save live Screener.in / yfinance responses for offline tests.

Usage:
    python scripts/capture_fixture.py TATATECH ASIANPAINT HDFCBANK      # pages + schedules + market + peers
    python scripts/capture_fixture.py --market-peers-only TATATECH ASIANPAINT

Writes into tests/fixtures/screener/<SYMBOL>/..., the same layout as the runtime
cache, so tests can use ``ScreenerClient(cache_dir=..., offline=True)``. Peer
company pages are cached in ./cache; only the compact peer snapshot
(``<SYMBOL>/peers/peers.json``) goes into the fixtures.
"""
import argparse
import json
import shutil
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from equity_agent.config import load_market_config  # noqa: E402
from equity_agent.errors import BankNotSupported  # noqa: E402
from equity_agent.market_data import download_market_data  # noqa: E402
from equity_agent.peers import get_peers  # noqa: E402
from equity_agent.screener_client import ScreenerClient  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures" / "screener"


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("symbols", nargs="*", default=["TATATECH", "ASIANPAINT", "HDFCBANK"])
    ap.add_argument("--market-peers-only", action="store_true", help="Keep saved pages; refresh market data and peers")
    a = ap.parse_args(argv)
    cfg = load_market_config()
    fx = ScreenerClient(cache_dir=FIXTURES, refresh=not a.market_peers_only, ttl_seconds=None if a.market_peers_only else 0)
    live = ScreenerClient(cache_dir=ROOT / "cache")
    for sym in a.symbols:
        try:
            raw = fx.fetch_company(sym)
        except BankNotSupported as exc:
            print(f"{sym}: saved page (bank) -> {exc}")
            continue
        if raw.logo_url and not a.market_peers_only:
            fx.fetch_binary(raw.symbol, "screener_logo", raw.logo_url)
        md = download_market_data(raw.symbol, cfg)
        mpath = FIXTURES / raw.symbol / "market" / "yfinance.json"
        mpath.parent.mkdir(parents=True, exist_ok=True)
        mpath.write_text(json.dumps(md.to_dict(), indent=1), encoding="utf-8")
        live_peer = ROOT / "cache" / raw.symbol / "peers" / "peers.json"
        if live_peer.exists():
            live_peer.unlink()
        ps = get_peers(live, raw, cfg)
        if ps is not None:
            dst = FIXTURES / raw.symbol / "peers" / "peers.json"
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy(live_peer, dst)
        print(f"{sym}: {raw.name}; {len(md.price_close)} daily prices; "
              f"{len(ps.peers) if ps else 0} peers ({', '.join(p.symbol for p in ps.peers) if ps else '-'})")


if __name__ == "__main__":
    main()
