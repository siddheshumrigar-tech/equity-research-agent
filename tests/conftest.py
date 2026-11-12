import os
import sys
from pathlib import Path

import pytest

os.environ["EQUITY_AGENT_NO_HISTORY"] = "1"  # tests must not write to memory/learnings.json

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

FIXTURES = ROOT / "tests" / "fixtures" / "screener"


@pytest.fixture(scope="session")
def fixtures_dir():
    return FIXTURES


@pytest.fixture(scope="session")
def client():
    from equity_agent.screener_client import ScreenerClient

    return ScreenerClient(cache_dir=FIXTURES, offline=True)


@pytest.fixture(scope="session")
def raw_tata(client):
    return client.fetch_company("TATATECH")


@pytest.fixture(scope="session")
def raw_asian(client):
    return client.fetch_company("ASIANPAINT")


@pytest.fixture(scope="session")
def fin_tata(raw_tata):
    from equity_agent.financials import build_financials

    return build_financials(raw_tata)


@pytest.fixture(scope="session")
def fin_asian(raw_asian):
    from equity_agent.financials import build_financials

    return build_financials(raw_asian)


def _analyse(symbol):
    from equity_agent.pipeline import analyse

    return analyse(symbol, cache_dir=FIXTURES, offline=True)


@pytest.fixture(scope="session")
def res_tata():
    return _analyse("TATATECH")


@pytest.fixture(scope="session")
def res_asian():
    return _analyse("ASIANPAINT")


@pytest.fixture(scope="session")
def out_dir(tmp_path_factory):
    return tmp_path_factory.mktemp("outputs")


@pytest.fixture(scope="session")
def built_tata(res_tata, out_dir):
    from equity_agent.pipeline import write_outputs

    return write_outputs(res_tata, out_dir / "tata", make_pdf=True)


@pytest.fixture(scope="session")
def built_asian(res_asian, out_dir):
    from equity_agent.pipeline import write_outputs

    return write_outputs(res_asian, out_dir / "asian", make_pdf=True)


@pytest.fixture(scope="session")
def recalc_tata(built_tata):
    import recalc

    if not recalc.available():
        pytest.skip("No recalculation engine (LibreOffice or the 'formulas' package) installed")
    return recalc.recalc(built_tata.xlsx_path)
