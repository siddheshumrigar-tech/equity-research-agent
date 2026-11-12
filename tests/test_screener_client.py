import pytest

from equity_agent.errors import BankNotSupported, NetworkError, SymbolNotFound
from equity_agent.screener_client import ScreenerClient, is_bank_layout, parse_number, parse_schedule


@pytest.mark.parametrize("text,expected", [
    ("1,234", 1234.0), ("-39", -39.0), ("17%", 17.0), ("17.57%", 17.57), ("", None), (None, None),
    ("₹ 2,45,348 Cr.", 245348.0), ("+15%", 15.0), ("₹ 676", 676.0), ("-1,328", -1328.0), ("  ", None),
])
def test_parse_number(text, expected):
    assert parse_number(text) == expected


def test_parse_tatatech_page(raw_tata):
    assert raw_tata.name == "Tata Technologies Ltd"
    assert raw_tata.company_id == "1284568"
    assert raw_tata.consolidated is True
    assert raw_tata.website.startswith("https://www.tatatechnologies.com")
    assert raw_tata.bse_code == "544028"
    assert "company-logos" in raw_tata.logo_url
    assert raw_tata.sector_path[-1] == "IT Enabled Services"
    pl = raw_tata.tables["profit-loss"]
    assert pl.headers[:2] == ["Mar 2019", "Mar 2020"] and pl.headers[-1] == "TTM"
    assert pl.rows["Sales"][pl.headers.index("Mar 2025")] == 5168
    assert pl.schedules["Expenses"] == "Expenses"
    bs = raw_tata.tables["balance-sheet"]
    assert bs.rows["Total Assets"][bs.headers.index("Mar 2026")] == 8860


def test_schedules_loaded(raw_tata):
    exp = raw_tata.schedule("profit-loss", "Expenses")
    assert set(exp) == {"Material Cost", "Employee Cost", "Other Expenses"}
    assert exp["Material Cost"].is_percent
    assert exp["Material Cost"].amounts["Mar 2026"] == 967
    oa = raw_tata.schedule("balance-sheet", "Other Assets")
    assert oa["Trade receivables"].values["Mar 2026"] == 1187


def test_parse_schedule_skips_meta_keys():
    payload = {"Trade Payables": {"Mar 2025": "1,234", "Sep 2025": "", "setAttributes": {"class": "x"}},
               "Material Cost %": {"Mar 2025": "12.5%", "tooltips": {"Mar 2025": "₹1,390 Cr"}, "isExpandable": "x"}}
    out = parse_schedule(payload)
    assert out["Trade Payables"].values == {"Mar 2025": 1234.0, "Sep 2025": None}
    assert out["Material Cost"].amounts == {"Mar 2025": 1390.0}


def test_bank_is_rejected(client):
    with pytest.raises(BankNotSupported):
        client.fetch_company("HDFCBANK")
    assert BankNotSupported.exit_code == 3


def test_invalid_symbol_text(client):
    with pytest.raises(SymbolNotFound):
        client.fetch_company("NOT A SYMBOL!")


def test_offline_cache_miss_is_network_error(client):
    with pytest.raises(NetworkError):
        client.fetch_company("NOSUCHCOMPANY")


class _Resp:
    def __init__(self, status, text=""):
        self.status_code = status
        self.text = text
        self.content = text.encode()


class _Session:
    def __init__(self):
        self.headers = {}
        self.calls = []

    def get(self, url, params=None, timeout=None):
        self.calls.append(url)
        return _Resp(404)


def test_unknown_symbol_404(tmp_path):
    sess = _Session()
    c = ScreenerClient(cache_dir=tmp_path, session=sess, min_interval=0)
    with pytest.raises(SymbolNotFound):
        c.fetch_company("NOSUCHCO")
    assert any("/company/NOSUCHCO/consolidated/" in u for u in sess.calls)
    assert any(u.endswith("/company/NOSUCHCO/") for u in sess.calls)


def test_bank_layout_detection(raw_tata):
    assert not is_bank_layout(raw_tata.tables)
