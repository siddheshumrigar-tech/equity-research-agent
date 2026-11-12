import io

import pytest
from PIL import Image

from equity_agent.brand import FALLBACK_HEX, dominant_colour, get_brand, palette_from_hex, text_on


def _png(img):
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    return buf.getvalue()


def test_red_logo_gives_red():
    img = Image.new("RGB", (64, 64), (255, 255, 255))
    for x in range(10, 54):
        for y in range(10, 54):
            img.putpixel((x, y), (200, 20, 30))
    hex_ = dominant_colour(img)
    r, g, b = int(hex_[0:2], 16), int(hex_[2:4], 16), int(hex_[4:6], 16)
    assert r > 150 and g < 80 and b < 80


def test_white_and_grey_logo_has_no_colour():
    img = Image.new("RGB", (64, 64), (250, 250, 250))
    for x in range(20, 40):
        for y in range(20, 40):
            img.putpixel((x, y), (120, 120, 120))
    assert dominant_colour(img) is None


def test_transparent_pixels_ignored():
    img = Image.new("RGBA", (64, 64), (0, 200, 0, 0))  # transparent green background
    for x in range(16, 48):
        for y in range(16, 48):
            img.putpixel((x, y), (20, 40, 200, 255))
    hex_ = dominant_colour(img)
    assert int(hex_[4:6], 16) > 150 and int(hex_[2:4], 16) < 120


def test_palette_shades_and_text():
    p = palette_from_hex("0664AE", "test")
    assert p.primary == "0664AE"
    assert p.dark != p.primary and p.light != p.primary
    assert text_on("FFFFFF") == "000000" and text_on("000000") == "FFFFFF"
    assert len(p.shades(7)) == 7


class _FailingClient:
    def fetch_binary(self, *a, **k):
        return None

    def fetch_text(self, *a, **k):
        return None


class _Client:
    def __init__(self, data):
        self.data = data

    def fetch_binary(self, *a, **k):
        return self.data

    def fetch_text(self, *a, **k):
        return None


def test_network_failure_falls_back():
    p = get_brand("X", "https://example.com/logo.png", "https://example.com", _FailingClient())
    assert p.primary == FALLBACK_HEX
    assert "Fallback" in p.source


def test_logo_bytes_used():
    img = Image.new("RGB", (64, 64), (230, 120, 10))
    p = get_brand("X", "https://cdn/logo.png", None, _Client(_png(img)))
    assert p.primary != FALLBACK_HEX
    assert p.logo_png is not None


def test_fixture_logo_tatatech(raw_tata, client):
    p = get_brand(raw_tata.symbol, raw_tata.logo_url, raw_tata.website, client)
    assert p.primary == "0664AE"
