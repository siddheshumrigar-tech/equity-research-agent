"""Company brand colour from its logo, and the palette used by the Dashboard and PDF.

Logo sources, in order: the logo Screener shows on the company page, then the
company website's apple-touch-icon / icon / og:image, then Google's favicon
service. The dominant colour is the most frequent of 8 quantized colours with
saturation > 0.25 and 0.15 < lightness < 0.85 (transparent pixels ignored).
If no logo yields such a colour (white, grey or SVG-only logos), Tesla green
70AD47 is used.
"""
from __future__ import annotations

import colorsys
import io
import re
import sys
from dataclasses import dataclass
from typing import List, Optional, Tuple
from urllib.parse import urljoin, urlparse

FALLBACK_HEX = "70AD47"


def _rgb(hex_: str) -> Tuple[int, int, int]:
    h = hex_.lstrip("#")
    return int(h[0:2], 16), int(h[2:4], 16), int(h[4:6], 16)


def _hex(rgb) -> str:
    return "".join(f"{max(0, min(255, int(round(c)))):02X}" for c in rgb)


def _hls_adjust(hex_: str, l_mul: float = 1.0, l_add: float = 0.0) -> str:
    r, g, b = (c / 255.0 for c in _rgb(hex_))
    h, l, s = colorsys.rgb_to_hls(r, g, b)
    l = min(1.0, max(0.0, l * l_mul + l_add))
    return _hex(tuple(c * 255 for c in colorsys.hls_to_rgb(h, l, s)))


def text_on(hex_: str) -> str:
    r, g, b = (c / 255.0 for c in _rgb(hex_))
    lum = 0.2126 * r + 0.7152 * g + 0.0722 * b
    return "000000" if lum > 0.55 else "FFFFFF"


@dataclass
class Palette:
    primary: str
    dark: str
    light: str
    card_bg: str
    source: str
    logo_png: Optional[bytes] = None

    @property
    def text_on_primary(self) -> str:
        return text_on(self.primary)

    @property
    def text_on_dark(self) -> str:
        return text_on(self.dark)

    @property
    def text_on_light(self) -> str:
        return text_on(self.light)

    def shades(self, n: int) -> List[str]:
        """``n`` shades of the primary colour, darkest first (for pie slices)."""
        if n <= 1:
            return [self.primary]
        out = []
        for i in range(n):
            t = i / (n - 1)  # 0 -> dark, 1 -> light
            out.append(_hls_adjust(self.primary, l_mul=0.65 + 0.25 * t, l_add=0.45 * t))
        return out


def palette_from_hex(hex_: str, source: str, logo_png: Optional[bytes] = None) -> Palette:
    hex_ = hex_.lstrip("#").upper()
    return Palette(primary=hex_, dark=_hls_adjust(hex_, l_mul=0.75), light=_hls_adjust(hex_, l_mul=0.6, l_add=0.4),
                   card_bg="F2F2F2", source=source, logo_png=logo_png)


def dominant_colour(image) -> Optional[str]:
    """Return the dominant brand-like colour of a PIL image as hex, or None."""
    from PIL import Image

    img = image.convert("RGBA")
    img.thumbnail((128, 128))
    data = img.get_flattened_data() if hasattr(img, "get_flattened_data") else img.getdata()
    px = [tuple(p[:3]) for p in data if p[3] >= 128]
    if len(px) < 20:
        return None
    strip = Image.new("RGB", (len(px), 1))
    strip.putdata(px)
    q = strip.quantize(colors=8, method=Image.Quantize.MEDIANCUT)
    pal = q.getpalette()
    counts = sorted(q.getcolors(), reverse=True)
    for count, idx in counts:
        r, g, b = pal[idx * 3: idx * 3 + 3]
        h, l, s = colorsys.rgb_to_hls(r / 255, g / 255, b / 255)
        if s > 0.25 and 0.15 < l < 0.85:
            return _hex((r, g, b))
    return None


def logo_to_png(image, height: int = 64, max_width: int = 150) -> bytes:
    """Flatten onto white and resize for the dashboard banner."""
    from PIL import Image

    img = image.convert("RGBA")
    bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
    bg.alpha_composite(img)
    rgb = bg.convert("RGB")
    w, h = rgb.size
    scale = min(height / h, max_width / w)
    rgb = rgb.resize((max(1, int(w * scale)), max(1, int(h * scale))), Image.LANCZOS)
    buf = io.BytesIO()
    rgb.save(buf, format="PNG")
    return buf.getvalue()


def _open(data: bytes):
    from PIL import Image

    try:
        img = Image.open(io.BytesIO(data))
        img.load()
        return img
    except Exception:
        return None


def website_icon_urls(html: str, base_url: str) -> List[str]:
    urls = []
    for rel in ("apple-touch-icon", "icon", "shortcut icon"):
        for m in re.finditer(r"<link[^>]+>", html, re.I):
            tag = m.group(0)
            r = re.search(r'rel=["\']([^"\']+)["\']', tag, re.I)
            h = re.search(r'href=["\']([^"\']+)["\']', tag, re.I)
            if r and h and r.group(1).lower().strip() == rel:
                urls.append(urljoin(base_url, h.group(1)))
    m = re.search(r'<meta[^>]+property=["\']og:image["\'][^>]+content=["\']([^"\']+)["\']', html, re.I)
    if m:
        urls.append(urljoin(base_url, m.group(1)))
    seen, out = set(), []
    for u in urls:
        if u not in seen and not u.lower().endswith(".svg"):
            seen.add(u)
            out.append(u)
    return out


def get_brand(symbol: str, logo_url: Optional[str], website: Optional[str], client) -> Palette:
    """Find the brand palette. ``client`` is a ScreenerClient (used for cached downloads)."""
    candidates: List[Tuple[str, str]] = []
    if logo_url:
        candidates.append(("Screener.in company logo", logo_url))
    if website:
        html = client.fetch_text(symbol, "website.html", website)
        if html:
            for i, u in enumerate(website_icon_urls(html, website)):
                candidates.append((f"company website icon ({urlparse(u).netloc})", u))
        domain = urlparse(website).netloc
        if domain:
            candidates.append(("Google favicon service", f"https://www.google.com/s2/favicons?domain={domain}&sz=128"))
    first_logo: Optional[bytes] = None
    for i, (label, url) in enumerate(candidates):
        data = client.fetch_binary(symbol, f"logo_{i}", url)
        if not data:
            continue
        img = _open(data)
        if img is None:
            continue
        if first_logo is None:
            first_logo = logo_to_png(img)
        colour = dominant_colour(img)
        if colour:
            return palette_from_hex(colour, f"{label}: {url}", logo_png=first_logo if i == 0 else logo_to_png(img))
    return palette_from_hex(FALLBACK_HEX, "Fallback colour (no usable logo colour found)", logo_png=first_logo)


def _main(argv):
    import argparse
    from pathlib import Path

    from .screener_client import ScreenerClient

    ap = argparse.ArgumentParser(description="Print the brand colour and save a palette swatch PNG.")
    ap.add_argument("symbol")
    ap.add_argument("--cache-dir", default="cache")
    ap.add_argument("--offline", action="store_true")
    ap.add_argument("--out", default="output")
    a = ap.parse_args(argv)
    client = ScreenerClient(cache_dir=a.cache_dir, offline=a.offline)
    raw = client.fetch_company(a.symbol)
    pal = get_brand(raw.symbol, raw.logo_url, raw.website, client)
    print(f"{raw.name}: primary #{pal.primary} dark #{pal.dark} light #{pal.light}  ({pal.source})")
    from PIL import Image, ImageDraw

    sw = Image.new("RGB", (400, 100), "white")
    d = ImageDraw.Draw(sw)
    for i, c in enumerate([pal.primary, pal.dark, pal.light, pal.card_bg]):
        d.rectangle([i * 100, 0, i * 100 + 99, 99], fill="#" + c)
    Path(a.out).mkdir(parents=True, exist_ok=True)
    path = Path(a.out) / f"{raw.symbol}_palette.png"
    sw.save(path)
    print(f"Swatch saved to {path}")
    return 0


if __name__ == "__main__":
    sys.exit(_main(sys.argv[1:]))
