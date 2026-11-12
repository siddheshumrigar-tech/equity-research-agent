"""Recalculate an .xlsx without Excel, for tests.

Uses LibreOffice headless when ``soffice`` is on PATH (closest to Excel), else the
pure-Python ``formulas`` package. Returns ``{(SHEET_UPPER, "B5"): value}``.
"""
from __future__ import annotations

import re
import shutil
import subprocess
import tempfile
from pathlib import Path


def soffice_path():
    for name in ("soffice", "libreoffice"):
        p = shutil.which(name)
        if p:
            return p
    for p in (r"C:\Program Files\LibreOffice\program\soffice.exe", "/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice"):
        if Path(p).exists():
            return p
    return None


def _formulas_available() -> bool:
    try:
        import formulas  # noqa: F401
        return True
    except ImportError:
        return False


def available() -> bool:
    return soffice_path() is not None or _formulas_available()


def _read_values(path: Path) -> dict:
    import openpyxl

    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    out[(ws.title.upper(), c.coordinate)] = c.value
    return out


def recalc(path: str | Path) -> dict:
    path = Path(path)
    office = soffice_path()
    if office:
        with tempfile.TemporaryDirectory() as td:
            subprocess.run([office, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", td, str(path)],
                           check=True, capture_output=True, timeout=180)
            return _read_values(Path(td) / path.name)
    import formulas

    xl = formulas.ExcelModel().loads(str(path)).finish()
    sol = xl.calculate()
    out = {}
    pat = re.compile(r"^'?\[[^\]]+\](.+?)'?!([A-Z]+\d+)$")
    for k, v in sol.items():
        m = pat.match(str(k))
        if not m:
            continue
        try:
            val = v.value[0, 0]
        except Exception:
            continue
        if hasattr(val, "item"):
            val = val.item()
        out[(m.group(1).upper(), m.group(2))] = val
    return out
