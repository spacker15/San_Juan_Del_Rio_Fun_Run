"""Build every scenario workbook once, recalculate them all in one LibreOffice call, share results."""
import os
import shutil
import subprocess
import sys
import tempfile
from datetime import date, datetime

import openpyxl
import pytest

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.join(ROOT, "build"))
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import build_workbook  # noqa: E402
import demo_data  # noqa: E402
import scenarios  # noqa: E402

SCENARIOS = {
    "blank": lambda: {},
    "demo": demo_data.generate,
    "draw": lambda: dict(demo_data.generate(), challenge_pick=4),
    "ties": scenarios.ties,
    **{f"thermo_{amt}": (lambda a=amt: scenarios.thermo(a)) for amt in scenarios.THERMO_AMOUNTS},
}


class Book:
    """Recalculated workbook values + the input data that produced them."""

    def __init__(self, path, data):
        self.wb = openpyxl.load_workbook(path, data_only=True)
        self.data = data

    def __getitem__(self, sheet):
        return self.wb[sheet]

    def kpi(self, nm):
        ws = self.wb["Calc_KPI"]
        for r in range(2, ws.max_row + 1):
            if str(ws[f"A{r}"].value or "").split("  ")[0] == nm:
                return ws[f"B{r}"].value
        raise KeyError(nm)


def num(v):
    """Cell value as a plain number (dates come back as datetimes when the cell has a date format)."""
    if isinstance(v, datetime):
        return (v.date() - date(1899, 12, 30)).days
    if isinstance(v, date):
        return (v - date(1899, 12, 30)).days
    return v


@pytest.fixture(scope="session")
def books():
    if not shutil.which("soffice"):
        pytest.skip("LibreOffice (soffice) is required to recalculate formulas")
    tmp = tempfile.mkdtemp(prefix="srw_")
    src, out = os.path.join(tmp, "src"), os.path.join(tmp, "out")
    os.makedirs(src)
    datas, paths = {}, []
    for key, make in SCENARIOS.items():
        datas[key] = make()
        paths.append(build_workbook.build(os.path.join(src, f"{key}.xlsx"), datas[key]))
    subprocess.run(["soffice", "--headless", "--calc", "--convert-to", "xlsx", "--outdir", out, *paths],
                   check=True, capture_output=True, timeout=1800)
    return {k: Book(os.path.join(out, f"{k}.xlsx"), datas[k]) for k in SCENARIOS}
