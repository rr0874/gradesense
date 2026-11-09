"""Regression checks. Run from the project folder:  python -m tests.test_backend  [path/to/scheme_image.png]"""
import sys, warnings; warnings.filterwarnings("ignore")
from unittest.mock import patch
from backend import Dataset, run_checks
from backend.catalog import apply_catalog
from backend.syllabus import COLS, _is_horizontal_dash, _is_leading_serial_column
import numpy as np

dash = np.full((12, 30), 255, dtype=np.uint8)
dash[5:7, 8:19] = 0
digit = np.full((12, 30), 255, dtype=np.uint8)
digit[2:11, 12:15] = 0
assert _is_horizontal_dash(dash) and not _is_horizontal_dash(digit), "numeric dash cells are separated from narrow digits"

columns = [(0, 20), (20, 120)] + [(120 + i * 20, 140 + i * 20) for i in range(len(COLS) - 1)]
rows = [(i * 10, (i + 1) * 10) for i in range(6)]
with patch("backend.syllabus._ocr_cells", side_effect=lambda _g, cells, col: (
           (["231AIMLPCC301"] * len(cells) if col == "code" else
            [str(i + 1) for i in range(len(cells))]), set())):
    assert _is_leading_serial_column(None, columns, rows, set()), "sequential leading row numbers are recognized"

ds = Dataset.from_workbook("data/AIML_4.xlsx", "data/AIML_4.xlsx")
assert ds.wb.info["students"] == 153 and ds.wb.info["terms"] == ["I", "II", "III", "IV"], ds.wb.info
assert all(c.ok for c in run_checks(ds) if c.level == "error"), "validation reports a data error"
print("ingest + validation: OK")

if len(sys.argv) > 1:
    from backend.syllabus import extract_scheme
    sch = extract_scheme(open(sys.argv[1], "rb").read())
    by = {c["code"]: c for c in sch.courses}
    assert sch.semester == "IV" and len(by) == 11 and sch.pass_pct == 40
    assert (by["231AIMLPCC204"]["ISE"], by["231AIMLPCC204"]["total"], by["231AIMLAEC201"]["L"], by["231AIMLEEM202"]["L"]) == (20, 125, 1, 1)
    assert "P" not in by["231AIMLVEC202"] and "P" not in by["231AIMLMDM202"]
    s, miss = apply_catalog(ds.wb.subjects, sch.to_catalog())
    assert s.in_catalog.sum() == 11
    print("syllabus extraction + catalog matching: OK", "| review items:", len(sch.problems))

# second layout (max marks / credits / pass rule printed above the columns; INT/POE/single-component blocks; no cumulative columns)
ds2 = Dataset.from_workbook("data/Sample_Student_Results_Gen_AI.xlsx", "gen-ai")
assert ds2.wb.info["students"] == 50 and ds2.wb.info["subjects"] == 8, ds2.wb.info
assert all(c.ok for c in run_checks(ds2) if c.level == "error"), "validation reports a data error on the second layout"
s2 = ds2.wb.subjects.set_index("abbr")
assert (s2.at["Gen AI", "max_total"], s2.at["Gen AI", "credits"], s2.at["RM", "max_int"], s2.at["PR", "max_oe_poe"]) == (100, 3, 25, 50)
assert "ISE+MSE 40%" in s2.at["Gen AI", "pass_rule"]
print("second layout ingest + validation: OK")
