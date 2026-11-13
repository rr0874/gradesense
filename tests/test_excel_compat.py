"""Regression checks for the college ERP and staff-output workbook variants."""
import io
import json
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")

import pandas as pd
from openpyxl import load_workbook
from backend import Dataset, run_checks
from backend.engine import Engine
from backend.report import build_generic_report, build_report

ROOT = Path(__file__).resolve().parents[1]


def ok(condition, label):
    assert condition, "FAILED: " + label


# The Sem III ERP places Term/subject headings above the student-header row,
# and uses parenthetical subject descriptors rather than course codes.
ai3 = Dataset.from_workbook(str(ROOT / "data" / "AI_3_data.xlsx"), "ai3")
ok(ai3.wb.info["students"] == 228 and ai3.wb.info["terms"] == ["III"], "shifted Sem III header is detected")
ok(ai3.wb.info["subjects"] == 24 and "DS-C (LAB)" in ai3.wb.subjects.abbr.values,
   "parenthetical subject labels are preserved")
ok(ai3.wb.marks.grade.astype(str).str.contains(r"\*$", regex=True).any(), "annotated grades are retained")
ok(all(c.ok for c in run_checks(ai3) if c.level == "error"), "annotated grades reconcile with ERP statuses")

# AI 5's raw sheet is the input paired with the supplied staff analysis tabs.
ai5 = Dataset.from_workbook(str(ROOT / "data" / "AI_5_data.xlsx"), "ai5")
ok(ai5.wb.info["students"] == 150 and ai5.wb.info["terms"] == ["V"], "AI 5 input remains supported")
ok(all(c.ok for c in run_checks(ai5) if c.level == "error"), "AI 5 values pass validation")
summary = Engine(ai5, log_path="NUL").ask("class summary").text
for expected in ("Final status: Pass 129, Fail 21", "86.0%"):
    ok(expected in summary, f"AI 5 analysis result includes {expected}")

scheme = json.loads((ROOT / "backend" / "report_schemes" / "semV.json").read_text(encoding="utf-8"))
stream = io.BytesIO()
build_report(ai5, scheme, stream)
report = load_workbook(stream, data_only=False)
ok({"Result_Analysis", "subjectwise", "Data", "Notes"} <= set(report.sheetnames),
   "staff-style analysis and auditable data sheets are produced")
ra, sw, data = report["Result_Analysis"], report["subjectwise"], report["Data"]
ok(ra["A1"].value == "Sr No" and ra["AD1"].value == "CGPA", "analysis headers follow the staff workbook")
ok(ra["B6"].value == ai5.wb.students.enroll.iloc[0] and ra["AD6"].value == ai5.wb.cumulative.cgpa.iloc[0],
   "analysis rows carry the source student and CGPA values")
ok(isinstance(ra["E6"].value, str) and ra["E6"].value.startswith("=IF(COUNTIFS("),
   "subject totals are linked to the auditable source data")
ok(ra["AG6"].value.startswith('=IF(AF6<>"Pass","Fail"') and "AD6>=8" in ra["AG6"].value,
   "overall class formulas use the student's own status and CGPA")
ok(ra["H159"].value == "Toppers (by CGPA)" and "LARGE(" in ra["I160"].value,
   "overall output includes formula-ranked CGPA toppers")
ok(sw["E8"].value == "ML I" and sw["E9"].value == 100 and sw["E10"].value.startswith("=COUNTIFS("),
   "subjectwise output includes labels, maxima, and live counts")
ok(data.max_row == len(ai5.wb.marks) + 1 and data["H2"].value in ("Y", "N"),
   "one auditable output row is generated per parsed student-subject result")

# A deliberately different, normalized long-table workbook is generated in memory.
# It has no semester catalog or per-workbook report scheme.
from openpyxl import Workbook as InputWorkbook
source = InputWorkbook()
sheet = source.active
sheet.append(["Student ID", "Student Name", "Semester", "Course Code", "Subject Name", "Marks Obtained",
              "Maximum Marks", "Grade", "Result", "Credits", "SGPA", "CGPA", "Percentage"])
for record in [
    ("U1001", "Student Alpha", "I", "CS101", "Computing", 88, 100, "A", "Pass", 4, 9.5, 7.0, 88),
    ("U1001", "Student Alpha", "I", "MA101", "Mathematics", 74, 100, "B", "Pass", 3, 9.5, 7.0, 74),
    ("U1002", "Student Beta", "I", "CS101", "Computing", 41, 100, "F", "Fail", 4, 5.5, 9.0, 41),
    ("U1002", "Student Beta", "I", "MA101", "Mathematics", 61, 100, "C", "Pass", 3, 5.5, 9.0, 61),
]:
    sheet.append(record)
buffer = io.BytesIO()
source.save(buffer)
flat = Dataset.from_workbook(buffer.getvalue(), "independent-flat-fixture")
ok(flat.wb.info["students"] == 2 and flat.wb.info["subjects"] == 2, "normalized long-table input is detected")
ok(flat.wb.terms.sgpa.notna().all() and flat.wb.cumulative.cgpa.notna().all(),
   "semester SGPA and cumulative CGPA remain separate")
flat_engine = Engine(flat, catalogs={}, log_path="NUL")
top_cgpa = flat_engine.ask("top 2 students by CGPA").table
top_default = flat_engine.ask("top 2 students").table
ok(flat_engine.ask("top 2 students by sgpa").table.iloc[0]["Enrollment"] == "U1002", "a file that has CGPA is ranked only by CGPA, even if SGPA is named")
ok(top_cgpa.iloc[0]["Enrollment"] == "U1002" and top_cgpa.iloc[0]["CGPA"] == 9.0, "ranking uses CGPA")
ok(top_default.iloc[0]["Enrollment"] == "U1002" and "CGPA" in top_default.columns, "with no metric named, the file's CGPA is the ranking")
ok("CGPA" in flat_engine.ask("class summary by CGPA").text, "class summary uses explicitly requested CGPA")
generic = io.BytesIO()
build_generic_report(flat, "I", generic)
generic_wb = load_workbook(generic, data_only=True)
ok(generic_wb.sheetnames == ["Result_Analysis", "SUBJECTWISE", "Counts", "Failed_Students"], "report has Result_Analysis, SUBJECTWISE, Counts and Failed_Students")
rs = generic_wb["Result_Analysis"]
ok(rs.freeze_panes == "D8" and rs["A3"].value == "Rank" and rs["B3"].value == "Student Unique No",
   "staff-style header band with frozen identifiers")
ok(rs["B8"].value == "U1002" and rs["B9"].value == "U1001",
   "file has CGPA, so ranking uses CGPA only (SGPA would put U1001 first)")
heads = {c.value: c.column for c in rs[3] if c.value}
ok("CGPA" in heads and "SGPA" not in heads, "ranking column shown is the metric actually used")
ok(rs.cell(8, heads["CGPA"]).value == 9.0 and rs.cell(8, heads["Class"]).value == "Distinction", "CGPA and class columns are filled")
ok(not any(c.data_type == "f" for row in rs.iter_rows() for c in row), "values are plain numbers, no formulas")
ok(all(c.fill.fill_type is None for row in rs.iter_rows(min_row=8) for c in row), "student rows carry no colour; only the header band is highlighted")
fs = generic_wb["Failed_Students"]
ok(fs["B8"].value == "U1002" and fs["A8"].value == rs["A8"].value and fs.max_row == 8,
   "failed students are copied as complete rows")
cn = generic_wb["Counts"]
cv = {cn.cell(r, 1).value: cn.cell(r, 5).value for r in range(5, 14)}
ok(cv["Total students"] == 2 and cv["No. of Students Failed / ATKT"] == 1 and cv["No. of Students Passed"] == 1 and cv["Overall pass percentage"] == 0.5,
   "Counts sheet totals come from the parsed results")
ok(cn["H6"].value == "Student Beta" and cn["G5"].value == "Top 3 by CGPA", "Counts sheet lists the top students by CGPA")
sw = generic_wb["SUBJECTWISE"]
ok(sw["A10"].value == "No of Students Appeared" and sw["E10"].value == 2 and sw["A19"].value.startswith("Subject percentage"),
   "SUBJECTWISE has the staff row labels and per-subject counts")
# a file with no CGPA anywhere is ranked by SGPA
flat.wb.cumulative["cgpa"] = float("nan")
nocg = io.BytesIO(); build_generic_report(flat, "I", nocg)
ns = load_workbook(nocg)["Result_Analysis"]
ok(ns["B8"].value == "U1001" and "SGPA" in [c.value for c in ns[3]], "without CGPA the report falls back to SGPA")

print("Excel compatibility tests: plain CGPA-first report passed")
