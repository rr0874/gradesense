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
top_sgpa = flat_engine.ask("top 2 students").table
ok(top_cgpa.iloc[0]["Enrollment"] == "U1002" and top_cgpa.iloc[0]["CGPA"] == 9.0
   and top_sgpa.iloc[0]["Enrollment"] == "U1001" and top_sgpa.iloc[0]["SGPA"] == 9.5,
   "explicit CGPA ranking and default SGPA ranking use their named metrics")
ok("CGPA" in flat_engine.ask("class summary by CGPA").text, "class summary uses explicitly requested CGPA")
generic = io.BytesIO()
build_generic_report(flat, "I", generic)
generic_wb = load_workbook(generic, data_only=True)
ok({"Result_Analysis", "subjectwise", "Data", "Notes"} <= set(generic_wb.sheetnames),
   "scheme-independent report uses the familiar result-analysis workbook tabs")
ok(generic_wb.sheetnames[0] == "Result_Analysis" and generic_wb.active.title == "Result_Analysis",
   "the report opens directly on the student result analysis")
result_sheet = generic_wb["Result_Analysis"]
ok(result_sheet.freeze_panes == "D15" and generic_wb["Data"].freeze_panes == "C2",
   "student identifiers, headers, and source columns remain frozen while scrolling")
ok(result_sheet["A14"].value == "Rank" and result_sheet["B14"].value == "Enrollment",
   "result analysis has clear filterable student-table headers")
last_col = result_sheet.cell(14, result_sheet.max_column).column_letter
ok(result_sheet.auto_filter.ref == f"A14:{last_col}{result_sheet.max_row}",
   "the student filter spans all report rows and columns")
ok(all(result_sheet.column_dimensions[cell.column_letter].width <= 34 for cell in result_sheet[14]),
   "main report columns have bounded readable widths")
ok(result_sheet["B15"].value == "U1001" and result_sheet["B16"].value == "U1002",
   "generic report rankings default to SGPA rather than CGPA")
ok(result_sheet["B5"].value == 2 and result_sheet["B7"].value == 1 and result_sheet["B10"].value == 50.0,
   "class summary counts and pass percentage are calculated from the workbook results")
ok(result_sheet["D4"].value == "Top performers · SGPA" and result_sheet["E5"].value == "Student Alpha",
   "top performers use the displayed ranking metric")
ok("Maximum marks (from source)" in [cell.value for cell in generic_wb["subjectwise"][13]],
   "subject-wise analysis remains available as a structured cross-check")
ok(generic_wb["Data"].max_row == len(flat.wb.marks) + 1
   and generic_wb["Data"]["A2"].value == "U1001" and generic_wb["Data"]["B2"].value == "Student Alpha",
   "every parsed result is retained with student identity in the auditable source tab")

ai5_report = io.BytesIO()
build_generic_report(ai5, "V", ai5_report)
ai5_output = load_workbook(ai5_report, data_only=True)
ai5_result = ai5_output["Result_Analysis"]
ok(ai5_result["B5"].value == 150 and ai5_result["B7"].value == 129
   and ai5_result["B8"].value == 21 and ai5_result["B10"].value == 86.0,
   "staff-example cohort totals are reproduced from parsed source data without hard-coded rows")
headers = {str(cell.value).splitlines()[1]: cell.column for cell in ai5_result[14]
           if isinstance(cell.value, str) and len(cell.value.splitlines()) >= 3}
enrollment_col = next(cell.column for cell in ai5_result[14] if cell.value == "Enrollment")
student_rows = {str(ai5_result.cell(row, enrollment_col).value): row
                for row in range(15, ai5_result.max_row + 1)}
for mark in ai5.wb.marks.itertuples():
    output_row = student_rows[str(mark.enroll)]
    output_col = headers[str(mark.code)]
    output_value = ai5_result.cell(output_row, output_col).value
    ok(output_value is None and pd.isna(mark.total) or output_value == mark.total,
       f"student/course output preserves source mark for {mark.enroll}/{mark.code}")
ai5_data = ai5_output["Data"]
ok(ai5_data.max_row == len(ai5.wb.marks) + 1,
   "auditable Data tab includes exactly one row for each parsed student-course result")
formula_book = load_workbook(ai5_report, data_only=False)
ok(not any(cell.data_type == "f" for row in formula_book["Result_Analysis"].iter_rows() for cell in row),
   "analysis values are materialized rather than dependent on Excel formula recalculation")

print("Excel compatibility tests: ERP variants, generic long-table input, and scheme-independent report passed")
