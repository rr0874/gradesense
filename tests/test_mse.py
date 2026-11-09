"""Focused tests for MSE parsing, outcomes, and whole-semester compatibility."""
import io
from openpyxl import Workbook
from openpyxl import load_workbook

from backend import Dataset, UnsupportedLayout
from backend.engine import Engine
from backend.mse import analyze_mse, validate_mse_dataset
from backend.report import build_mse_report
from backend.mse_divisions import read_division_mse_workbook
from backend.report import build_mse_analysis_ii


def mse_workbook(values):
    book = Workbook()
    sheet = book.active
    sheet.append(["Sr No", "Student Unique No", "Student Name", "Term - I",
                  None, None, None, None, None, None, None, None, None, None,
                  "Total", "Credit Reg", "Grade Total", "SGPA"])
    sheet.append([None, None, None, "Mathematics ( 231ABC123456 )"])
    sheet.append([None, None, None, "Int Passing( 0 )"])
    sheet.append([None, None, None, "Ext Passing( 0 )"])
    sheet.append([None, None, None, "MSE", "Tot", "P/F", "ESE", "P/F", "Tot.", "Gr.",
                  "Cr. Pt.", "Cr. Er.", "Pt.", "St."])
    for index, (mark, expected_grade) in enumerate(values, start=1):
        grade = expected_grade if expected_grade is not None else (
            "AB" if str(mark).strip().casefold() in {"ab", "absent"} else "P"
        )
        sheet.append([index, f"EN240000{index:02d}", f"Student {index}", mark, 0, "P", 0, "P",
                      0, grade, 3, 3, 4, "Pass", 0, 3, 0, 5.0])
    output = io.BytesIO()
    book.save(output)
    return output.getvalue()


def division_workbook():
    book = Workbook()
    first = book.active
    first.title = "SY- A"
    for sheet, title, subjects, rows in [
        (first, "Second Year Division - A Marksheet", ["FAI", "IPRM", "NO MARKS", "DASH"], [
            ["A01", "Student One", 11, "Abs", None, "-"],
            ["A02", "Student Two", 12, None, None, "-"],
        ]),
        (book.create_sheet("SY- B"), "Second Year Division - B Marksheet", ["DMGT", "DS"], [
            ["B01", "Student Three", 18, 0],
        ]),
    ]:
        sheet.append(["College"])
        sheet.append(["Department"])
        sheet.append([title])
        sheet.append([None, "Roll No", "Name", "Marks", *([None] * (len(subjects) - 1)), "Total"])
        sheet.append([None, None, None, *subjects])
        for roll, name, *marks in rows:
            sheet.append([None, roll, name, *marks, None])
    output = io.BytesIO()
    book.save(output)
    return output.getvalue()


parsed = Dataset.from_workbook(mse_workbook([
    (11, "P"), (12, "P"), ("AB", "AB"), ("Absent", ""), ("-", "P"), (30, "P"),
]), "mse-fixture", include_mse_rows=True)
validate_mse_dataset(parsed)
analysis = analyze_mse(parsed, "I")
subject = analysis.subject_summary.iloc[0]
assert (subject["Passed"], subject["Failed"], subject["Absent"], subject["Unreported"]) == (2, 1, 2, 1)
outcomes = analysis.student_results.set_index("Student")["Overall MSE result"]
assert outcomes["Student 1"] == "Fail", "11/30 is below the pass threshold"
assert outcomes["Student 2"] == "Pass", "12/30 meets the pass threshold"
assert outcomes["Student 3"] == outcomes["Student 4"] == "Absent", "AB and Absent are distinct absence markers"
assert outcomes["Student 5"] == "Unreported" and outcomes["Student 6"] == "Pass"
assert analysis.totals == {
    "students": 6, "pass": 2, "fail": 1, "absent": 2, "absent_any": 2, "unreported": 1
}

# A wide MSE register with the common three-row subject/mark header is mapped by
# the explicit source column positions, not by sorting or fuzzy name matching.
wide_book = Workbook()
wide_sheet = wide_book.active
wide_sheet.append(["Student Marks Register"])
wide_sheet.append([None, None, None, None, None, "DS-C", "IPR"])
wide_sheet.append([None, None, None, None, None, "Int Passing (20)", "Int Passing (20)"])
wide_sheet.append([None, None, None, None, None, "Ext Passing (20)", "Ext Passing (0)"])
wide_sheet.append(["Sr No", "Student Unique No", "Student Seat No", "Student Name",
                   "Student Mother Name", "Marks (30)", "Marks (30)"])
wide_sheet.append([1, "UID001", "T001", "Student One", "Parent One", 11, "Abs"])
wide_sheet.append([2, "UID002", "T002", "Student Two", "Parent Two", 12, 10])
wide_buffer = io.BytesIO()
wide_book.save(wide_buffer)
wide_ds = Dataset.from_workbook(wide_buffer.getvalue(), "wide-mse-fixture", include_mse_rows=True)
wide_analysis = analyze_mse(wide_ds, "MSE")
wide_marks = wide_ds.wb.marks.set_index(["enroll", "code"])
assert wide_marks.at[("UID001", "MSE-C6"), "mse"] == 11
assert wide_marks.at[("UID001", "MSE-C7"), "mse_status"] == "AB"
assert wide_analysis.subject_summary[["Subject", "Failed", "Passed", "Absent"]].values.tolist() == [
    ["DS-C", 1, 1, 0], ["IPR", 1, 0, 1]
]
assert wide_analysis.student_results["Overall MSE result"].tolist() == ["Fail", "Fail"]
wide_report = io.BytesIO()
build_mse_report(wide_ds, wide_analysis, wide_report)
wide_report.seek(0)
excel = load_workbook(wide_report, data_only=True)
assert excel.sheetnames == ["Result_Analysis", "subjectwise", "Data", "Notes"]
assert excel["Result_Analysis"]["B15"].value == "UID001"
assert excel["Result_Analysis"]["C15"].value == "Student One"
assert excel["Result_Analysis"]["D15"].value == 11
assert excel["Result_Analysis"]["E15"].value == "Fail"
assert excel["Result_Analysis"]["F15"].value == "AB"
assert excel["Result_Analysis"]["G15"].value == "Absent"
assert excel["Result_Analysis"]["N15"].value == "Fail"
assert excel["subjectwise"]["A6"].value == "DS-C"
assert excel["Data"].max_row == 1 + len(wide_ds.wb.marks)
assert excel["Data"]["A2"].value == "UID001" and excel["Data"]["E2"].value == 11
assert excel["Data"]["F3"].value == "AB" and excel["Data"]["G3"].value == "Absent"

for extra_row, expected_error in (
    ([3, "UID001", "T003", "Duplicate Student", "Parent", 20, 21], "occurs more than once"),
    ([3, None, "T003", "Missing ID", "Parent", 20, 21], "incomplete student ID/name"),
):
    malformed = Workbook()
    malformed_sheet = malformed.active
    malformed_sheet.append(["Student Marks Register"])
    malformed_sheet.append([None, None, None, None, None, "DS-C", "IPR"])
    malformed_sheet.append([None, None, None, None, None, "Int Passing (20)", "Int Passing (20)"])
    malformed_sheet.append([None, None, None, None, None, "Ext Passing (20)", "Ext Passing (0)"])
    malformed_sheet.append(["Sr No", "Student Unique No", "Student Seat No", "Student Name",
                            "Student Mother Name", "Marks (30)", "Marks (30)"])
    malformed_sheet.append([1, "UID001", "T001", "Student One", "Parent One", 11, "Abs"])
    malformed_sheet.append(extra_row)
    malformed_buffer = io.BytesIO()
    malformed.save(malformed_buffer)
    try:
        Dataset.from_workbook(malformed_buffer.getvalue(), "misaligned-wide-mse", include_mse_rows=True)
    except UnsupportedLayout as ex:
        assert expected_error in str(ex)
    else:
        raise AssertionError("Ambiguous student/mark rows must be rejected, not silently combined.")

# A normalized MSE table can be read without inventing an overall grade/status.
flat_book = Workbook()
flat_sheet = flat_book.active
flat_sheet.append(["Student ID", "Student Name", "Subject", "MSE Marks"])
flat_sheet.append(["S001", "Student One", "Mathematics", 11])
flat_sheet.append(["S002", "Student Two", "Mathematics", 12])
flat_sheet.append(["S003", "Student Three", "Mathematics", "Absent"])
flat_bytes = io.BytesIO()
flat_book.save(flat_bytes)
flat_ds = Dataset.from_workbook(flat_bytes.getvalue(), "flat-mse", include_mse_rows=True)
validate_mse_dataset(flat_ds)
flat_analysis = analyze_mse(flat_ds, "I")
assert flat_analysis.subject_summary[["Failed", "Passed", "Absent"]].iloc[0].tolist() == [1, 1, 1]
assert flat_analysis.student_results["Overall MSE result"].tolist() == ["Fail", "Pass", "Absent"]

division_bytes = division_workbook()
division_data = read_division_mse_workbook(
    division_bytes, "MSE MARKSHEET 2026-27 (SEM-I).xlsx"
)
assert division_data.academic_year == "2026-27"
assert len(division_data.students) == 3 and len(division_data.records) == 10
division_summary = division_data.summary
assert division_summary[["Class", "Division", "Course", "Students Appeared", "Student Absent",
                         "Students Passed", "Students Failed", "Unreported"]].values.tolist() == [
    ["Second Year", "A", "FAI", 2, 0, 1, 1, 0],
    ["Second Year", "A", "IPRM", 0, 2, 0, 0, 0],
    ["Second Year", "A", "NO MARKS", 0, 2, 0, 0, 0],
    ["Second Year", "A", "DASH", 0, 2, 0, 0, 0],
    ["Second Year", "B", "DMGT", 1, 0, 1, 0, 0],
    ["Second Year", "B", "DS", 1, 0, 0, 1, 0],
]
division_report = io.BytesIO()
build_mse_analysis_ii(division_data, division_report, division_bytes)
division_report.seek(0)
division_excel = load_workbook(division_report, data_only=False)
assert division_excel.sheetnames == ["ISE MSE Analysis", "Checks", "SY- A", "SY- B"]
summary_sheet = division_excel["ISE MSE Analysis"]
assert summary_sheet["A1"].value == "MSE Performance Analysis (Odd Semester)"
assert summary_sheet["A2"].value == "Academic Year : 2026-27"
assert summary_sheet["A5"].value == "Second Year" and summary_sheet["C5"].value == "FAI"
assert summary_sheet["D5"].value == "=COUNT('SY- A'!$D$6:$D$500)"
assert summary_sheet["E5"].value == "=COUNTA('SY- A'!$C$6:$C$500)-D5"
assert summary_sheet["F5"].value == '=COUNTIF(\'SY- A\'!$D$6:$D$500,">="&12)'
assert summary_sheet["G5"].value == '=COUNTIF(\'SY- A\'!$D$6:$D$500,"<"&12)'
assert summary_sheet["D6"].value == "=COUNT('SY- A'!$E$6:$E$500)"
assert summary_sheet["E6"].value == "=COUNTA('SY- A'!$C$6:$C$500)-D6"
assert summary_sheet["D7"].value == "-" and summary_sheet["G7"].value == "-"
assert "A5:A10" in {str(item) for item in summary_sheet.merged_cells.ranges}
assert summary_sheet["A1"].fill.fgColor.rgb == "FF9DC3E6"
checks_sheet = division_excel["Checks"]
assert checks_sheet["H5"].value == '=IF(AND(D5=E5,F5=G5),"OK","MISMATCH")'
assert checks_sheet["D5"].value == "=COUNTA('SY- A'!$C$6:$C$500)"
assert checks_sheet["H7"].value == "OK"
assert checks_sheet["D7"].value == "=COUNTA('SY- A'!$C$6:$C$500)"
assert division_excel.calculation.calcMode == "auto"
division_report.seek(0)
cached_excel = load_workbook(division_report, data_only=True)
assert [cached_excel["ISE MSE Analysis"].cell(5, col).value for col in range(4, 8)] == [2, 0, 1, 1]
assert [cached_excel["ISE MSE Analysis"].cell(6, col).value for col in range(4, 8)] == [0, 2, 0, 0]
assert cached_excel["Checks"]["H5"].value == "OK"
assert [cached_excel["Checks"].cell(5, col).value for col in range(4, 8)] == [2, 2, 2, 2]
assert cached_excel["Checks"]["H7"].value == "OK"
source_sheet = division_excel["SY- A"]
assert source_sheet["B4"].value == "Roll No" and source_sheet["D6"].value == 11
assert source_sheet["B4"].style_id == load_workbook(
    io.BytesIO(division_bytes), data_only=False
)["SY- A"]["B4"].style_id
assert division_data.records.loc[
    division_data.records["Course"].eq("IPRM"), "Outcome"
].tolist() == ["Absent", "Absent"]
assert division_data.records.loc[
    division_data.records["Course"].eq("DASH"), "Outcome"
].tolist() == ["Absent", "Absent"]

# New students added beneath the original roster stay inside the formula ranges.
extended_source = load_workbook(io.BytesIO(division_bytes))
extended_sheet = extended_source["SY- A"]
for roll, name, mark in [
    ("A03", "Dummy Student Three", 12),
    ("A04", "Dummy Student Four", 11),
    ("A05", "Dummy Student Five", 18),
]:
    extended_sheet.append([None, roll, name, mark, None, None, None, None])
extended_bytes = io.BytesIO()
extended_source.save(extended_bytes)
extended_analysis = read_division_mse_workbook(
    extended_bytes.getvalue(), "MSE MARKSHEET 2026-27 (SEM-I).xlsx"
)
extended_report = io.BytesIO()
build_mse_analysis_ii(extended_analysis, extended_report, extended_bytes.getvalue())
extended_report.seek(0)
extended_cached = load_workbook(extended_report, data_only=True)
extended_summary = extended_cached["ISE MSE Analysis"]
extended_checks = extended_cached["Checks"]
assert extended_summary["D5"].value + extended_summary["E5"].value == 5
assert extended_checks["D5"].value == 5
assert extended_checks["E5"].value == 5
assert extended_checks["H5"].value == "OK"

invalid = Dataset.from_workbook(mse_workbook([(31, "P")]), "invalid-mse-fixture")
try:
    validate_mse_dataset(invalid)
except UnsupportedLayout as ex:
    assert "outside the valid range 0–30" in str(ex)
else:
    raise AssertionError("MSE marks above 30 must be rejected")

# The existing whole-semester path still ingests and analyses the sample unchanged.
whole_semester = Dataset.from_workbook("data/AIML_4.xlsx", "whole-semester-regression")
assert whole_semester.wb.info["students"] == 153
assert whole_semester.wb.info["terms"] == ["I", "II", "III", "IV"]
assert len(whole_semester.wb.marks) == 7275
summary = Engine(whole_semester, catalogs={}, log_path="NUL").ask("class summary").text
assert "Pass 133" in summary
assert whole_semester.wb.marks.status.value_counts().to_dict()["Pass"] > 0

print("MSE tests: parsing, 11/12 threshold, AB/Absent, range validation, and semester regression passed")
