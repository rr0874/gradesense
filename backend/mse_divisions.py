"""Strict parser and summary model for multi-division MSE marksheet workbooks."""
import io
import math
import re
from dataclasses import dataclass

import pandas as pd
from openpyxl import load_workbook

from .ingest import UnsupportedLayout

PASS_MARK = 12
MAX_MARKS = 30


@dataclass
class DivisionMSE:
    records: pd.DataFrame
    academic_year: str

    @property
    def divisions(self):
        return self.records[["Class", "Division"]].drop_duplicates().reset_index(drop=True)

    @property
    def students(self):
        return self.records[["Class", "Division", "Roll No", "Student Name"]].drop_duplicates()

    @property
    def subjects(self):
        return self.records[["Class", "Division", "Course"]].drop_duplicates()

    @property
    def summary(self):
        rows = []
        keys = ["Class", "Division", "Course"]
        for (class_name, division, course), group in self.records.groupby(keys, sort=False, dropna=False):
            outcomes = group["Outcome"]
            rows.append({
                "Class": class_name,
                "Division": division,
                "Course": course,
                "Students Appeared": int(outcomes.isin(["Pass", "Fail"]).sum()),
                "Student Absent": int(outcomes.eq("Absent").sum()),
                "Students Passed": int(outcomes.eq("Pass").sum()),
                "Students Failed": int(outcomes.eq("Fail").sum()),
                "Unreported": int(outcomes.eq("Unreported").sum()),
                "Enrolled": int(len(group)),
            })
        return pd.DataFrame(rows)


def _text(value):
    if value is None:
        return ""
    if isinstance(value, float) and math.isfinite(value) and value.is_integer():
        value = int(value)
    return re.sub(r"\s+", " ", str(value)).strip()


def _key(value):
    return re.sub(r"[^a-z0-9]+", " ", _text(value).casefold()).strip()


def _student_id(value):
    return _text(value).upper()


def _mark(value, sheet, row, course):
    if value is None or not _text(value) or _text(value).casefold() in {"-", "—", "–"}:
        return None, "Absent"
    if isinstance(value, bool):
        raise UnsupportedLayout(f"{sheet}, row {row}, {course}: boolean cell cannot be an MSE mark.")
    if isinstance(value, (int, float)):
        mark = float(value)
        if not math.isfinite(mark) or mark < 0 or mark > MAX_MARKS:
            raise UnsupportedLayout(
                f"{sheet}, row {row}, {course}: MSE mark {_text(value)!r} is outside 0–{MAX_MARKS}."
            )
        return mark, "Pass" if mark >= PASS_MARK else "Fail"
    marker = re.sub(r"[\s.]+", "", _text(value)).casefold()
    if marker in {"ab", "abs", "absent", "a"}:
        return "AB", "Absent"
    raise UnsupportedLayout(
        f"{sheet}, row {row}, {course}: unrecognized mark {_text(value)!r}; "
        "expected a number from 0 to 30, an absence marker, or a blank."
    )


def _class_division(ws):
    heading = " ".join(_text(ws.cell(row, col).value) for row in range(1, 8)
                       for col in range(1, min(ws.max_column or 1, 8) + 1))
    class_match = re.search(r"\b(first|second|third|fourth|final)\s+year\b", heading, re.I)
    if not class_match:
        raise UnsupportedLayout(f"{ws.title}: could not identify the class/year from the sheet title.")
    class_name = f"{class_match.group(1).title()} Year"
    division_match = re.search(r"\bdivision\s*[-:]?\s*([A-Z])\b", heading, re.I)
    if not division_match:
        division_match = re.search(r"(?:^|[-\s])([A-Z])\s*$", ws.title.strip(), re.I)
    if not division_match:
        raise UnsupportedLayout(f"{ws.title}: could not identify the division letter.")
    return class_name, division_match.group(1).upper()


def _sheet_records(ws):
    identity = None
    for row in range(1, min(ws.max_row or 0, 30) + 1):
        columns = {}
        for col in range(1, (ws.max_column or 0) + 1):
            key = _key(ws.cell(row, col).value)
            if key in {"roll no", "roll number", "student roll no"}:
                columns.setdefault("roll", col)
            elif key in {"name", "student name"}:
                columns.setdefault("name", col)
        if {"roll", "name"} <= columns.keys():
            identity = (row, columns)
            break
    if identity is None:
        raise UnsupportedLayout(f"{ws.title}: no row with both Roll No and Name headers was found.")
    header_row, identity_cols = identity

    marks_row = None
    marks_start = None
    total_col = None
    for row in range(max(1, header_row - 2), min(ws.max_row or 0, header_row + 3) + 1):
        for col in range(1, (ws.max_column or 0) + 1):
            key = _key(ws.cell(row, col).value)
            if key == "marks" or re.fullmatch(r"marks 30", key):
                marks_row, marks_start = row, col
                break
        if marks_row is not None:
            for col in range(marks_start + 1, (ws.max_column or 0) + 1):
                if _key(ws.cell(marks_row, col).value) in {"total", "grand total"}:
                    total_col = col
                    break
            break
    if marks_row is None:
        raise UnsupportedLayout(f"{ws.title}: no Marks heading was found beside the student columns.")

    scan_start = min(header_row, marks_row) + 1
    scan_end = min(ws.max_row or 0, max(header_row, marks_row) + 4)
    subject_row, subject_cols, subjects = None, [], []
    for row in range(scan_start, scan_end + 1):
        labels = []
        for col in range(marks_start, (total_col or (ws.max_column or 0) + 1)):
            label = _text(ws.cell(row, col).value)
            if not label or _key(label) in {"marks", "marks 30", "total", "percent marks"}:
                continue
            labels.append((col, label))
        if labels:
            subject_row, subject_cols, subjects = row, [c for c, _ in labels], [v for _, v in labels]
            break
    if not subject_cols:
        raise UnsupportedLayout(f"{ws.title}: subject names were not found under the Marks heading.")
    normalized = [_key(subject) for subject in subjects]
    if len(set(normalized)) != len(normalized):
        raise UnsupportedLayout(f"{ws.title}: subject headings are duplicated; source columns cannot be distinguished safely.")

    class_name, division = _class_division(ws)
    rows = []
    first_data_row = max(header_row, marks_row, subject_row) + 1
    for row in range(first_data_row, (ws.max_row or 0) + 1):
        roll = _student_id(ws.cell(row, identity_cols["roll"]).value)
        name = _text(ws.cell(row, identity_cols["name"]).value)
        source_marks = [ws.cell(row, col).value for col in subject_cols]
        has_result = any(value is not None and _text(value) for value in source_marks)
        if not roll and not name:
            if has_result:
                raise UnsupportedLayout(f"{ws.title}, row {row}: marks exist without a roll number and student name.")
            continue
        if not roll or not name:
            raise UnsupportedLayout(f"{ws.title}, row {row}: student roll number or name is missing; refusing to map marks.")
        for course, col, source_value in zip(subjects, subject_cols, source_marks):
            mark, outcome = _mark(source_value, ws.title, row, course)
            rows.append({
                "Class": class_name,
                "Division": division,
                "Roll No": roll,
                "Student Name": name,
                "Course": course,
                "Source Sheet": ws.title,
                "Source Row": row,
                "Source Column": col,
                "Source Mark": source_value,
                "MSE Marks": mark,
                "Outcome": outcome,
            })
    if not rows:
        raise UnsupportedLayout(f"{ws.title}: no student marks rows were found.")
    frame = pd.DataFrame(rows)
    duplicate = frame.duplicated(["Roll No", "Course"], keep=False)
    if duplicate.any():
        record = frame.loc[duplicate].iloc[0]
        raise UnsupportedLayout(
            f"{ws.title}: roll number {record['Roll No']} appears more than once for {record['Course']}."
        )
    return frame


def read_division_mse_workbook(source, filename=""):
    """Parse one sheet per class/division and retain source-cell coordinates for audit."""
    if isinstance(source, (bytes, bytearray)):
        source = io.BytesIO(source)
    try:
        workbook = load_workbook(source, data_only=True, read_only=False)
    except Exception as ex:
        raise UnsupportedLayout(f"Could not open the division MSE workbook ({ex.__class__.__name__}).") from ex
    if not workbook.worksheets:
        raise UnsupportedLayout("The workbook has no worksheets.")
    frames = []
    for worksheet in workbook.worksheets:
        frames.append(_sheet_records(worksheet))
    records = pd.concat(frames, ignore_index=True)
    division_keys = records[["Class", "Division", "Roll No"]].drop_duplicates()
    if division_keys.duplicated(["Class", "Division", "Roll No"]).any():
        raise UnsupportedLayout("A roll number is duplicated across worksheets for the same class and division.")
    name_by_student = records.groupby(["Class", "Division", "Roll No"])["Student Name"].nunique()
    if name_by_student.gt(1).any():
        raise UnsupportedLayout("A roll number is associated with multiple names in the same class/division.")

    year_match = re.search(r"\b(20\d{2}\s*-\s*\d{2,4})\b",
                           filename or getattr(source, "name", "") or "", re.I)
    academic_year = re.sub(r"\s+", "", year_match.group(1)) if year_match else "Not stated in workbook"
    return DivisionMSE(records, academic_year)
