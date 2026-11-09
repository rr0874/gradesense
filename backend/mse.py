"""MSE-specific classification for ERP workbooks with explicit MSE columns."""
from dataclasses import dataclass

import pandas as pd

from .ingest import UnsupportedLayout

MSE_PASS_MARK = 12
MSE_MAX_MARKS = 30


@dataclass
class MSEAnalysis:
    term: str
    subject_summary: pd.DataFrame
    student_results: pd.DataFrame
    totals: dict


def validate_mse_dataset(ds):
    wb = ds.wb
    if "has_mse" not in wb.subjects or not wb.subjects.has_mse.fillna(False).any():
        raise UnsupportedLayout(
            "No explicit MSE component columns were found. MSE analysis currently supports the "
            "multi-level ERP subject-block layout with an MSE column."
        )
    if "mse_status" not in wb.marks:
        raise UnsupportedLayout("The workbook's MSE values could not be validated.")

    subjects = wb.subjects[wb.subjects.has_mse.fillna(False)][["term", "code", "label"]]
    records = wb.marks.merge(subjects, on=["term", "code"], how="inner")
    for row in records.itertuples():
        status = str(row.mse_status).strip()
        if status.startswith("Invalid:"):
            raise UnsupportedLayout(
                f"Unrecognized MSE mark {status.removeprefix('Invalid:').strip()!r} "
                f"for {row.enroll}, {row.label}."
            )
        if status == "Marked":
            if pd.isna(row.mse):
                raise UnsupportedLayout(f"An MSE mark is missing for {row.enroll}, {row.label}.")
            if row.mse < 0 or row.mse > MSE_MAX_MARKS:
                raise UnsupportedLayout(
                    f"MSE mark {row.mse:g} for {row.enroll}, {row.label} is outside the valid range 0–30."
                )
        elif status not in {"AB", "Unreported"}:
            raise UnsupportedLayout(f"Could not determine the MSE outcome for {row.enroll}, {row.label}.")


def analyze_mse(ds, term):
    validate_mse_dataset(ds)
    wb = ds.wb
    subjects = wb.subjects[
        (wb.subjects.term == term) & wb.subjects.has_mse.fillna(False)
    ].copy()
    if subjects.empty:
        raise UnsupportedLayout(f"Term {term} has no MSE component columns.")

    roster = wb.terms.loc[wb.terms.term == term, ["enroll"]].drop_duplicates()
    students = wb.students.merge(roster, on="enroll", how="inner")[["enroll", "name"]]
    if students.empty:
        raise UnsupportedLayout(f"Term {term} has no student records for MSE analysis.")
    marks = wb.marks[wb.marks.term == term].merge(
        subjects[["term", "code", "label"]], on=["term", "code"], how="inner"
    ).merge(students, on="enroll", how="left")
    if marks.name.isna().any():
        raise UnsupportedLayout(f"Term {term} has MSE marks for students missing from its roster.")
    if marks.duplicated(["enroll", "code"]).any():
        raise UnsupportedLayout(f"Term {term} contains duplicate MSE results for a student and subject.")

    marks["Outcome"] = "Unreported"
    marks.loc[marks.mse_status == "AB", "Outcome"] = "Absent"
    marks.loc[(marks.mse_status == "Marked") & (marks.mse < MSE_PASS_MARK), "Outcome"] = "Fail"
    marks.loc[(marks.mse_status == "Marked") & (marks.mse >= MSE_PASS_MARK), "Outcome"] = "Pass"

    summary = []
    expected = len(students)
    for subject in subjects.itertuples():
        rows = marks[marks.code == subject.code]
        counts = rows.Outcome.value_counts()
        absent_rows = int(counts.get("Absent", 0))
        passed = int(counts.get("Pass", 0))
        failed = int(counts.get("Fail", 0))
        unreported = max(0, expected - len(rows)) + int(counts.get("Unreported", 0))
        summary.append({
            "Subject": subject.label,
            "Course code": subject.code,
            "Students": expected,
            "Passed": passed,
            "Failed": failed,
            "Absent": absent_rows,
            "Unreported": unreported,
            "Pass rate (%)": round(100 * passed / (passed + failed), 2) if passed + failed else None,
        })
    subject_summary = pd.DataFrame(summary)

    student_outcomes = []
    by_student = {enroll: group for enroll, group in marks.groupby("enroll")}
    for student in students.itertuples(index=False):
        rows = by_student.get(student.enroll, marks.iloc[0:0])
        outcomes = set(rows.Outcome)
        unreported = max(0, len(subjects) - len(rows)) + int(rows.Outcome.eq("Unreported").sum())
        if "Fail" in outcomes:
            outcome = "Fail"
        elif "Absent" in outcomes:
            outcome = "Absent"
        elif unreported:
            outcome = "Unreported"
        else:
            outcome = "Pass"
        student_outcomes.append({
            "Enrollment": student.enroll,
            "Student": student.name,
            "MSE subjects": len(subjects),
            "Marked / reported": int(rows.Outcome.isin(["Pass", "Fail"]).sum()) + int(rows.Outcome.eq("Absent").sum()),
            "Pass": int(rows.Outcome.eq("Pass").sum()),
            "Fail": int(rows.Outcome.eq("Fail").sum()),
            "Absent": int(rows.Outcome.eq("Absent").sum()),
            "Unreported": unreported,
            "Overall MSE result": outcome,
        })
    student_results = pd.DataFrame(student_outcomes)
    totals = {
        "students": len(student_results),
        "pass": int(student_results["Overall MSE result"].eq("Pass").sum()),
        "fail": int(student_results["Overall MSE result"].eq("Fail").sum()),
        "absent": int(student_results["Overall MSE result"].eq("Absent").sum()),
        "absent_any": int(student_results["Absent"].gt(0).sum()),
        "unreported": int(student_results["Overall MSE result"].eq("Unreported").sum()),
    }
    return MSEAnalysis(term, subject_summary, student_results, totals)
