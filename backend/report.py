"""Stage 5 - staff-style result analysis workbook (Result_Analysis + subjectwise) from an ingested sheet and a scheme JSON.
Marks live once on a `Data` sheet; every total, class and count in the other sheets is a formula over it.
    python -m backend.report data/AI_5.xlsx backend/report_schemes/semV.json out.xlsx"""
import io, json, posixpath, re, sys, zipfile
from xml.etree import ElementTree
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter as L
from .mse_divisions import DivisionMSE, PASS_MARK

FIRST_ROW = 6
MAX_ROW = 500

FONT = lambda **k: Font(name="Arial", size=10, **k)
BLUE, GREY = FONT(color="0000FF"), PatternFill("solid", fgColor="D9E1F2")
THIN = Border(*(Side(style="thin"),) * 4)
GROUPS = [("Theory", "Theory"), ("MDM", "MDM"), ("OE", "OE"), ("Lab", "Laboratory")]
CLASS_F = ('=IF({s}<>"Pass","Fail",IF({g}>=8,"Distinction",IF({g}>=7,"First Class",IF({g}>=6,"Second Class",IF({g}>=5,"Pass","Fail")))))')

def _put(ws, ref, v, font=None, fmt=None, bold=False, fill=None, align=None, border=True):
    c = ws[ref]; c.value = v; c.font = font or FONT(bold=bold)
    if fmt: c.number_format = fmt
    if fill: c.fill = fill
    if border: c.border = THIN
    c.alignment = Alignment(horizontal=align or "center", vertical="center", wrap_text=True)
    return c

def _merge(ws, rng, v, **k):
    ws.merge_cells(rng); _put(ws, rng.split(":")[0], v, **k)
    for row in ws[rng]:
        for c in row: c.border = THIN

def build_generic_report(ds, term, path):
    """Build a scheme-independent, value-based report from the validated parsed workbook."""
    wb = ds.wb
    if term not in set(wb.terms.term):
        raise ValueError(f"Term {term} is not present in the workbook.")

    subjects = wb.subjects[wb.subjects.term == term].copy().reset_index(drop=True)
    marks = wb.marks[wb.marks.term == term].copy()
    term_results = wb.terms[wb.terms.term == term].drop(columns="term_n", errors="ignore")
    active_enrollments = set(term_results.enroll) | set(marks.enroll)
    students = wb.students[wb.students.enroll.isin(active_enrollments)].copy().merge(
        term_results, on="enroll", how="left")
    if students.empty or subjects.empty or marks.empty:
        raise ValueError(f"Term {term} has no student or subject records to report.")
    if marks.duplicated(["enroll", "code"]).any():
        raise ValueError(f"Term {term} contains duplicate student-course rows; resolve these before exporting.")

    labels = subjects.set_index("code").label.to_dict()
    used = set()
    display = {}
    for row in subjects.itertuples():
        name = str(row.label).strip() or str(row.abbr).strip() or str(row.code)
        if name.casefold() in used:
            name = f"{name} [{row.code}]"
        used.add(name.casefold())
        display[row.code] = name
    marks["Subject"] = marks.code.map(display)
    if marks.Subject.isna().any():
        raise ValueError(f"Term {term} has result rows for courses not present in its subject list.")
    subject_keys = {row.code: f"__course_{index}" for index, row in enumerate(subjects.itertuples())}
    marks["__subject_key"] = marks.code.map(subject_keys)
    pivot = marks.pivot(index="enroll", columns="__subject_key", values="total")
    pivot = pivot.reindex(columns=list(subject_keys.values()))
    student_report = students.merge(pivot, left_on="enroll", right_index=True, how="left")

    if term == wb.terms.sort_values("term_n").iloc[-1].term and not wb.cumulative.empty:
        cumulative = wb.cumulative.copy()
        student_report = student_report.merge(cumulative, on="enroll", how="left", suffixes=("", "_cumulative"))
    elif "cgpa" not in student_report:
        student_report["cgpa"] = float("nan")

    latest = term == wb.terms.sort_values("term_n").iloc[-1].term
    if latest and "status" in student_report:
        outcomes = student_report.status.fillna("").astype(str).str.strip()
    else:
        def derive_outcome(values):
            values = values.dropna().astype(str).str.strip()
            if values.eq("Fail").any(): return "Fail"
            if values.eq("AB").any(): return "AB"
            if len(values) and values.eq("Pass").all(): return "Pass"
            return "Unreported"
        outcomes = marks.groupby("enroll").status.apply(derive_outcome)
        outcomes = student_report.enroll.map(outcomes).fillna("No result")
    student_report["Outcome"] = outcomes
    student_report["Result total"] = marks.groupby("enroll").total.sum(min_count=1).reindex(student_report.enroll).to_numpy()
    student_report["Backlog / non-pass subjects"] = student_report.enroll.map(
        marks[marks.status.isin(["Fail", "AB"])].groupby("enroll").Subject.apply(
            lambda values: ", ".join(values.astype(str))))

    score = "sgpa" if "sgpa" in student_report and student_report.sgpa.notna().any() else "cgpa"
    if score not in student_report:
        student_report[score] = float("nan")
    student_report = student_report.sort_values(
        score, ascending=False, na_position="last", kind="stable").reset_index(drop=True)
    student_report.insert(0, "Rank", range(1, len(student_report) + 1))

    subject_stats = []
    for row in subjects.itertuples():
        result = marks[marks.code == row.code]
        maximum = row.max_total
        maximum = float(maximum) if pd.notna(maximum) and maximum > 0 else None
        average = result.total.mean() if result.total.notna().any() else None
        subject_stats.append({
            "Subject": display[row.code],
            "Course code": row.code,
            "Students with results": int(len(result)),
            "Passed": int(result.status.eq("Pass").sum()),
            "Failed": int(result.status.eq("Fail").sum()),
            "Absent": int(result.status.eq("AB").sum()),
            "Pass rate (%)": round(100 * result.status.eq("Pass").mean(), 2) if len(result) else None,
            "Average marks": round(float(average), 2) if average is not None else None,
            "Maximum marks (from source)": maximum,
            "Average marks (%)": round(100 * average / maximum, 2) if average is not None and maximum else None,
            "Credits (from source)": float(row.credits) if pd.notna(row.credits) else None,
        })
    subject_frame = pd.DataFrame(subject_stats)

    book = Workbook()
    overview = book.active
    overview.title = "subjectwise"
    result_sheet = book.create_sheet("Student Results")
    result_sheet.title = "Result_Analysis"
    subjects_sheet = overview
    source_sheet = book.create_sheet("Data")
    notes_sheet = book.create_sheet("Notes")
    navy = "17365D"
    pale = "E8EFF7"
    gold = "B08A35"
    thin_grey = Side(style="thin", color="D8E0EA")
    header_fill = PatternFill("solid", fgColor=navy)
    sub_fill = PatternFill("solid", fgColor=pale)
    title_font = Font(name="Aptos Display", size=18, bold=True, color=navy)
    header_font = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
    body_font = Font(name="Aptos", size=10, color="243247")

    def format_table(ws, header_row, first_col=1, last_col=None, freeze=None):
        last_col = last_col or ws.max_column
        for cell in ws[header_row][first_col - 1:last_col]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(bottom=Side(style="medium", color=gold))
        ws.row_dimensions[header_row].height = 32
        for row in ws.iter_rows(min_row=header_row + 1, max_col=last_col):
            for cell in row:
                cell.font = body_font
                cell.alignment = Alignment(
                    horizontal="center" if isinstance(cell.value, (int, float)) else "left",
                    vertical="center", wrap_text=True)
                cell.border = Border(bottom=thin_grey)
                if cell.data_type == "n":
                    cell.number_format = "0.##"
        ws.freeze_panes = freeze or f"A{header_row + 1}"
        ws.auto_filter.ref = f"{L(first_col)}{header_row}:{L(last_col)}{max(header_row, ws.max_row)}"
        ws.sheet_view.showGridLines = False
        ws.sheet_view.zoomScale = 90
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.page_setup.orientation = "landscape"

    # The first tab summarizes per-subject results in the staff workbook's familiar layout.
    overview.sheet_view.showGridLines = False
    overview.sheet_view.zoomScale = 95
    overview["A1"] = "Subject-wise Result Analysis"
    overview["A1"].font = title_font
    overview.merge_cells("A1:D1")
    overview["A2"] = f"Term {term} · {ds.name} · {len(students):,} students"
    overview["A2"].font = Font(name="Aptos", size=11, color="64748B")
    overview.merge_cells("A2:D2")
    metric_name = "CGPA" if score == "cgpa" else "SGPA"
    overview["A4"] = "Measure"
    overview["B4"] = "Result"
    for cell in overview[4][:2]:
        cell.fill, cell.font = header_fill, header_font
    stats = [
        ("Students appeared", len(students)),
        ("Subjects in term", len(subjects)),
        ("Students recorded as passed", overall_pass := int(student_report.Outcome.eq("Pass").sum())),
        ("Students failed / ATKT", int(student_report.Outcome.isin(["Fail", "ATKT"]).sum())),
        ("Students absent", int(student_report.Outcome.eq("AB").sum())),
        ("Pass percentage (%)", round(100 * overall_pass / len(students), 2) if len(students) else None),
    ]
    for row_idx, (label, value) in enumerate(stats, 5):
        overview.cell(row_idx, 1, label)
        overview.cell(row_idx, 2, value)
        overview.cell(row_idx, 1).font = Font(name="Aptos", size=10, color="243247")
        overview.cell(row_idx, 2).font = Font(name="Aptos", size=10, bold=True, color=navy)
        for cell in overview[row_idx][:2]:
            cell.border = Border(bottom=thin_grey)
            cell.alignment = Alignment(vertical="center")
    overview["D4"] = f"Top performers by {metric_name}"
    overview.merge_cells("D4:F4")
    overview["D4"].fill, overview["D4"].font = header_fill, header_font
    top_performers = student_report.dropna(subset=[score]).head(3)
    for row_idx, (_, student) in enumerate(top_performers.iterrows(), 5):
        overview.cell(row_idx, 4, f"{row_idx - 4}{'st' if row_idx == 5 else 'nd' if row_idx == 6 else 'rd'}")
        overview.cell(row_idx, 5, student["name"])
        overview.cell(row_idx, 6, student[score])
        overview.cell(row_idx, 6).number_format = "0.00"
        for cell in overview[row_idx][3:6]:
            cell.border = Border(bottom=thin_grey)
            cell.alignment = Alignment(vertical="center")
    overview["D9"] = "Ranking metric"
    overview["E9"] = metric_name
    overview["D9"].font = Font(name="Aptos", size=9, color="64748B")
    overview["E9"].font = Font(name="Aptos", size=9, bold=True, color=navy)
    overview["A12"] = "Subject results"
    overview["A12"].font = Font(name="Aptos", size=13, bold=True, color=navy)
    for col, title in enumerate(subject_frame.columns, 1):
        overview.cell(13, col, title)
    for row_idx, values in enumerate(subject_frame.itertuples(index=False, name=None), 10):
        for col, value in enumerate(values, 1):
            overview.cell(row_idx + 4, col, value)
    format_table(overview, 13, last_col=len(subject_frame.columns), freeze="A14")
    for col, width in enumerate([30, 20, 16, 12, 12, 12, 15, 16, 22, 18, 17], 1):
        overview.column_dimensions[L(col)].width = width

    details = student_report[["Rank", "enroll", "name"]].rename(
        columns={"enroll": "Enrollment", "name": "Student name"}).copy()
    if "seat" in student_report:
        details.insert(3, "Seat number", student_report["seat"].to_numpy())
    subject_columns = list(pivot.columns)
    subject_display_keys = {key: display[code] for code, key in subject_keys.items()}
    for subject in subject_columns:
        details[subject] = student_report[subject].to_numpy()
    for source_column, display_name in (
        ("Result total", "Result total"), ("sgpa", "SGPA"), ("cgpa", "CGPA"),
        ("pct", "Percentage (%)"), ("grade", "Overall grade"), ("Outcome", "Overall status"),
        ("Backlog / non-pass subjects", "Non-pass / absent subjects"),
    ):
        if source_column in student_report:
            details[display_name] = student_report[source_column].to_numpy()

    # Familiar staff-style student table, with bounded widths and explicit panes.
    result_sheet.sheet_view.showGridLines = False
    result_sheet.sheet_view.zoomScale = 75
    last_column = len(details.columns)
    result_sheet["A1"] = f"RESULT ANALYSIS — TERM {term}"
    result_sheet["A1"].font = title_font
    result_sheet.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last_column)
    result_sheet["A2"] = f"{ds.name} · Ranked by {metric_name} · {len(students):,} students · {len(subjects)} subjects"
    result_sheet["A2"].font = Font(name="Aptos", size=10, color="64748B")
    result_sheet.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last_column)
    result_sheet["A4"] = "Class summary"
    result_sheet["A4"].font = Font(name="Aptos", size=12, bold=True, color=navy)
    for row_idx, (label, value) in enumerate(stats, 5):
        result_sheet.cell(row_idx, 1, label)
        result_sheet.cell(row_idx, 2, value)
        result_sheet.cell(row_idx, 1).font = Font(name="Aptos", size=9, color="243247")
        result_sheet.cell(row_idx, 2).font = Font(name="Aptos", size=9, bold=True, color=navy)
        for cell in result_sheet[row_idx][:2]:
            cell.border = Border(bottom=thin_grey)
    result_sheet["D4"] = f"Top performers · {metric_name}"
    result_sheet.merge_cells("D4:F4")
    result_sheet["D4"].fill, result_sheet["D4"].font = header_fill, header_font
    for row_idx, (_, student) in enumerate(top_performers.iterrows(), 5):
        result_sheet.cell(row_idx, 4, f"{row_idx - 4}")
        result_sheet.cell(row_idx, 5, student["name"])
        result_sheet.cell(row_idx, 6, student[score])
        result_sheet.cell(row_idx, 6).number_format = "0.00"
        for cell in result_sheet[row_idx][3:6]:
            cell.border = Border(bottom=thin_grey)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
    result_sheet["A12"] = "Student results"
    result_sheet["A12"].font = Font(name="Aptos", size=13, bold=True, color=navy)

    for col, name in enumerate(details.columns, 1):
        heading = str(name)
        if name in subject_columns:
            course_code = next(code for code, key in subject_keys.items() if key == name)
            source = subjects[subjects.code == course_code]
            maximum = source.iloc[0].max_total if len(source) else None
            max_text = f"{maximum:g}" if pd.notna(maximum) and maximum > 0 else "not stated"
            heading = f"{subject_display_keys[name]}\n{course_code}\nMax: {max_text}"
        result_sheet.cell(14, col, heading)
    for row_idx, values in enumerate(details.itertuples(index=False, name=None), 15):
        for col, value in enumerate(values, 1):
            result_sheet.cell(row_idx, col, None if pd.isna(value) else value)
    format_table(result_sheet, 14, last_col=len(details.columns), freeze="D15")
    result_sheet.row_dimensions[14].height = 58
    for col, name in enumerate(details.columns, 1):
        width = 14
        if name == "Enrollment": width = 19
        elif name == "Student name": width = 26
        elif name == "Seat number": width = 13
        elif name in subject_columns: width = 17
        elif name == "Non-pass / absent subjects": width = 34
        elif name == "Overall status": width = 17
        result_sheet.column_dimensions[L(col)].width = width
    result_sheet.auto_filter.ref = f"A14:{L(last_column)}{result_sheet.max_row}"
    result_sheet.freeze_panes = "E15" if "Seat number" in details else "D15"
    result_sheet.print_title_rows = "1:14"
    result_sheet.page_setup.fitToWidth = 1
    result_sheet.page_setup.fitToHeight = 0

    for col, width in enumerate([30, 20, 16, 12, 12, 12, 15, 16, 22, 18, 17], 1):
        overview.column_dimensions[L(col)].width = width

    identity = wb.students[["enroll", "name"]].drop_duplicates("enroll").set_index("enroll").name
    source = marks.copy()
    source.insert(0, "Student name", source.enroll.map(identity))
    source.insert(1, "Term", term)
    source_columns = ["enroll", "Student name", "Term", "code", "Subject", "total"]
    source_columns += [column for column in ("ise", "mse", "internal", "int", "ese", "poe", "other")
                       if column in source and source[column].notna().any()]
    source_columns += [column for column in ("grade", "status", "credits", "earned") if column in source]
    source = source[source_columns].rename(columns={
        "enroll": "Enrollment", "code": "Course code", "total": "Marks",
        "ise": "ISE", "mse": "MSE", "internal": "Internal", "int": "INT",
        "ese": "ESE", "poe": "POE", "other": "Other marks",
        "grade": "Grade", "status": "Status", "credits": "Credits", "earned": "Credits earned",
    })
    source_headers = list(source.columns)
    for col, title in enumerate(source_headers, 1):
        source_sheet.cell(1, col, title)
    for row_idx, values in enumerate(source.itertuples(index=False, name=None), 2):
        for col, value in enumerate(values, 1):
            source_sheet.cell(row_idx, col, None if pd.isna(value) else value)
    format_table(source_sheet, 1, last_col=len(source_headers), freeze="C2")
    for col, title in enumerate(source_headers, 1):
        source_sheet.column_dimensions[L(col)].width = min(30, max(13, len(str(title)) + 3))
    source_sheet.column_dimensions["A"].width = 20
    source_sheet.column_dimensions["B"].width = 26
    source_sheet.column_dimensions["E"].width = 30

    notes = [
        ("About this workbook", "Generated from the active uploaded result workbook after its data checks."),
        ("Ranking metric", f"Students are ordered by {metric_name}; SGPA and CGPA are included separately when the source provides them."),
        ("Source fields", "Subject marks, grades, statuses, credits, and identifiers are carried over from the parsed source sheet."),
        ("Unknown maximum marks", "A blank maximum or average percentage means the source workbook did not provide a reliable maximum; no maximum was guessed."),
        ("Outcome rules", "No institution-specific grade bands or pass thresholds are invented. Outcomes shown here come from the sheet or the subject-level statuses."),
        ("Navigation", "Result_Analysis freezes student identifiers and the header row. subjectwise and Data freeze their headers for scrolling."),
    ]
    notes_sheet.sheet_view.showGridLines = False
    notes_sheet["A1"] = "Report notes"
    notes_sheet["A1"].font = title_font
    for row_idx, (heading, text) in enumerate(notes, 3):
        notes_sheet.cell(row_idx, 1, heading).font = Font(name="Aptos", size=10, bold=True, color=navy)
        notes_sheet.cell(row_idx, 2, text).font = body_font
        notes_sheet.cell(row_idx, 2).alignment = Alignment(wrap_text=True, vertical="top")
        notes_sheet.row_dimensions[row_idx].height = 34
    notes_sheet.column_dimensions["A"].width = 25
    notes_sheet.column_dimensions["B"].width = 110
    notes_sheet.freeze_panes = "A3"
    notes_sheet.sheet_view.zoomScale = 95

    book._sheets = [result_sheet, overview, source_sheet, notes_sheet]
    book.active = 0
    book.calculation.fullCalcOnLoad = True
    book.calculation.forceFullCalc = True
    book.save(path)
    return path

def build_mse_report(ds, analysis, path):
    """Build the same four-tab, auditable report structure for MSE results."""
    wb = ds.wb
    subjects = wb.subjects[
        (wb.subjects.term == analysis.term) & wb.subjects.has_mse.fillna(False)
    ].copy()
    students = wb.students[wb.students.enroll.isin(
        wb.terms.loc[wb.terms.term == analysis.term, "enroll"]
    )].copy()
    subject_marks = wb.marks[wb.marks.term == analysis.term].merge(
        subjects[["term", "code", "label"]], on=["term", "code"], how="inner"
    )
    subject_marks = subject_marks.merge(students[["enroll", "name"]], on="enroll", how="left")
    if subjects.empty or students.empty or subject_marks.empty:
        raise ValueError(f"Term {analysis.term} has no MSE result rows to export.")
    if subject_marks.duplicated(["enroll", "code"]).any():
        raise ValueError("Duplicate student/subject MSE results prevent a reliable workbook export.")
    if subject_marks.name.isna().any():
        raise ValueError("Some MSE results do not map to a known student; workbook export was stopped.")

    subject_names = {}
    used_names = set()
    for row in subjects.itertuples():
        label = str(row.label).strip() or str(row.code)
        key = label.casefold()
        if key in used_names:
            label = f"{label} [{row.code}]"
        used_names.add(label.casefold())
        subject_names[row.code] = label

    subject_marks["Subject"] = subject_marks.code.map(subject_names)
    subject_marks["MSE outcome"] = "Unreported"
    subject_marks.loc[subject_marks.mse_status.eq("AB"), "MSE outcome"] = "Absent"
    subject_marks.loc[
        subject_marks.mse_status.eq("Marked") & subject_marks.mse.lt(12), "MSE outcome"
    ] = "Fail"
    subject_marks.loc[
        subject_marks.mse_status.eq("Marked") & subject_marks.mse.ge(12), "MSE outcome"
    ] = "Pass"

    book = Workbook()
    result = book.active
    result.title = "Result_Analysis"
    overview = book.create_sheet("subjectwise")
    source = book.create_sheet("Data")
    notes = book.create_sheet("Notes")
    navy, pale, gold = "17365D", "E8EFF7", "B08A35"
    header_fill = PatternFill("solid", fgColor=navy)
    pale_fill = PatternFill("solid", fgColor=pale)
    header_font = Font(name="Aptos", size=10, bold=True, color="FFFFFF")
    title_font = Font(name="Aptos Display", size=18, bold=True, color=navy)
    body_font = Font(name="Aptos", size=10, color="243247")
    grey_side = Side(style="thin", color="D8E0EA")

    def format_sheet(sheet, header_row, last_col, freeze):
        sheet.sheet_view.showGridLines = False
        sheet.sheet_view.zoomScale = 85
        sheet.freeze_panes = freeze
        sheet.auto_filter.ref = f"A{header_row}:{L(last_col)}{sheet.max_row}"
        sheet.sheet_properties.pageSetUpPr.fitToPage = True
        sheet.page_setup.orientation = "landscape"
        sheet.page_setup.fitToWidth = 1
        sheet.page_setup.fitToHeight = 0
        for cell in sheet[header_row][:last_col]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            cell.border = Border(bottom=Side(style="medium", color=gold))
        sheet.row_dimensions[header_row].height = 34
        for row in sheet.iter_rows(min_row=header_row + 1, max_col=last_col):
            for cell in row:
                cell.font = body_font
                cell.alignment = Alignment(vertical="center", wrap_text=True)
                cell.border = Border(bottom=grey_side)
                if isinstance(cell.value, (int, float)):
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    metric_rows = [
        ("Students", analysis.totals["students"]),
        ("Overall Pass", analysis.totals["pass"]),
        ("Overall Fail", analysis.totals["fail"]),
        ("Overall Absent", analysis.totals["absent"]),
        ("Absent in any subject", analysis.totals["absent_any"]),
        ("Overall Unreported", analysis.totals["unreported"]),
    ]
    subject_order = list(subjects.code)
    label_order = [subject_names[code] for code in subject_order]

    result["A1"] = f"MSE RESULT ANALYSIS — {analysis.term}"
    result["A1"].font = title_font
    result.merge_cells(start_row=1, start_column=1, end_row=1, end_column=3 + 2 * len(subject_order) + 7)
    result["A2"] = f"{ds.name} · 30-mark MSE · Pass threshold: 12"
    result["A2"].font = Font(name="Aptos", size=10, color="64748B")
    result.merge_cells(start_row=2, start_column=1, end_row=2, end_column=3 + 2 * len(subject_order) + 7)
    result["A4"] = "Overall summary"
    result["A4"].font = Font(name="Aptos", size=12, bold=True, color=navy)
    for col, header in enumerate(("Measure", "Count"), 1):
        cell = result.cell(5, col, header)
        cell.fill, cell.font = header_fill, header_font
    for row, (label, value) in enumerate(metric_rows, 6):
        result.cell(row, 1, label)
        result.cell(row, 2, value)

    detail_start = 14
    details = ["Rank", "Enrollment", "Student name"]
    subject_output_columns = {}
    for code, label in zip(subject_order, label_order):
        subject_output_columns[code] = (len(details) + 1, len(details) + 2)
        details.extend((f"{label} · Marks / 30", f"{label} · Outcome"))
    details.extend(["MSE subjects", "Reported", "Pass", "Fail", "Absent", "Unreported", "Overall MSE result"])
    for col, heading in enumerate(details, 1):
        result.cell(detail_start, col, heading)

    student_outcome = analysis.student_results.set_index("Enrollment")
    mark_lookup = subject_marks.set_index(["enroll", "code"])
    for rank, student in enumerate(students.itertuples(index=False), 1):
        row_number = detail_start + rank
        result.cell(row_number, 1, rank)
        result.cell(row_number, 2, student.enroll)
        result.cell(row_number, 3, student.name)
        for code in subject_order:
            mark_col, outcome_col = subject_output_columns[code]
            try:
                record = mark_lookup.loc[(student.enroll, code)]
            except KeyError:
                record = None
            if record is not None:
                score = record.mse if pd.notna(record.mse) else (
                    "AB" if record.mse_status == "AB" else None
                )
                result.cell(row_number, mark_col, score)
                result.cell(row_number, outcome_col, record["MSE outcome"])
        totals = student_outcome.loc[student.enroll]
        for offset, key in enumerate(
            ["MSE subjects", "Marked / reported", "Pass", "Fail", "Absent", "Unreported",
             "Overall MSE result"],
            len(details) - 6,
        ):
            result.cell(row_number, offset, totals[key])
    format_sheet(result, detail_start, len(details), f"D{detail_start + 1}")
    result.freeze_panes = f"D{detail_start + 1}"
    result.column_dimensions["A"].width = 9
    result.column_dimensions["B"].width = 20
    result.column_dimensions["C"].width = 26
    for code, (mark_col, outcome_col) in subject_output_columns.items():
        result.column_dimensions[L(mark_col)].width = 20
        result.column_dimensions[L(outcome_col)].width = 18
    for col in range(1, len(details) + 1):
        if not result.column_dimensions[L(col)].width:
            result.column_dimensions[L(col)].width = 16
    result.auto_filter.ref = f"A{detail_start}:{L(len(details))}{detail_start + len(students)}"
    result.print_title_rows = f"1:{detail_start}"

    overview["A1"] = "MSE Subject-wise Result Analysis"
    overview["A1"].font = title_font
    overview.merge_cells("A1:H1")
    overview["A2"] = f"{ds.name} · {len(students):,} students · Marks out of 30 · Pass mark 12"
    overview["A2"].font = Font(name="Aptos", size=10, color="64748B")
    overview.merge_cells("A2:H2")
    overview_headers = list(analysis.subject_summary.columns)
    for col, heading in enumerate(overview_headers, 1):
        overview.cell(5, col, heading)
    for row, values in enumerate(analysis.subject_summary.itertuples(index=False, name=None), 6):
        for col, value in enumerate(values, 1):
            overview.cell(row, col, None if pd.isna(value) else value)
    format_sheet(overview, 5, len(overview_headers), "A6")
    for col, heading in enumerate(overview_headers, 1):
        overview.column_dimensions[L(col)].width = 22 if heading in {"Subject", "Course code"} else 16

    source_headers = [
        "Enrollment", "Student name", "Subject", "Course code", "MSE marks / 30",
        "MSE status", "Outcome", "Pass threshold",
    ]
    for col, heading in enumerate(source_headers, 1):
        source.cell(1, col, heading)
    for row, (_, record) in enumerate(subject_marks.iterrows(), 2):
        values = [
            record.enroll, record.name, record.Subject, record.code,
            record.mse if pd.notna(record.mse) else ("AB" if record.mse_status == "AB" else None),
            record.mse_status, record["MSE outcome"], 12,
        ]
        for col, value in enumerate(values, 1):
            source.cell(row, col, value)
    format_sheet(source, 1, len(source_headers), "C2")
    for col, width in enumerate([20, 26, 24, 18, 19, 16, 16, 16], 1):
        source.column_dimensions[L(col)].width = width

    notes["A1"] = "Report notes"
    notes["A1"].font = title_font
    notes_rows = [
        ("Source", "Student identity, subject names, and MSE values are mapped directly from the uploaded workbook. Data has one row per source student and subject."),
        ("Maximum marks", "MSE marks are interpreted on a 30-mark scale, as specified by the Marks (30) source headers."),
        ("Pass rule", "Numeric MSE marks below 12 are Fail; marks from 12 through 30 are Pass."),
        ("Absence", "AB, Abs, and Absent are recorded as Absent, not as a numeric zero."),
        ("Missing values", "Blank or dash source marks remain unreported; they are not treated as zero."),
        ("Overall result", "A student is an overall Fail if any MSE subject fails; otherwise Absent if any subject is absent; Unreported if any result is missing; Pass only if all subjects pass."),
        ("Audit trail", "The Data sheet contains the source marks, outcomes, student IDs, and subject column mapping used for this analysis."),
    ]
    for row, (heading, text) in enumerate(notes_rows, 3):
        notes.cell(row, 1, heading).font = Font(name="Aptos", size=10, bold=True, color=navy)
        notes.cell(row, 2, text).font = body_font
        notes.cell(row, 2).alignment = Alignment(wrap_text=True, vertical="top")
        notes.row_dimensions[row].height = 36
    notes.column_dimensions["A"].width = 25
    notes.column_dimensions["B"].width = 110
    notes.freeze_panes = "A3"
    notes.sheet_view.showGridLines = False
    book.active = 0
    book.save(path)
    return path

def _save_workbook_with_formula_caches(book, path, cached_values):
    namespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    document_relationships = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    package_relationships = "http://schemas.openxmlformats.org/package/2006/relationships"
    ElementTree.register_namespace("", namespace)
    ElementTree.register_namespace("r", document_relationships)

    buffer = io.BytesIO()
    book.save(buffer)
    buffer.seek(0)
    with zipfile.ZipFile(buffer, "r") as source_zip:
        workbook_xml = ElementTree.fromstring(source_zip.read("xl/workbook.xml"))
        relationships_xml = ElementTree.fromstring(source_zip.read("xl/_rels/workbook.xml.rels"))
        targets = {
            relationship.attrib["Id"]: relationship.attrib["Target"]
            for relationship in relationships_xml.findall(f"{{{package_relationships}}}Relationship")
        }
        sheet_paths = {}
        for sheet in workbook_xml.find(f"{{{namespace}}}sheets"):
            target = targets[sheet.attrib[f"{{{document_relationships}}}id"]]
            sheet_paths[sheet.attrib["name"]] = (
                target.lstrip("/") if target.startswith("/")
                else posixpath.normpath(posixpath.join("xl", target))
            )
        path_sheets = {sheet_paths[name]: values for name, values in cached_values.items()}
        patched = io.BytesIO()
        with zipfile.ZipFile(patched, "w", zipfile.ZIP_DEFLATED) as output_zip:
            for item in source_zip.infolist():
                contents = source_zip.read(item.filename)
                if item.filename in path_sheets:
                    root = ElementTree.fromstring(contents)
                    cells = {
                        cell.attrib["r"]: cell
                        for cell in root.iter(f"{{{namespace}}}c")
                        if "r" in cell.attrib
                    }
                    for coordinate, value in path_sheets[item.filename].items():
                        cell = cells.get(coordinate)
                        if cell is None or cell.find(f"{{{namespace}}}f") is None:
                            raise ValueError(f"Could not cache the MSE summary formula at {coordinate}.")
                        cached = cell.find(f"{{{namespace}}}v")
                        if cached is None:
                            cached = ElementTree.SubElement(cell, f"{{{namespace}}}v")
                        if isinstance(value, str):
                            cell.attrib["t"] = "str"
                            cached.text = value
                        else:
                            cell.attrib.pop("t", None)
                            cached.text = str(value)
                    contents = ElementTree.tostring(root, encoding="utf-8", xml_declaration=True)
                output_zip.writestr(item, contents)

    payload = patched.getvalue()
    if hasattr(path, "write"):
        path.seek(0)
        path.truncate(0)
        path.write(payload)
    else:
        with open(path, "wb") as output_file:
            output_file.write(payload)


def build_mse_analysis_ii(analysis: DivisionMSE, path, source_workbook):
    """Add the live MSE summary and checks to the original, formatted division workbook."""
    records = analysis.records.copy()
    if records.empty:
        raise ValueError("The division workbook contains no student-subject MSE records.")
    if records.duplicated(["Class", "Division", "Roll No", "Course"]).any():
        raise ValueError("Duplicate class/division/student/course records prevent a reliable export.")
    if not isinstance(source_workbook, (bytes, bytearray)):
        raise ValueError("The original division workbook is required to preserve its formatted marksheets.")

    book = load_workbook(io.BytesIO(source_workbook), data_only=False)
    if "ISE MSE Analysis" in book.sheetnames or "Checks" in book.sheetnames:
        raise ValueError("The uploaded workbook already contains MSE summary sheets; use the original division workbook.")
    missing_sheets = sorted(set(records["Source Sheet"]) - set(book.sheetnames))
    if missing_sheets:
        raise ValueError(f"Original division worksheet(s) are missing: {missing_sheets}.")

    summary = analysis.summary
    class_order = {"First Year": 0, "Second Year": 1, "Third Year": 2, "Fourth Year": 3, "Final Year": 4}
    division_order = {"A": 0, "B": 1, "C": 2, "D": 3}
    summary["__class_order"] = summary["Class"].map(class_order).fillna(99)
    summary["__division_order"] = summary["Division"].map(division_order).fillna(99)
    course_order = {}
    for class_name in records["Class"].drop_duplicates():
        class_divisions = records.loc[records["Class"] == class_name, "Division"].drop_duplicates()
        first_division = min(class_divisions, key=lambda value: division_order.get(value, 99))
        ordered_courses = records.loc[
            records["Class"].eq(class_name) & records["Division"].eq(first_division),
            "Course",
        ].drop_duplicates().tolist()
        course_order[class_name] = {course: index for index, course in enumerate(ordered_courses)}
    summary["__course_order"] = summary.apply(
        lambda row: course_order[row["Class"]].get(row["Course"], 99), axis=1
    )
    summary = summary.sort_values(
        ["__class_order", "__division_order", "__course_order"], kind="stable"
    ).drop(columns=["__class_order", "__division_order", "__course_order"]).reset_index(drop=True)
    student_keys = records[["Class", "Division", "Roll No", "Student Name"]].drop_duplicates()
    if len(student_keys) != records.groupby(["Class", "Division", "Roll No"]).ngroups:
        raise ValueError("A student identifier maps to conflicting names in the source records.")

    report = book.create_sheet("ISE MSE Analysis", 0)
    checks = book.create_sheet("Checks", 1)
    formula_caches = {"ISE MSE Analysis": {}, "Checks": {}}
    report.sheet_view.showGridLines = False
    report.sheet_view.zoomScale = 90
    checks.sheet_view.showGridLines = False

    thin = Side(style="thin", color="000000")
    cell_border = Border(left=thin, right=thin, top=thin, bottom=thin)
    title_fill = PatternFill("solid", fgColor="FF9DC3E6")
    header_fill = PatternFill("solid", fgColor="FFB4C7E7")
    title_font = Font(name="Calibri", size=14, bold=True)
    header_font = Font(name="Calibri", size=12, bold=True)
    body_font = Font(name="Calibri", size=12)

    def style_range(sheet, cell_range, fill, font):
        for row in sheet[cell_range]:
            for cell in row:
                cell.fill = fill
                cell.font = font
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.border = cell_border

    report.merge_cells("A1:G1")
    report["A1"] = "MSE Performance Analysis (Odd Semester)"
    report.merge_cells("A2:G2")
    report["A2"] = f"Academic Year : {analysis.academic_year}"
    style_range(report, "A1:G2", title_fill, title_font)
    for cell_range, value in (
        ("A3:A4", "Class"), ("B3:B4", "Division (if any)"),
        ("C3:C4", "Course Name"), ("D3:G3", "MSE Performance"),
    ):
        report.merge_cells(cell_range)
        report[cell_range.split(":")[0]] = value
        style_range(report, cell_range, header_fill, header_font)
    for col, title in enumerate(
        ["Students Appeared", "Student Absent", "Students Passed", "Students Failed"], 4
    ):
        report.cell(4, col, title)
    style_range(report, "D4:G4", header_fill, header_font)
    for row in (1, 2, 3, 4):
        report.row_dimensions[row].height = 24

    check_headers = [
        "Class", "Div", "Course", "Total students", "Appeared + Absent",
        "Passed + Failed", "Appeared", "Status",
    ]
    checks["A1"] = "Cross-checks: Appeared+Absent must equal total students; Passed+Failed must equal Appeared"
    checks["A1"].font = Font(name="Calibri", size=12, bold=True)
    for col, title in enumerate(check_headers, 1):
        checks.cell(4, col, title)
    for cell in checks[4]:
        cell.fill, cell.font, cell.border = header_fill, Font(name="Calibri", size=11, bold=True), cell_border
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    for row_index, result in enumerate(summary.to_dict("records"), 5):
        key = (result["Class"], result["Division"], result["Course"])
        group = records[
            records["Class"].eq(key[0])
            & records["Division"].eq(key[1])
            & records["Course"].eq(key[2])
        ]
        source_sheets = group["Source Sheet"].drop_duplicates().tolist()
        source_columns = group["Source Column"].drop_duplicates().tolist()
        if len(source_sheets) != 1 or len(source_columns) != 1:
            raise ValueError(f"{key}: source marks do not map to one unique worksheet column.")
        sheet_name = source_sheets[0]
        sheet = book[sheet_name]
        source_column = int(source_columns[0])
        mark_column = L(source_column)
        quoted_sheet = "'" + sheet_name.replace("'", "''") + "'"
        marks_ref = f"{quoted_sheet}!${mark_column}${FIRST_ROW}:${mark_column}${MAX_ROW}"
        names_ref = f"{quoted_sheet}!$C${FIRST_ROW}:$C${MAX_ROW}"

        previous = summary.iloc[row_index - 6] if row_index > 5 else None
        report.cell(row_index, 1, result["Class"] if previous is None or previous["Class"] != result["Class"] else None)
        new_division = previous is None or (previous["Class"], previous["Division"]) != (result["Class"], result["Division"])
        report.cell(row_index, 2, result["Division"] if new_division else None)
        report.cell(row_index, 3, result["Course"])
        def is_empty_source_mark(value):
            return pd.isna(value) or (
                isinstance(value, str) and value.strip() in {"", "-", "—", "–"}
            )
        all_empty = group["Source Mark"].map(is_empty_source_mark).all()
        if all_empty:
            for col in range(4, 8):
                report.cell(row_index, col, "-")
        else:
            appeared = int(group["Outcome"].isin(["Pass", "Fail"]).sum())
            absent = int(group["Outcome"].eq("Absent").sum())
            passed = int(group["Outcome"].eq("Pass").sum())
            failed = int(group["Outcome"].eq("Fail").sum())
            report.cell(row_index, 4, f"=COUNT({marks_ref})")
            report.cell(row_index, 5, f"=COUNTA({names_ref})-D{row_index}")
            report.cell(row_index, 6, f'=COUNTIF({marks_ref},">="&{PASS_MARK})')
            report.cell(row_index, 7, f'=COUNTIF({marks_ref},"<"&{PASS_MARK})')
            formula_caches["ISE MSE Analysis"].update({
                f"D{row_index}": appeared,
                f"E{row_index}": absent,
                f"F{row_index}": passed,
                f"G{row_index}": failed,
            })
        for cell in report[row_index][:7]:
            cell.font, cell.border = body_font, cell_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        report.row_dimensions[row_index].height = 21.75
        check_row = row_index
        for col, value in enumerate([result["Class"], result["Division"], result["Course"]], 1):
            checks.cell(check_row, col, value)
        if all_empty:
            checks.cell(check_row, 4, f"=COUNTA({quoted_sheet}!$C${FIRST_ROW}:$C${MAX_ROW})")
            for col in range(5, 8):
                checks.cell(check_row, col, "-")
            checks.cell(check_row, 8, "OK")
            formula_caches["Checks"].update({
                f"D{check_row}": len(group),
            })
        else:
            checks.cell(check_row, 4, f"=COUNTA({quoted_sheet}!$C${FIRST_ROW}:$C${MAX_ROW})")
            checks.cell(check_row, 5, f"='ISE MSE Analysis'!D{row_index}+'ISE MSE Analysis'!E{row_index}")
            checks.cell(check_row, 6, f"='ISE MSE Analysis'!F{row_index}+'ISE MSE Analysis'!G{row_index}")
            checks.cell(check_row, 7, f"='ISE MSE Analysis'!D{row_index}")
            checks.cell(check_row, 8, f'=IF(AND(D{check_row}=E{check_row},F{check_row}=G{check_row}),"OK","MISMATCH")')
            formula_caches["Checks"].update({
                f"D{check_row}": len(group),
                f"E{check_row}": appeared + absent,
                f"F{check_row}": passed + failed,
                f"G{check_row}": appeared,
                f"H{check_row}": "OK" if len(group) == appeared + absent and appeared == passed + failed else "MISMATCH",
            })
        for cell in checks[check_row][:8]:
            cell.font, cell.border = Font(name="Calibri", size=11), cell_border
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        checks.row_dimensions[check_row].height = 15

    row_index = 5
    while row_index < 5 + len(summary):
        current_class = summary.iloc[row_index - 5]["Class"]
        class_end = row_index + 1
        while class_end < 5 + len(summary) and summary.iloc[class_end - 5]["Class"] == current_class:
            class_end += 1
        if class_end - row_index > 1:
            report.merge_cells(start_row=row_index, start_column=1, end_row=class_end - 1, end_column=1)
        div_start = row_index
        while div_start < class_end:
            current_division = summary.iloc[div_start - 5]["Division"]
            div_end = div_start + 1
            while div_end < class_end and summary.iloc[div_end - 5]["Division"] == current_division:
                div_end += 1
            if div_end - div_start > 1:
                report.merge_cells(start_row=div_start, start_column=2, end_row=div_end - 1, end_column=2)
            div_start = div_end
        row_index = class_end

    note_row = 5 + len(summary) + 1
    report.cell(note_row, 1,
                'Rules: marks >= 12 = Pass, < 12 = Fail; "AB", blank, or dash = Absent. '
                'Courses with no marks entered are shown as "-". Counts are live formulas on the division sheets.')
    report.cell(note_row, 1).font = Font(name="Calibri", size=10, italic=True)
    for col, width in enumerate([16, 14, 20, 16, 18, 16, 12], 1):
        report.column_dimensions[L(col)].width = width
    report.freeze_panes = "D5"
    report.sheet_properties.pageSetUpPr.fitToPage = True
    report.page_setup.orientation = "landscape"
    report.page_setup.fitToWidth = 1
    report.page_setup.fitToHeight = 1
    report.print_title_rows = "1:4"
    for col, width in enumerate([14, 8, 12, 16, 18, 16, 12, 24], 1):
        checks.column_dimensions[L(col)].width = width
    checks.freeze_panes = "A5"

    book.calculation.fullCalcOnLoad = True
    book.calculation.forceFullCalc = True
    book.calculation.calcMode = "auto"
    book.active = 0
    _save_workbook_with_formula_caches(book, path, formula_caches)
    return path

def build_report(ds, scheme, path):
    wb, subs = ds.wb, scheme["subjects"]
    missing = sorted(set(wb.subjects.abbr) - set(subs))
    if missing: raise ValueError(f"Scheme has no entry for sheet subject(s): {missing}")
    keys = [k for k in subs if k in set(wb.subjects.abbr)]                       # scheme order = staff column order
    label = {k: subs[k]["label"] for k in keys}
    c0, c1 = 5, 4 + len(keys); col = {k: L(c0 + i) for i, k in enumerate(keys)}      # subject columns E.. (same letters on both sheets)
    tot, cg, gr, stc, cl, rk = (L(c1 + i) for i in range(1, 7))
    m = wb.marks.merge(wb.subjects[["term", "code", "abbr"]], on=["term", "code"])
    S = wb.students.merge(wb.cumulative, on="enroll", how="left").reset_index(drop=True)
    order = {e: i for i, e in enumerate(S.enroll)}; korder = {k: i for i, k in enumerate(keys)}
    m = m.assign(o=m.enroll.map(order), k=m.abbr.map(korder)).sort_values(["o", "k"]).reset_index(drop=True)
    n, nd = len(S), len(m); r0, r1 = 6, 5 + n; d1 = nd + 1                      # student rows / last Data row
    book = Workbook(); ra = book.active; ra.title = "Result_Analysis"; sw = book.create_sheet("subjectwise"); dt = book.create_sheet("Data"); nt = book.create_sheet("Notes")
    rng = lambda c: f"Data!${c}$2:${c}${d1}"

    # ---- Data: one row per student x subject taken
    for j, h in enumerate(["Enrollment", "Subject", "Group", "Total", "Max", "Status", "Ratio", "Counted in total"], 1): _put(dt, f"{L(j)}1", h, bold=True, fill=GREY)
    for i, r in enumerate(m.itertuples(), 2):
        k = r.abbr
        for j, v in enumerate([r.enroll, label[k], subs[k]["group"], None if pd.isna(r.total) else float(r.total)], 1): _put(dt, f"{L(j)}{i}", v, font=BLUE, border=False, align="left" if j < 3 else "center")
        _put(dt, f"E{i}", f"=INDEX(subjectwise!${L(c0)}$9:${L(c1)}$9,MATCH(B{i},subjectwise!${L(c0)}$8:${L(c1)}$8,0))", border=False)
        _put(dt, f"F{i}", r.status, font=BLUE, border=False)
        _put(dt, f"G{i}", f'=IF(AND(ISNUMBER(D{i}),E{i}>0),D{i}/E{i},"")', fmt="0.000", border=False)
        _put(dt, f"H{i}", "Y" if subs[k]["counted"] else "N", font=BLUE, border=False)
    for c, w in zip("ABCDEFGH", (14, 12, 9, 8, 8, 8, 8, 16)): dt.column_dimensions[c].width = w
    dt.freeze_panes = "A2"

    # ---- subjectwise
    last = L(c1)
    for i, t in enumerate(scheme["header"], 1): _merge(sw, f"A{i}:{last}{i}" if i < 3 else f"A{i}:{last}{i}", t, bold=True, border=False)
    _merge(sw, f"A5:{last}5", "Subject Wise Result Analysis", bold=True, border=False)
    _merge(sw, "A7:D8", "Subjects", bold=True, fill=GREY)
    for g, gl in GROUPS:
        cs = [col[k] for k in keys if subs[k]["group"] == g]
        if cs: _merge(sw, f"{cs[0]}7:{cs[-1]}7", gl, bold=True, fill=GREY)
    for k in keys:
        _put(sw, f"{col[k]}8", label[k], bold=True, fill=GREY); _put(sw, f"{col[k]}9", subs[k]["max"], font=BLUE)
    _merge(sw, "A9:D9", "Maximum marks", bold=True)
    B = rng("B")
    rows = [("No of Students Appeared", '=COUNTIFS({B},{c}$8)'),
            ("No of Result Declared", '=COUNTIFS({B},{c}$8,{F},"?*")'),
            ("No. of Students with FCWD", '=COUNTIFS({B},{c}$8,{F},"Pass",{G},">="&$E$22)'),
            ("No. of Students with First Class", '=COUNTIFS({B},{c}$8,{F},"Pass",{G},">="&$E$23,{G},"<"&$E$22)'),
            ("No. of Students with Second Class", '=COUNTIFS({B},{c}$8,{F},"Pass",{G},">="&$E$24,{G},"<"&$E$23)'),
            ("No. of Students with pass class", '=COUNTIFS({B},{c}$8,{F},"Pass",{G},"<"&$E$24)'),
            ("No. of Students Failed", '=COUNTIFS({B},{c}$8,{F},"Fail")'),
            ("No of Students Passed", '=COUNTIFS({B},{c}$8,{F},"Pass")'),
            ("No of Students Absent", '=COUNTIFS({B},{c}$8,{F},"AB")'),
            ("Subject percentage %", '=IF({c}10=0,"",ROUND({c}17/{c}10*100,2))')]
    for i, (lab, f) in enumerate(rows, 10):
        _merge(sw, f"A{i}:D{i}", lab, bold=True)
        for k in keys: _put(sw, f"{col[k]}{i}", f.format(B=B, F=rng("F"), G=rng("G"), c=col[k]), fmt="0.00" if i == 19 else "0")
    for i, (lab, v) in enumerate([("FCWD from (% of max marks)", scheme["bands"]["fcwd"]), ("First Class from", scheme["bands"]["first"]), ("Second Class from", scheme["bands"]["second"])], 22):
        _merge(sw, f"A{i}:D{i}", lab, bold=True); _put(sw, f"E{i}", v, font=BLUE, fmt="0%")
    _merge(sw, "F22:M24", "Bands apply to students who passed the subject; passers below the Second Class band are counted as pass class.", border=False)
    for j, (role, nm) in enumerate(scheme["signatures"].items()):
        a = 1 if j == 0 else 8
        _merge(sw, f"{L(a)}27:{L(a + 3)}27", role, bold=True, border=False); _merge(sw, f"{L(a)}28:{L(a + 3)}28", nm, border=False)
    sw.column_dimensions["A"].width = sw.column_dimensions["B"].width = sw.column_dimensions["C"].width = sw.column_dimensions["D"].width = 9
    for k in keys: sw.column_dimensions[col[k]].width = 9
    sw.freeze_panes = "E10"

    # ---- Result_Analysis
    for ref, t in (("A", "Sr No"), ("B", "Student Unique No"), ("C", "Student Seat No"), ("D", "Student Name")): _merge(ra, f"{ref}1:{ref}5", t, bold=True, fill=GREY)
    for g, gl in GROUPS:
        cs = [col[k] for k in keys if subs[k]["group"] == g]
        if cs: _merge(ra, f"{cs[0]}1:{cs[-1]}2", gl, bold=True, fill=GREY)
    for k in keys:
        _put(ra, f"{col[k]}3", f"=subjectwise!{col[k]}8", bold=True, fill=GREY); _put(ra, f"{col[k]}4", f"=subjectwise!{col[k]}9"); _put(ra, f"{col[k]}5", "Tot.", bold=True, fill=GREY)
    _merge(ra, f"{tot}1:{tot}4", "Total", bold=True, fill=GREY); _put(ra, f"{tot}5", scheme["total_max"], font=BLUE, bold=True)
    for c, t in ((cg, "CGPA"), (gr, "GRADE"), (stc, "STATUS"), (cl, "Class")): _merge(ra, f"{c}1:{c}5", t, bold=True, fill=GREY)
    _merge(ra, f"{rk}1:{rk}5", "Rank key (helper)", fill=GREY)
    A, EN = rng("A"), rng("A")
    for i, s in enumerate(S.itertuples(), r0):
        for c, v in zip("ABCD", [i - 5, s.enroll, getattr(s, "seat", ""), s.name]): _put(ra, f"{c}{i}", v, font=BLUE, align="left" if c == "D" else None)
        for k in keys:
            c = col[k]
            _put(ra, f"{c}{i}", f'=IF(COUNTIFS({A},$B{i},{B},{c}$3)=0,"",SUMIFS({rng("D")},{A},$B{i},{B},{c}$3))')
        _put(ra, f"{tot}{i}", f'=SUMIFS({rng("D")},{A},$B{i},{rng("H")},"Y")', bold=True)
        _put(ra, f"{cg}{i}", None if pd.isna(s.cgpa) else float(s.cgpa), font=BLUE, fmt="0.00"); _put(ra, f"{gr}{i}", s.grade, font=BLUE); _put(ra, f"{stc}{i}", s.status, font=BLUE)
        _put(ra, f"{cl}{i}", CLASS_F.format(s=f"{stc}{i}", g=f"{cg}{i}")); _put(ra, f"{rk}{i}", f"={cg}{i}-ROW()/10000000", fmt="0.0000000")
    ra.column_dimensions[rk].hidden = True
    ra.column_dimensions["B"].width, ra.column_dimensions["D"].width = 14, 30
    for k in keys: ra.column_dimensions[col[k]].width = 7
    for c in (tot, cg, gr, stc, cl): ra.column_dimensions[c].width = 12
    ra.freeze_panes = f"{L(c0)}{r0}"
    s0 = r1 + 4; CL, ST = f"${cl}${r0}:${cl}${r1}", f"${stc}${r0}:${stc}${r1}"
    summary = [("No of Students Appeared", f"=COUNTA($D${r0}:$D${r1})"), ("No of Result Declared", f'=COUNTIF({ST},"?*")'),
               ("No. of Students with FCWD", f'=COUNTIF({CL},"Distinction")'), ("No. of Students with First Class", f'=COUNTIF({CL},"First Class")'),
               ("No. of Students with Second Class", f'=COUNTIF({CL},"Second Class")'), ("No. of Students with pass class", f'=COUNTIF({CL},"Pass")'),
               ("No. of Students Failed", f'=COUNTIF({CL},"Fail")'), ("No of Students Passed", f"=F{s0}-F{s0 + 6}"), ("No of Students Absent", f'=COUNTIF({ST},"AB")'),
               ("Percentage %", f"=IF(F{s0}=0,\"\",ROUND(F{s0 + 7}/F{s0}*100,2))")]
    for i, (lab, f) in enumerate(summary, s0):
        _merge(ra, f"C{i}:E{i}", lab, bold=True); _put(ra, f"F{i}", f, fmt="0.00" if "Percentage" in lab else "0")
    RK = f"${rk}${r0}:${rk}${r1}"
    _merge(ra, f"H{s0}:M{s0}", "Toppers (by CGPA)", bold=True, fill=GREY)
    for j, o in enumerate(["1st", "2nd", "3rd"], 1):
        i = s0 + j; _put(ra, f"H{i}", o, bold=True)
        _merge(ra, f"I{i}:K{i}", f"=INDEX($D${r0}:$D${r1},MATCH(LARGE({RK},{j}),{RK},0))"); _put(ra, f"L{i}", f"=INDEX(${cg}${r0}:${cg}${r1},MATCH(LARGE({RK},{j}),{RK},0))", fmt="0.00")
    # ---- Notes
    notes = ["Where the numbers come from", f"Marks, status, CGPA, grade: the ERP sheet (blue = typed-in inputs). Maximum marks, groups, credit flags: {scheme['source']}.",
             f"Total (max {scheme['total_max']}) leaves out the subjects flagged N on the Data sheet (audit / non-credit: FST IV and LL).",
             "Class = Fail if the sheet's Status is not Pass; otherwise by CGPA: 8+ Distinction, 7+ First, 6+ Second, 5+ Pass.",
             "Subjectwise classes use marks / maximum marks of passers; bands are the blue cells on subjectwise (E22:E24).",
             "Change a maximum on subjectwise row 9 and every ratio and class count follows."]
    for i, t in enumerate(notes, 1): nt[f"A{i}"].value = t; nt[f"A{i}"].font = FONT(bold=i == 1)
    nt.column_dimensions["A"].width = 130
    book.save(path)
    return path

def main(argv):
    from .store import Dataset
    xlsx, sch, out = argv[:3]
    build_report(Dataset.from_workbook(xlsx), json.load(open(sch)), out); print("wrote", out)

if __name__ == "__main__": main(sys.argv[1:])
