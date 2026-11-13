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

# ---- plain, staff-style report for ANY parsed workbook (no scheme file needed) -------------------------------------
HIGHLIGHT = "FFF2CC"                       # the only colour used: title + header band (Excel light gold); change freely
CLASS_BANDS = ((8.0, "Distinction"), (7.0, "First Class"), (6.0, "Second Class"), (5.0, "Pass Class"))   # CGPA (or SGPA) >= value
SUBJECT_CLASS = {"FCWD": ("O", "A+", "A"), "First Class": ("B+", "B"), "Second Class": ("C",), "Pass Class": ("P",)}   # subject grade -> class
STD_MAX = (10, 15, 20, 25, 30, 40, 50, 60, 70, 75, 80, 100)
PASS_FRACTION = 0.4                        # passing marks / maximum marks, used only when the sheet prints no maximum
DEFAULT_COLLEGE = "D. Y. PATIL COLLEGE OF ENGG. & TECH, KOLHAPUR"
DEFAULT_DEPARTMENT = "DEPARTMENT OF CSE (Artificial Intelligence and Machine Learning)"   # pass department= to build_generic_report per file
GROUP_ORDER = ("Theory", "MDM", "OE", "Elective", "Laboratory")
_PLAIN = lambda **k: Font(name="Calibri", size=11, **k)
_MED, _THIN = Side(style="medium"), Side(style="thin")

def _group(row):
    """Theory / MDM / OE / Laboratory from the subject's own type, name and code; unknown -> Theory."""
    name, code, typ = str(row.label), str(row.code), str(getattr(row, "type", ""))
    code = re.sub(r"\s+", "", code)
    if typ == "MDM" or "MDM" in code.upper() or re.search(r"\(\s*MDM\s*\)", name, re.I): return "MDM"
    if typ == "OEC" or re.search(r"OEC|\(\s*OE\s*\)", code + " " + name, re.I): return "OE"
    if pd.notna(getattr(row, "credits", float("nan"))) and row.credits == 0: return "Theory"      # zero-credit audit courses (FST, LLC ...) sit with Theory
    if getattr(row, "part", "") == "Practical" or re.search(r"\blab\b|\(lab\)|lab$", name, re.I) or re.search(r"P\d{3}$", code): return "Laboratory"
    if re.search(r"PEC|PE[LM]\d{3}$|ELECTIVE", code.upper() + " " + name.upper()): return "Elective"
    return "Theory"

def _max_marks(row, observed):
    """Maximum marks of a subject: printed maximum if the sheet has one, else passing marks / 40 %, never below the highest mark seen."""
    if pd.notna(row.max_total) and row.max_total > 0: return float(row.max_total)
    est = (float(row.int_pass or 0) + float(row.ext_pass or 0)) / PASS_FRACTION if pd.notna(row.int_pass) or pd.notna(row.ext_pass) else 0
    if pd.isna(observed): return est or None
    if est >= observed and est > 0: return est
    return next((float(m) for m in STD_MAX if m >= observed), None)

def _grade_class(g):
    g = re.sub(r"[^A-Za-z+]", "", str(g)).upper()
    return next((c for c, gs in SUBJECT_CLASS.items() if g in gs), None)

def rank_metric(student_report):
    """CGPA is the first priority whenever the file has it; SGPA only when there is no CGPA at all."""
    for m in ("cgpa", "sgpa"):
        if m in student_report and student_report[m].notna().any(): return m
    return None

def _class_of(score):
    return "" if pd.isna(score) else next((n for lim, n in CLASS_BANDS if score >= lim), "")

def _write_table(ws, title, sub, frame, cols, groups, maxrow, total_max, last_fixed):
    """cols: [(key, heading)] fixed/summary columns; groups: [(group, [(key, subject, max)])]. Header = rows 1-5 like the staff sheet."""
    box = Border(_THIN, _THIN, _THIN, _THIN)
    fixed = [c for c in cols if c[0] in ("rank", "enroll", "name")]; tail = [c for c in cols if c not in fixed]
    n_sub = sum(len(g[1]) for g in groups); last = len(fixed) + n_sub + len(tail)
    def put(r, c, v, bold=False, fill=False, edge=True, fmt=None):
        cell = ws.cell(r, c, v); cell.font = _PLAIN(bold=bold); cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if edge: cell.border = box
        if fill: cell.fill = PatternFill("solid", fgColor=HIGHLIGHT)
        if fmt: cell.number_format = fmt
        return cell
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=last); put(1, 1, title, True, True)
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=last); put(2, 1, sub, False, False, False)
    ws.cell(2, 1).alignment = Alignment(horizontal="center")
    for i, (k, h) in enumerate(fixed, 1):
        ws.merge_cells(start_row=3, start_column=i, end_row=7, end_column=i); put(3, i, h, True, True)
        for r in range(3, 8): ws.cell(r, i).border = box
    c = len(fixed) + 1; keycol = {}
    for g, items in groups:
        if g and len(items):
            ws.merge_cells(start_row=3, start_column=c, end_row=4, end_column=c + len(items) - 1); put(3, c, g, True, True)
            for r in (3, 4):
                for cc in range(c, c + len(items)): ws.cell(r, cc).border = box
        for k, subj, mx in items:
            put(5, c, subj, True, True); put(6, c, mx); put(7, c, "Tot.", True, True); keycol[k] = c; c += 1
    for k, h in tail:
        ws.merge_cells(start_row=3, start_column=c, end_row=6, end_column=c); put(3, c, h, True, True)
        for r in range(3, 7): ws.cell(r, c).border = box
        put(7, c, total_max if k == "total" else None, True); keycol[k] = c; c += 1
    for i, k in enumerate([k for k, _ in fixed], 1): keycol[k] = i
    for r, row in enumerate(frame.to_dict("records"), 8):
        for k, col in keycol.items():
            v = row.get(k); v = None if v is None or (isinstance(v, float) and pd.isna(v)) else v
            cell = ws.cell(r, col, v); cell.font = _PLAIN(); cell.border = box
            if isinstance(v, float): cell.number_format = "0.00" if k in ("cgpa", "sgpa") else "0.##"
            cell.alignment = Alignment(horizontal="left" if k == "name" else "center", vertical="center")
    ws.column_dimensions["A"].width = 7; ws.column_dimensions["B"].width = 16; ws.column_dimensions["C"].width = 30
    for col in range(len(fixed) + 1, last + 1): ws.column_dimensions[L(col)].width = 11
    ws.freeze_panes = f"{L(len(fixed) + 1)}8"

def _banner(ws, last, rows):
    box = Border(_THIN, _THIN, _THIN, _THIN)
    for r, text in rows:
        ws.merge_cells(start_row=r, start_column=1, end_row=r, end_column=last)
        c = ws.cell(r, 1, text); c.font = _PLAIN(bold=True); c.alignment = Alignment(horizontal="center", vertical="center")
        if r == rows[0][0]: c.fill = PatternFill("solid", fgColor=HIGHLIGHT)

def _subjectwise_sheet(ws, college, department, term, groups, items, marks):
    """Subject Wise Result Analysis: one column per subject; counts come from the subject's own status and grade columns."""
    box = Border(_THIN, _THIN, _THIN, _THIN)
    keys = [k for _, its in groups for k, *_ in its]; last = 4 + len(keys)
    _banner(ws, last, [(1, college), (2, department), (3, f"RESULT ANALYSIS (Term - {term})")])
    ws.merge_cells(start_row=5, start_column=1, end_row=5, end_column=last)
    c = ws.cell(5, 1, "Subject Wise Result Analysis"); c.font = _PLAIN(bold=True); c.alignment = Alignment(horizontal="center")
    def put(r, c, v, bold=False, fill=False, fmt=None):
        cell = ws.cell(r, c, v); cell.font = _PLAIN(bold=bold); cell.border = box; cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        if fill: cell.fill = PatternFill("solid", fgColor=HIGHLIGHT)
        if fmt: cell.number_format = fmt
    ws.merge_cells("A7:D8"); put(7, 1, "Subjects", True, True)
    for r in (7, 8):
        for cc in range(1, 5): ws.cell(r, cc).border = box
    col = 5; code_of = {k: code for k, _, _, _, code, _ in items}; mx_of = {k: m for k, _, m, *_ in items}
    labels = ["No of Students Appeared", "No of Result Declared", "No. of Students with FCWD", "No. of Students with First Class",
              "No. of Students with Second Class", "No. of Students with pass class", "No. of Students with Failed",
              "No of Students Passed", "No of Students Absent", "Subject percentage   %"]
    ws.merge_cells("A9:D9"); put(9, 1, None)
    for i, t in enumerate(labels, 10):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=4); put(i, 1, t, True)
        for cc in range(1, 5): ws.cell(i, cc).border = box
    for g, its in groups:
        if its:
            ws.merge_cells(start_row=7, start_column=col, end_row=7, end_column=col + len(its) - 1); put(7, col, g, True, True)
            for cc in range(col, col + len(its)): ws.cell(7, cc).border = box
        for k, name, mx, *_ in its:
            m = marks[marks.code == code_of[k]]; st = m.status.astype(str).str.strip()
            passed, failed, absent = int(st.eq("Pass").sum()), int(st.eq("Fail").sum()), int(st.eq("AB").sum())
            appeared = passed + failed + absent
            cls = m[st.eq("Pass")].grade.map(_grade_class) if m.grade.astype(str).str.strip().replace({"nan": "", "None": ""}).ne("").any() else None
            cnt = [None if cls is None else int(cls.eq(c).sum()) for c in ("FCWD", "First Class", "Second Class", "Pass Class")]
            vals = [appeared, passed + failed, *cnt, failed, passed, absent, round(100 * passed / appeared, 2) if appeared else None]
            put(8, col, name, True, True); put(9, col, mx)
            for i, v in enumerate(vals, 10): put(i, col, v, fmt="0.00" if i == 19 else None)
            ws.column_dimensions[L(col)].width = max(9, min(14, len(name) + 2)); col += 1
    for cc in "ABCD": ws.column_dimensions[cc].width = 9
    ws.freeze_panes = "E10"

def _counts_sheet(ws, S, metric, college, department, term):
    """Overall counts (class bands, failed/ATKT, passed, absent, pass %) and the top 3 by the ranking metric."""
    box = Border(_THIN, _THIN, _THIN, _THIN)
    _banner(ws, 10, [(1, college), (2, department), (3, f"RESULT ANALYSIS (Term - {term})")])
    st = S.status.astype(str).str.strip().str.casefold(); n = len(S)
    cls = S["class"]; passed = int(st.eq("pass").sum())
    rows = [("Total students", n)] + [(f"No. of Students with {nm}", int(cls.eq(nm).sum())) for _, nm in CLASS_BANDS] + [
            ("No. of Students Failed / ATKT", int(st.str.contains("fail|atkt").sum())), ("No. of Students Passed", passed),
            ("No. of Students Absent", int(st.isin(["ab", "absent"]).sum())), ("Overall pass percentage", passed / n if n else None)]
    rows[1] = ("No. of Students with Distinction", rows[1][1])
    rows[4] = ("No. of Students with pass class", rows[4][1])
    for i, (lab, v) in enumerate(rows, 5):
        ws.merge_cells(start_row=i, start_column=1, end_row=i, end_column=4)
        c = ws.cell(i, 1, lab); c.font = _PLAIN(bold=True)
        for cc in range(1, 5): ws.cell(i, cc).border = box
        c = ws.cell(i, 5, v); c.font = _PLAIN(); c.border = box; c.alignment = Alignment(horizontal="center")
        if lab == "Overall pass percentage": c.number_format = "0.00%"
    if metric:
        ws.merge_cells("G5:J5"); c = ws.cell(5, 7, f"Top 3 by {metric.upper()}"); c.font = _PLAIN(bold=True); c.fill = PatternFill("solid", fgColor=HIGHLIGHT)
        for cc in range(7, 11): ws.cell(5, cc).border = box
        top = S.dropna(subset=[metric]).sort_values(metric, ascending=False, kind="stable").head(3)
        for i, r in enumerate(top.itertuples(), 6):
            for cc, v in zip(range(7, 11), (f"{i - 5}.", r.name, getattr(r, metric), r.enroll)):
                c = ws.cell(i, cc, v); c.font = _PLAIN(); c.border = box
                c.alignment = Alignment(horizontal="left" if cc == 8 else "center")
                if cc == 9: c.number_format = "0.00"
    for cc, w in zip("ABCDEFGHIJ", (9, 9, 9, 9, 12, 3, 5, 30, 8, 16)): ws.column_dimensions[cc].width = w

def build_generic_report(ds, term, path, college=DEFAULT_COLLEGE, department=DEFAULT_DEPARTMENT):
    """Plain staff-style result analysis for any parsed workbook: one row per student, subjects grouped, max row, Tot. row,
    Total / CGPA / Grade / Status / Class. Ranked by CGPA when the file has one, else by SGPA. Failed students' full rows go on a second sheet."""
    wb = ds.wb
    if term not in set(wb.terms.term): raise ValueError(f"Term {term} is not present in the workbook.")
    subjects = wb.subjects[wb.subjects.term == term].reset_index(drop=True)
    marks = wb.marks[wb.marks.term == term]
    tr = wb.terms[wb.terms.term == term].drop(columns="term_n", errors="ignore")
    ids = list(dict.fromkeys(list(tr.enroll) + list(marks.enroll)))
    if not ids or subjects.empty or marks.empty: raise ValueError(f"Term {term} has no student or subject records to report.")
    if marks.duplicated(["enroll", "code"]).any(): raise ValueError(f"Term {term} contains duplicate student-course rows; resolve these before exporting.")
    S = wb.students[wb.students.enroll.isin(ids)].merge(tr, on="enroll", how="left")
    if term == wb.terms.sort_values("term_n").iloc[-1].term and not wb.cumulative.empty:   # CGPA belongs to the last term of the file
        S = S.merge(wb.cumulative.drop(columns=[c for c in ("status",) if c in S], errors="ignore"), on="enroll", how="left", suffixes=("", "_cum"))
        if "status_cum" in S: S["status"] = S.status_cum
    # subjects: unique display names, grouped Theory / MDM / OE / Laboratory, source order inside a group
    seen, items = set(), []
    for r in subjects.itertuples():
        name = str(r.label).strip() or str(r.abbr).strip() or str(r.code)
        if name.casefold() in seen: name = f"{name} [{r.code}]"
        seen.add(name.casefold()); mx = _max_marks(r, marks[marks.code == r.code].total.max())
        items.append((f"s{len(items)}", name, mx, _group(r), r.code, r.credits))
    kinds = {g for *_, g, _, _ in items}
    groups = [(g, [(k, n, m) for k, n, m, gg, _, _ in items if gg == g]) for g in GROUP_ORDER if g in kinds] if len(kinds) > 1 else [("Subjects", [(k, n, m) for k, n, m, *_ in items])]
    key_of = {code: k for k, _, _, _, code, _ in items}
    piv = marks.assign(k=marks.code.map(key_of)).pivot(index="enroll", columns="k", values="total")
    S = S.merge(piv, left_on="enroll", right_index=True, how="left")
    counted = [k for k, _, m, _, _, cr in items if not (pd.notna(cr) and cr == 0)]          # 0-credit audit subjects stay visible but out of Total
    S["total"] = S["total"] if "total" in S and S["total"].notna().any() else S[counted].sum(axis=1, min_count=1)
    common = set(subjects[subjects.is_common].code) if "is_common" in subjects else {c for *_, c, _ in items}
    mxs = [m for k, _, m, _, c, _ in items if k in counted and c in common]
    only_common = all(c in common for k, _, _, _, c, _ in items if k in counted)
    total_max = sum(mxs) if mxs and only_common and all(m is not None for m in mxs) else None   # optional subjects (MDM / electives) make a sheet-wide maximum ambiguous
    def outcome(g):
        v = g.dropna().astype(str).str.strip()
        return "Fail" if v.eq("Fail").any() else "AB" if v.eq("AB").any() else "Pass" if len(v) and v.eq("Pass").all() else "Unreported"
    if "status" not in S or S.status.replace({"": None, "nan": None}).isna().all(): S["status"] = S.enroll.map(marks.groupby("enroll").status.apply(outcome)).fillna("No result")
    S["status"] = S.status.replace({"": None, "nan": None})
    metric = rank_metric(S)
    if metric:
        S["rank"] = S[metric].rank(method="min", ascending=False)
        S = S.sort_values("rank", kind="stable", na_position="last")
        S["rank"] = S["rank"].map(lambda v: None if pd.isna(v) else int(v))
    else: S["rank"] = range(1, len(S) + 1)
    S["class"] = [_class_of(g) for g in (S[metric] if metric else [float("nan")] * len(S))]
    S = S.reset_index(drop=True)
    cols = [("rank", "Rank"), ("enroll", "Student Unique No"), ("name", "Student Name"), ("total", "Total")]
    if metric: cols.append((metric, metric.upper()))
    if "grade" in S and S.grade.replace({"": None, "nan": None}).notna().any(): cols.append(("grade", "Grade"))
    cols += [("status", "Status"), ("class", "Class")]
    S["grade"] = S["grade"].replace({"nan": None, "": None}) if "grade" in S else None
    book = Workbook(); ws = book.active; ws.title = "Result_Analysis"
    how = f"ranked by {metric.upper()}" if metric else "no CGPA / SGPA in the file, so not ranked"
    _write_table(ws, f"RESULT ANALYSIS - TERM {term}", f"{ds.name}  |  {len(S)} students  |  {how}", S, cols, groups, None, total_max, 3)
    _subjectwise_sheet(book.create_sheet("SUBJECTWISE"), college, department, term, groups, items, marks)
    _counts_sheet(book.create_sheet("Counts"), S, metric, college, department, term)
    bad = S[S.status.astype(str).str.contains("fail", case=False, na=False)]
    fs = book.create_sheet("Failed_Students")
    _write_table(fs, f"FAILED STUDENTS - TERM {term}", f"{len(bad)} of {len(S)} students (complete rows copied from Result_Analysis)", bad, cols, groups, None, total_max, 3)
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
