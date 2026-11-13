"""Stage 1 - universal ingestion for the multi-level ERP mark sheet (.xlsx).
Nothing about a specific semester or branch is hard-coded: terms, subjects, block widths and
column meanings are all discovered from the sheet's own header rows."""
import io, re
from dataclasses import dataclass
import pandas as pd

class UnsupportedLayout(ValueError):
    """Raised when a workbook doesn't look like a layout we have an adapter for."""

TYPES = "PCC|HSSM|MDM|OEC|VSEC|VEC|AEC|CEP|CCA|EEM|MC|FY"
CODE_RX = re.compile(rf"^\d{{3}}(?P<dept>[A-Z]*?)(?P<type>{TYPES})(?P<part>[LPM])?(?P<num>\d{{3}})$")
PARTS = {"L": "Theory", "P": "Practical", "M": "Other", None: "Theory"}
SUMMARY = {"total": "total", "credit reg": "cr_reg", "credit er": "cr_earned", "grade total": "grade_total", "sgpa": "sgpa", "%": "pct"}
FINAL = {"final percentage": "final_pct", "cgpa": "cgpa", "grade": "grade", "status": "status"}
COMPONENTS = {"ise": "ise", "mse": "mse", "int": "int", "ese": "ese", "poe": "poe"}          # mark components a block can have
AFTER = {"gr": "grade", "cr pt": "credits", "cr er": "earned", "pt": "gpoints", "st": "status"}
MAXCOL = {"ise": "max_ise", "mse": "max_mse", "ese": "max_ese", "int": "max_int", "poe": "max_oe_poe"}
NUMF = ["ise", "mse", "internal", "int", "ese", "poe", "other", "total", "credits", "earned", "gpoints"]
ALLF = ["ise", "mse", "internal", "int", "ese", "poe", "other", "total", "grade", "credits", "earned", "gpoints", "status"]
SUBJ_RX = r"\(\s*[A-Za-z0-9][A-Za-z0-9 ]{4,}[A-Za-z0-9](?:-\d+)?(?:\([^()]*\))?\s*\)\s*$"

@dataclass
class Workbook:
    students: pd.DataFrame      # one row per student
    terms: pd.DataFrame         # one row per student x term (totals, credits, SGPA)
    subjects: pd.DataFrame      # one row per term x subject
    marks: pd.DataFrame         # one row per student x subject actually taken
    cumulative: pd.DataFrame    # final %, CGPA, grade, status printed at end of sheet
    info: dict

def _norm(v):
    """Header label -> comparable key: drops '(20)' max-mark notes, dots and case."""
    return re.sub(r"[.\s]+", " ", re.sub(r"\(\s*[\d.]+\s*\)", " ", str(v))).strip().lower()
def _paren(v):
    m = re.search(r"\(\s*([\d.]+)\s*\)", str(v)); return float(m.group(1)) if m else float("nan")
def _num(s): return pd.to_numeric(s, errors="coerce")
def _mse_status(value):
    if pd.isna(value): return "Unreported"
    text = str(value).strip()
    if not text or text.casefold() in {"-", "—", "–", "nan", "none"}: return "Unreported"
    if re.sub(r"[\s.]+", "", text).casefold() in {"ab", "abs", "absent"}: return "AB"
    if pd.notna(pd.to_numeric(value, errors="coerce")): return "Marked"
    return f"Invalid: {text}"
def _limit(v):
    m = re.search(r"\(\s*(\d+)\s*\)", str(v)); return int(m.group(1)) if m else None
def parse_code(code):
    m = CODE_RX.match(str(code))
    return (m["dept"], m["type"], PARTS[m["part"]]) if m else ("", "Other", "Theory")

def _load(src, sheet):
    if isinstance(src, (bytes, bytearray)): src = io.BytesIO(src)
    try:
        return pd.read_excel(src, header=None, sheet_name=sheet)
    except Exception as ex:
        raise UnsupportedLayout(f"Could not open as an .xlsx workbook ({ex.__class__.__name__}). "
                                "Legacy/HTML-style .xls exports need a separate adapter.") from ex

def _layout(raw):
    """Find the student, term, subject-name, and sub-header rows."""
    h = next((i for i in range(min(15, len(raw))) if str(raw.iat[i, 0]).strip().lower().startswith("sr")), None)
    if h is None: raise UnsupportedLayout("No 'Sr No' header row found in the first 15 rows.")
    term_rows = [(i, sum(bool(re.match(r"^term\s*-?\s*\S+", str(v), re.I)) for v in raw.iloc[i]))
                 for i in range(min(max(h + 1, 15), len(raw)))]
    term_row, count = max(term_rows, key=lambda item: item[1], default=(None, 0))
    if not count: raise UnsupportedLayout("No 'Term - X' markers found in the header rows.")
    tcols = [c for c, v in raw.iloc[term_row].items() if re.match(r"^term\s*-?\s*\S+", str(v), re.I)]
    first = tcols[0]
    end = next((c for c in range(first + 1, raw.shape[1]) if _norm(raw.iat[term_row, c]) == "total"), raw.shape[1])
    coded_candidates = [(raw.iloc[i].astype(str).str.contains(SUBJ_RX).sum(), i)
                        for i in range(max(0, h - 4), min(h + 5, len(raw))) if i not in (h, term_row)]
    coded_candidates = [candidate for candidate in coded_candidates if candidate[0]]
    candidates = []
    for i in range(max(0, h - 4), min(h + 5, len(raw))):
        if i in (h, term_row): continue
        values = [_norm(raw.iat[i, c]) for c in range(first, end)
                  if isinstance(raw.iat[i, c], str) and str(raw.iat[i, c]).strip()]
        if values: candidates.append((len(set(values)), i))
    subj = (max(coded_candidates)[1] if coded_candidates else max(candidates, default=(0, None))[1])
    if subj is None: subj = next((i for i in range(h, min(h + 4, len(raw)))
                                  if raw.iloc[i].astype(str).str.contains(SUBJ_RX).sum() >= 1), None)
    rows = range(h + 1, min(h + 12, len(raw)))
    sub = max(rows, key=lambda i: (raw.iloc[i].map(_norm) == "gr").sum(), default=None)
    if subj is None or sub is None or (raw.iloc[sub].map(_norm) == "gr").sum() == 0 or sub <= subj:
        raise UnsupportedLayout("Found the student header but not both a subject-name row and a column-header row containing "
                                "'Gr.' (grade) - this looks like a different layout.")
    return h, term_row, subj, sub

def _col(top, pat, skip=None):
    return next((c for c, v in top.items() if re.search(pat, str(v), re.I) and not (skip and re.search(skip, str(v), re.I))), None)

def _meta(raw, h, term_row, srow, sub, c):
    """Per-subject facts printed in the rows above the column headers (limits, max marks, credits, pass rule)."""
    m = {}
    for r in range(srow + 1, sub):
        if r in (h, term_row): continue
        t = str(raw.iat[r, c])
        if re.search(r"int\w*\s*pass", t, re.I): m["int_pass"] = _limit(t)
        elif re.search(r"ext\w*\s*pass", t, re.I): m["ext_pass"] = _limit(t)
        mm = re.search(r"max\s*marks\s*\(\s*([\d.]+)\s*\)", t, re.I)
        if mm: m["max_total"] = float(mm.group(1))
        mc = re.search(r"credits?\s*\(\s*([\d.]+)\s*\)", t, re.I)
        if mc: m["credits_declared"] = float(mc.group(1))
        if re.match(r"\s*pass\s*:", t, re.I): m["pass_rule"] = re.sub(r"^\s*pass\s*:\s*", "", t, flags=re.I).strip()
    return m

def _fields(labels, where):
    """Column position -> field name for one subject block, anchored on the 'Gr.' column (the total sits just before it)."""
    if "gr" not in labels or labels.index("gr") == 0 or labels[labels.index("gr") - 1] != "tot":
        raise UnsupportedLayout(f"{where}: expected '... Tot., Gr., Cr. Pt., Cr. Er., Pt., St.' columns.")
    gi = labels.index("gr"); names = {gi - 1: "total"}; extra = []
    for i in range(gi - 1):
        lab = labels[i]
        if lab in ("p/f", ""): continue
        f = "internal" if lab == "tot" else COMPONENTS.get(lab, "other")
        if f in names.values(): extra.append(lab)
        else: names[i] = f
    for j in range(gi, len(labels)):
        if labels[j] in AFTER:
            if AFTER[labels[j]] in names.values():
                raise UnsupportedLayout(f"{where}: repeated '{labels[j]}' column - two subject blocks look merged.")
            names[j] = AFTER[labels[j]]
    missing = {"grade", "status", "total"} - set(names.values())
    if missing: raise UnsupportedLayout(f"{where}: header lacks {sorted(missing)}.")
    return names, extra

def _read_flat(raw, include_mse_rows=False):
    """Read normalized one-student/one-course-per-row ERP exports without a semester template."""
    def col_for(row, patterns):
        return next((c for c, value in enumerate(row)
                     if any(re.search(pattern, _norm(value), re.I) for pattern in patterns)), None)

    original = None
    for header in range(min(25, len(raw))):
        row = raw.iloc[header].tolist()
        mse_col = col_for(row, (r"^mse(?: marks?)?$", r"mid.?semester(?: exam)? marks"))
        enrollment_col = col_for(row, (r"\b(enrollment|enrolment|enroll|registration|roll)\b.*\b(no|number|id)\b",
                                       r"^student id$"))
        name_col = col_for(row, (r"^student name$", r"^name$"))
        subject_col = col_for(row, (r"^subject( name)?$", r"^course( name)?$", r"^paper( name)?$", r"^module( name)?$"))
        marks_col = col_for(row, (r"total marks", r"marks obtained", r"obtained marks", r"score", r"^marks$", r"^total$"))
        if marks_col is None and include_mse_rows:
            marks_col = mse_col
        grade_col = col_for(row, (r"^letter grade$", r"^grade$"))
        status_col = col_for(row, (r"^result$", r"^status$", r"^outcome$", r"pass.?fail"))
        has_outcome = grade_col is not None or status_col is not None
        if (all(value is not None for value in (enrollment_col, name_col, subject_col, marks_col))
                and (has_outcome or (include_mse_rows and mse_col is not None))):
            original = (header, row, enrollment_col, name_col, subject_col, marks_col,
                        grade_col, status_col, mse_col)
            break
    if original is None:
        raise UnsupportedLayout(
            "No supported ERP table was found. A normalized table needs one row per student and course, "
            "with student ID/name, subject, marks, and an explicit grade or result/status. "
            "Multi-level ERP sheets with subject blocks are also supported."
        )
    header, labels, enrollment_col, name_col, subject_col, marks_col, grade_col, status_col, mse_col = original
    code_col = col_for(labels, (r"^course code$", r"^subject code$", r"^paper code$", r"^code$"))
    term_col = col_for(labels, (r"^term$", r"^semester$", r"^sem$"))
    sgpa_col = col_for(labels, (r"^sgpa$", r"semester grade point average"))
    cgpa_col = col_for(labels, (r"^cgpa$", r"cumulative grade point average"))
    pct_col = col_for(labels, (r"^percentage$", r"^percent$", r"^result percentage$"))
    credits_col = col_for(labels, (r"^credits?$", r"^credit points?$"))
    max_col = col_for(labels, (r"maximum marks", r"max marks", r"maximum score", r"max score"))
    seat_col = col_for(labels, (r"seat number", r"seat no"))
    body = raw.iloc[header + 1:].copy().reset_index(drop=True)
    body = body[body[enrollment_col].notna() & body[name_col].notna() & body[subject_col].notna()].copy()
    if body.empty:
        raise UnsupportedLayout("Found an ERP table header, but no student/course rows appear below it.")

    def text(value):
        if pd.isna(value): return ""
        if isinstance(value, float) and value.is_integer(): value = int(value)
        return re.sub(r"\s+", " ", str(value)).strip()

    enroll = body[enrollment_col].map(text).str.upper()
    names = body[name_col].map(text)
    term_values = body[term_col].map(text) if term_col is not None else pd.Series("I", index=body.index)
    term_values = term_values.map(lambda value: re.sub(r"^(?:term|semester|sem)\s*[-:]?\s*", "", value, flags=re.I) or "I")
    subject_names = body[subject_col].map(text)
    codes = body[code_col].map(text) if code_col is not None else subject_names
    status_alias = {"pass": "Pass", "passed": "Pass", "p": "Pass", "clear": "Pass", "cleared": "Pass",
                    "fail": "Fail", "failed": "Fail", "f": "Fail", "absent": "AB", "ab": "AB"}
    grades = body[grade_col].map(text).str.upper() if grade_col is not None else pd.Series("", index=body.index)
    if status_col is not None:
        statuses = body[status_col].map(text).map(lambda value: status_alias.get(value.casefold(), value))
    else:
        statuses = grades.map(lambda value: "Fail" if value in ("F", "FAIL") else "AB" if value in ("AB", "ABSENT") else "Pass")
    keep = statuses.ne("") | grades.ne("") | body[marks_col].notna()
    if include_mse_rows and mse_col is not None:
        keep |= body[mse_col].map(_mse_status).isin(["Marked", "AB"])
    body, enroll, names, term_values, subject_names, codes, grades, statuses = (
        value.loc[keep].reset_index(drop=True) for value in
        (body, enroll, names, term_values, subject_names, codes, grades, statuses))
    if body.empty:
        raise UnsupportedLayout("The ERP table contains no rows with a mark, grade, or result status.")

    term_order = list(dict.fromkeys(term_values))
    term_numbers = {term: index + 1 for index, term in enumerate(term_order)}
    students = pd.DataFrame({"enroll": enroll, "name": names})
    if seat_col is not None:
        seat = body[seat_col].map(text)
        students["seat"] = seat
    students = students.drop_duplicates("enroll").reset_index(drop=True)
    students["prefix"] = students.enroll.str.extract(r"^([A-Z]+)")[0]
    students["batch"] = 2000 + pd.to_numeric(students.enroll.str.extract(r"^[A-Z]+(\d{2})")[0], errors="coerce")

    mse_source = body[mse_col].map(text) if mse_col is not None else pd.Series("", index=body.index)
    mark_rows = pd.DataFrame({
        "enroll": enroll, "term": term_values, "code": codes,
        "total": _num(body[marks_col]).to_numpy(), "mse": _num(body[mse_col]).to_numpy() if mse_col is not None else float("nan"),
        "mse_status": mse_source.map(_mse_status).to_numpy(), "grade": grades.to_numpy(),
        "status": statuses.to_numpy(),
    })
    for field in ("ise", "internal", "int", "ese", "poe", "other", "credits", "earned", "gpoints"):
        mark_rows[field] = float("nan")
    if credits_col is not None: mark_rows["credits"] = _num(body[credits_col]).to_numpy()
    mark_rows = mark_rows[["enroll", "term", "code"] + ALLF + ["mse_status"]]

    subject_lookup = pd.DataFrame({"term": term_values, "code": codes, "abbr": subject_names})
    subject_lookup = subject_lookup.drop_duplicates(["term", "code"], keep="first")
    subjects = []
    for row in subject_lookup.itertuples(index=False):
        group = body[(term_values == row.term) & (codes == row.code)]
        result = mark_rows[(mark_rows.term == row.term) & (mark_rows.code == row.code)]
        maximum = _num(group[max_col]).dropna().median() if max_col is not None else float("nan")
        credits = _num(group[credits_col]).dropna().median() if credits_col is not None else float("nan")
        dept, typ, part = parse_code(row.code)
        subjects.append({"term": row.term, "term_n": term_numbers[row.term], "abbr": row.abbr,
                         "code": row.code, "dept": dept, "type": typ, "part": part,
                         "int_pass": None, "ext_pass": None, "enrolled": len(result),
                         "credits": credits, "max_total": maximum, "pass_rule": "",
                         "has_mse": mse_col is not None,
                         "has_mse": mse_col is not None,
                         **{f"max_{key.lower()}": float("nan") for key in ("ise", "mse", "ese", "int", "oe_poe")}})
    subjects = pd.DataFrame(subjects)
    subjects["share"] = subjects.enrolled.div(students.groupby("enroll").size().sum()).round(3)
    subjects["is_common"] = subjects.share >= 0.9
    subjects["label"] = subjects.abbr
    subjects["name"] = subjects.abbr

    term_values_frame = pd.DataFrame({"enroll": enroll, "term": term_values})
    for field, source_col in (("sgpa", sgpa_col), ("pct", pct_col)):
        term_values_frame[field] = _num(body[source_col]).to_numpy() if source_col is not None else float("nan")
    if credits_col is not None:
        term_values_frame["cr_reg"] = _num(body[credits_col]).to_numpy()
    else:
        term_values_frame["cr_reg"] = float("nan")
    term_values_frame["cr_earned"] = float("nan")
    term_values_frame["grade_total"] = float("nan")
    terms = term_values_frame.groupby(["enroll", "term"], sort=False, as_index=False).first()
    terms["term_n"] = terms.term.map(term_numbers)
    cumulative = pd.DataFrame({"enroll": students.enroll})
    if cgpa_col is not None:
        cum_values = pd.DataFrame({"enroll": enroll, "cgpa": _num(body[cgpa_col])}).dropna(subset=["cgpa"]).drop_duplicates("enroll", keep="last")
        cumulative = cumulative.merge(cum_values, on="enroll", how="left")
    info = {"students": len(students), "terms": term_order, "subjects": len(subjects), "rows": len(raw),
            "mse_available": mse_col is not None,
            "mse_available": mse_col is not None,
            "term_max_marks": {}, "warnings": ["Normalized student-course table; missing grading rules were not inferred."]}
    return Workbook(students, terms, subjects, mark_rows, cumulative, info)

def _read_wide_mse(raw):
    """Read an explicitly column-labeled, one-row-per-student MSE marks register."""
    def value_text(value):
        if pd.isna(value): return ""
        if isinstance(value, float) and value.is_integer(): value = int(value)
        return re.sub(r"\s+", " ", str(value)).strip()

    def identity_columns(row):
        columns = {}
        for col, value in enumerate(row):
            key = _norm(value)
            if key in {"student unique no", "student unique number", "student id",
                       "enrollment no", "enrollment number", "enrolment no",
                       "roll no", "roll number"}:
                columns.setdefault("enroll", col)
            elif key in {"student name", "name"}:
                columns.setdefault("name", col)
            elif key in {"student seat no", "seat no", "seat number"}:
                columns.setdefault("seat", col)
        return columns

    identity_row = next(
        ((i, identity_columns(raw.iloc[i].tolist()))
         for i in range(min(40, len(raw)))
         if {"enroll", "name"} <= set(identity_columns(raw.iloc[i].tolist()))),
        None,
    )
    if identity_row is None:
        raise UnsupportedLayout("MSE register needs clearly labeled Student ID/Unique No and Student Name columns.")
    identity_header, ids = identity_row

    def is_mse_marks_header(value):
        text = _norm(value)
        raw_text = str(value).strip().casefold()
        if re.fullmatch(r"marks\s*\(\s*30\s*\)", raw_text): return True
        if re.fullmatch(r"marks\s+out\s+of\s+30", raw_text): return True
        return text in {"mse", "mse marks", "mse mark", "mid semester marks", "mid semester exam marks"}

    nearby_headers = range(max(0, identity_header - 3), min(len(raw), identity_header + 4))
    marks_header = max(
        ((sum(is_mse_marks_header(value) for value in raw.iloc[row]), row) for row in nearby_headers),
        default=(0, identity_header),
    )[1]
    header = max(identity_header, marks_header)
    mark_columns = [col for col, value in enumerate(raw.iloc[marks_header]) if is_mse_marks_header(value)]
    if not mark_columns:
        raise UnsupportedLayout("No explicit MSE/MSE Marks or Marks (30) columns were found beside the student header.")

    subject_columns = {}
    ignored = re.compile(r"\b(?:int|ext|internal|external)?\s*passing\b|\bmarks?\b|\bmse\b|\bterm\b|\bsemester\b", re.I)
    for col in mark_columns:
        subject = ""
        for row in range(marks_header - 1, max(-1, marks_header - 7), -1):
            candidate = value_text(raw.iat[row, col])
            if candidate and not ignored.search(candidate):
                subject = candidate
                break
        if not subject:
            raise UnsupportedLayout(
                f"Could not identify the subject heading for MSE marks in column {col + 1}; "
                "each MSE marks column must have a subject name above it."
            )
        subject_columns[col] = subject
    duplicated = pd.Series(list(subject_columns.values())).duplicated(keep=False)
    if duplicated.any():
        seen = {}
        for col, subject in subject_columns.items():
            seen[subject] = seen.get(subject, 0) + 1
            if seen[subject] > 1:
                subject_columns[col] = f"{subject} (column {col + 1})"

    students, mark_rows = [], []
    seen_enrollments = set()
    for row_index in range(header + 1, len(raw)):
        row = raw.iloc[row_index]
        enroll = value_text(row.iloc[ids["enroll"]]).upper()
        name = value_text(row.iloc[ids["name"]])
        marks = [row.iloc[col] for col in mark_columns]
        has_marks = any(not _mse_status(value) == "Unreported" for value in marks)
        if not enroll and not name:
            if has_marks:
                raise UnsupportedLayout(f"Row {row_index + 1} has MSE marks but no student ID or name.")
            continue
        if not enroll or not name:
            raise UnsupportedLayout(f"Row {row_index + 1} has an incomplete student ID/name; no marks were assigned.")
        if enroll in seen_enrollments:
            raise UnsupportedLayout(f"Student ID {enroll} occurs more than once; refusing to combine or misassign marks.")
        seen_enrollments.add(enroll)
        student = {"enroll": enroll, "name": name, "prefix": re.match(r"[A-Z]+", enroll).group(0)
                   if re.match(r"[A-Z]+", enroll) else ""}
        if "seat" in ids:
            student["seat"] = value_text(row.iloc[ids["seat"]])
        students.append(student)
        for col, subject in subject_columns.items():
            raw_mark = row.iloc[col]
            status = _mse_status(raw_mark)
            number = pd.to_numeric(raw_mark, errors="coerce")
            mark_rows.append({"enroll": enroll, "term": "MSE", "code": f"MSE-C{col + 1}",
                              "total": number, "mse": number, "mse_status": status,
                              "grade": "", "status": "AB" if status == "AB" else "Unreported"})
    if not students:
        raise UnsupportedLayout("No student rows were found below the MSE register headers.")

    student_frame = pd.DataFrame(students)
    student_frame["batch"] = float("nan")
    marks = pd.DataFrame(mark_rows)
    for field in ("ise", "internal", "int", "ese", "poe", "other", "credits", "earned", "gpoints"):
        marks[field] = float("nan")
    marks = marks[["enroll", "term", "code"] + ALLF + ["mse_status"]]
    subjects = pd.DataFrame([
        {"term": "MSE", "term_n": 1, "abbr": subject, "label": subject, "name": subject,
         "code": f"MSE-C{col + 1}", "dept": "", "type": "Other", "part": "Theory",
         "int_pass": None, "ext_pass": None, "enrolled": len(students), "credits": float("nan"),
         "max_total": 30.0, "pass_rule": "MSE marks out of 30; 12 or above passes",
         "has_mse": True, "max_ise": float("nan"), "max_mse": 30.0, "max_ese": float("nan"),
         "max_int": float("nan"), "max_oe_poe": float("nan"), "share": 1.0, "is_common": True}
        for col, subject in subject_columns.items()
    ])
    terms = pd.DataFrame([
        {"enroll": student.enroll, "term": "MSE", "term_n": 1,
         "sgpa": float("nan"), "pct": float("nan"), "cr_reg": float("nan"),
         "cr_earned": float("nan"), "grade_total": float("nan")}
        for student in student_frame.itertuples()
    ])
    return Workbook(student_frame, terms, subjects, marks,
                    pd.DataFrame({"enroll": student_frame.enroll}),
                    {"students": len(students), "terms": ["MSE"], "subjects": len(subjects),
                     "rows": len(raw), "mse_available": True, "term_max_marks": {},
                     "warnings": ["Explicit wide MSE register; each source mark remains linked to its student ID and subject column."]})

def read_workbook(src, sheet=0, include_mse_rows=False) -> Workbook:
    raw = _load(src, sheet)
    try:
        h, term_row, srow, sub = _layout(raw)
    except UnsupportedLayout as hierarchical_error:
        from .universal import read_auto         # local import avoids a circular import
        errs = [str(hierarchical_error)]
        tries = [lambda: _read_flat(raw, include_mse_rows)] + ([lambda: _read_wide_mse(raw), lambda: read_auto(raw)] if include_mse_rows else [])
        for attempt in tries:
            try: return attempt()
            except UnsupportedLayout as ex: errs.append(str(ex))
        raise UnsupportedLayout(" ".join(errs))
    body = raw.iloc[sub + 1:].reset_index(drop=True)
    # ... the rest of your existing function continues unchanged from here
    ce, cn, cs = (_col(raw.iloc[h], p, k) for p, k in ((r"unique|enrol", None), (r"\bname\b", r"mother|father"), (r"seat", None)))
    ce, cn = 1 if ce is None else ce, 2 if cn is None else cn
    body = body[body[ce].notna() & body[cn].notna()].reset_index(drop=True)
    if body.empty: raise UnsupportedLayout("Header found but no student rows below it.")
    enroll = body[ce].astype(str).str.strip().str.upper()
    stu = pd.DataFrame({"enroll": enroll, "name": body[cn].astype(str).str.strip().str.replace(r"\s+", " ", regex=True)})
    if cs is not None: stu["seat"] = pd.to_numeric(body[cs], errors="coerce").map(lambda v: str(int(v)) if pd.notna(v) else "")
    stu["prefix"] = stu.enroll.str.extract(r"^([A-Z]+)")[0]
    stu["batch"] = 2000 + pd.to_numeric(stu.enroll.str.extract(r"^[A-Z]+(\d{2})")[0], errors="coerce")

    top = raw.iloc[term_row]
    tcols = [(c, re.sub(r"^term\s*-?\s*", "", str(v), flags=re.I).strip()) for c, v in top.items()
             if re.match(r"^term\s*-?\s*\S+", str(v), re.I)]
    if not tcols: raise UnsupportedLayout("No 'Term - X' markers found in the header row.")
    ends = [c for c, _ in tcols[1:]] + [raw.shape[1]]
    fcols = {FINAL[_norm(v)]: c for c, v in top.items() if _norm(v) in FINAL and c > tcols[-1][0] + 5}

    terms, subjects, marks, warnings, term_max = [], [], [], [], {}
    for n, ((c0, label), c1) in enumerate(zip(tcols, ends), 1):
        starts = [c for c in range(c0, c1) if re.search(SUBJ_RX, str(raw.iat[srow, c]))]
        if not starts: starts = [c for c in range(c0, c1) if pd.notna(raw.iat[srow, c]) and str(raw.iat[srow, c]).strip()]
        s0 = next((c for c in range(starts[-1] + 1 if starts else c0, c1) if _norm(top[c]) == "total"), None)
        sts = [c for c in range(c0, s0) if _norm(raw.iat[sub, c]) == "st"] if s0 else []
        if sts:                                   # structural: each block ends at its 'St.' column
         sx = [c0] + [e + 1 for e in sts[:-1]]
         if all(raw.iloc[srow, a:b].notna().any() for a, b in zip(sx, sx[1:] + [sts[-1] + 1])): starts = sx
        if not starts or s0 is None: raise UnsupportedLayout(f"Term {label}: could not locate subject blocks or the term summary.")
        term_max[label] = _paren(top[s0])
        for j, c in enumerate(starts):
            end = starts[j + 1] if j + 1 < len(starts) else s0
            head = next((str(v) for v in raw.iloc[srow, c:end] if pd.notna(v) and str(v).strip()), "")
            m = re.match(r"^\s*(.+?)\s*\(\s*([A-Za-z0-9][A-Za-z0-9 ]{4,}[A-Za-z0-9](?:-\d+)?(?:\([^()]*\))?)\s*\)\s*$", head)
            abbr, code = (m.group(1).strip(), re.sub(r"\s+", "", m.group(2))) if m else (head.strip(),) * 2
            labels = [_norm(raw.iat[sub, k]) for k in range(c, end)]
            names, extra = _fields(labels, f"Term {label}, {abbr}")
            if extra: warnings.append(f"Term {label}, {abbr}: ignored extra columns {extra}")
            meta = _meta(raw, h, term_row, srow, sub, c)
            b = body.iloc[:, [c + i for i in names]].copy(); b.columns = list(names.values())
            mse_status = b["mse"].map(_mse_status) if "mse" in b else pd.Series("Unreported", index=b.index)
            g = b["grade"]
            keep = g.notna() & ~g.astype(str).str.strip().isin(["", "-", "nan", "None"])
            if include_mse_rows: keep |= mse_status.isin(["Marked", "AB"])
            b = b[keep].copy()
            mse_status = mse_status.loc[b.index]
            b["grade"], b["status"] = b.grade.astype(str).str.strip(), b.status.astype(str).str.strip()
            for f in NUMF: b[f] = _num(b[f]) if f in b else float("nan")
            b = b[ALLF]
            b["mse_status"] = mse_status
            b.insert(0, "enroll", enroll[b.index].values); b.insert(1, "term", label); b.insert(2, "code", code)
            marks.append(b)
            dept, typ, part = parse_code(code)
            row = {"term": label, "term_n": n, "abbr": abbr, "code": code, "dept": dept, "type": typ, "part": part,
                   "int_pass": meta.get("int_pass"), "ext_pass": meta.get("ext_pass"), "enrolled": len(b),
                   "credits": b.credits.median() if b.credits.notna().any() else meta.get("credits_declared", float("nan")),
                   "max_total": meta.get("max_total", float("nan")), "pass_rule": meta.get("pass_rule", ""),
                   "has_mse": "mse" in names.values()}
            for lab, col in MAXCOL.items():
                k = next((c + i for i, f in names.items() if f == COMPONENTS[lab]), None)
                row[col] = _paren(raw.iat[sub, k]) if k is not None else float("nan")
            subjects.append(row)
        sm = {}
        for k in range(s0, c1):
            lab = _norm(top[k])
            if lab in SUMMARY and SUMMARY[lab] not in sm: sm[SUMMARY[lab]] = _num(body[k]).values
        t = pd.DataFrame({"enroll": enroll.values, "term": label, "term_n": n, **sm})
        terms.append(t)

    terms = pd.concat(terms, ignore_index=True)
    terms = terms[terms[["sgpa", "grade_total", "cr_reg"]].notna().any(axis=1)] if "sgpa" in terms else terms
    subj = pd.DataFrame(subjects); N = len(stu)
    subj["share"] = (subj.enrolled / N).round(3); subj["is_common"] = subj.share >= 0.9
    dup = subj.groupby(["term", "abbr"]).code.transform("nunique") > 1
    subj["label"] = subj.abbr.where(~dup, subj.abbr + " (" + subj.part + ")")
    subj["name"] = subj.label
    cum = pd.DataFrame({"enroll": enroll.values, **{f: (body[c].astype(str).str.strip().values if f in ("grade", "status") else _num(body[c]).values)
                                                    for f, c in fcols.items()}})
    info = {"students": N, "terms": [l for _, l in tcols], "subjects": len(subj), "rows": len(raw),
            "mse_available": bool(subj.has_mse.any()), "term_max_marks": term_max, "warnings": warnings}
    return Workbook(stu, terms, subj, pd.concat(marks, ignore_index=True), cum, info)
