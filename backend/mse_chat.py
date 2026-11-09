"""Deterministic chat answers over validated MSE analysis records."""
import re

import pandas as pd

from .engine import Answer


def _matches(text, value):
    return re.search(rf"(?<!\w){re.escape(str(value).casefold())}(?!\w)", text.casefold()) is not None


def answer_division_mse(analysis, question):
    records = analysis.records.copy()
    text = question.casefold().strip()
    if not text:
        return Answer("Ask about MSE overall summaries, passed/failed/absent results, a subject, class, division, or student.")

    selected = records
    class_aliases = {
        "first year": ("first year", "fy"),
        "second year": ("second year", "sy"),
        "third year": ("third year", "ty"),
        "fourth year": ("fourth year", "fourth"),
        "final year": ("final year", "final"),
    }
    for class_name, aliases in class_aliases.items():
        if any(_matches(text, alias) for alias in aliases):
            selected = selected[selected["Class"].eq(class_name)]
            break
    division = re.search(r"\b(?:division|div)\s*[-:]?\s*([a-z])\b", text)
    if division:
        selected = selected[selected["Division"].str.casefold().eq(division.group(1).casefold())]

    subject = next(
        (course for course in sorted(records["Course"].dropna().unique(), key=len, reverse=True)
         if _matches(text, course)),
        None,
    )
    if subject:
        selected = selected[selected["Course"].eq(subject)]

    summary_request = bool(re.search(r"\b(summary|overall|count|counts|how many|statistics|performance)\b", text))
    status_words = {
        value for value, pattern in (
            ("Absent", r"\b(absent|absentees?|ab)\b"),
            ("Fail", r"\b(fail(?:ed|s|ing)?|failure|backlogs?)\b"),
            ("Pass", r"\b(pass(?:ed|es)?|success(?:ful)?)\b"),
        ) if re.search(pattern, text)
    }
    status = next(iter(status_words)) if len(status_words) == 1 else None
    if summary_request and len(status_words) > 1:
        status = None
    if status:
        selected = selected[selected["Outcome"].eq(status)]

    roll_match = re.search(r"\b[a-z]*\d+[a-z0-9]*\b", text)
    if roll_match:
        selected = selected[selected["Roll No"].str.casefold().eq(roll_match.group(0).casefold())]
    else:
        name_tokens = [
            token for token in re.findall(r"[a-z]+", text)
            if len(token) > 2 and token not in {
                "who", "what", "show", "list", "give", "find", "tell", "about", "student",
                "students", "result", "results", "analysis", "summary", "overall", "class",
                "division", "div", "passed", "pass", "failed", "fail", "absent", "absentees",
                "count", "many", "number", "marks", "mark", "mse", "the", "for", "with",
                "from", "and", "how", "did", "perform", "please", "can", "you",
            }
        ]
        if name_tokens:
            names = selected["Student Name"].str.casefold()
            matched = pd.Series(True, index=selected.index)
            for token in name_tokens:
                matched &= names.str.contains(re.escape(token), regex=True)
            if matched.any():
                selected = selected[matched]

    if selected.empty:
        return Answer("No matching MSE results were found. Check the student name/roll number, subject, class, or division.")

    if summary_request and not roll_match:
        frame = selected.groupby(
            ["Class", "Division", "Course", "Outcome"], dropna=False
        ).size().unstack(fill_value=0).reset_index()
        for outcome in ("Pass", "Fail", "Absent", "Unreported"):
            if outcome not in frame:
                frame[outcome] = 0
        frame = frame[["Class", "Division", "Course", "Pass", "Fail", "Absent", "Unreported"]]
        return Answer(f"MSE outcome counts for {len(selected)} subject results.", frame)

    columns = ["Class", "Division", "Roll No", "Student Name", "Course", "MSE Marks", "Outcome"]
    frame = selected[columns].sort_values(
        ["Class", "Division", "Course", "Roll No"], kind="stable"
    ).reset_index(drop=True)
    if status:
        return Answer(f"Found {len(frame)} {status.lower()} MSE result(s).", frame)
    return Answer(f"Found {len(frame)} matching MSE result(s).", frame)


def answer_single_mse(analysis, question):
    text = question.casefold().strip()
    if not text:
        return Answer("Ask about MSE overall outcomes, passed/failed/absent students, a subject, or a student.")

    students = analysis.student_results.copy()
    subjects = analysis.subject_summary.copy()
    subject = next(
        (value for value in sorted(
            set(subjects["Subject"].dropna().astype(str)) |
            set(subjects["Course code"].dropna().astype(str)),
            key=len, reverse=True,
        ) if _matches(text, value)),
        None,
    )
    if subject:
        mask = subjects["Subject"].astype(str).str.casefold().eq(subject.casefold())
        if "Course code" in subjects:
            mask |= subjects["Course code"].astype(str).str.casefold().eq(subject.casefold())
        row = subjects[mask]
        if len(row):
            return Answer(f"MSE summary for {subject}.", row.reset_index(drop=True))

    status = None
    if re.search(r"\b(absent|absentees?)\b", text):
        status = "Absent"
    elif re.search(r"\b(fail(?:ed|s|ing)?|failure|backlogs?)\b", text):
        status = "Fail"
    elif re.search(r"\b(pass(?:ed|es)?|success(?:ful)?)\b", text):
        status = "Pass"
    if status:
        result_column = "Overall MSE result"
        frame = students[students[result_column].eq(status)].reset_index(drop=True)
        return Answer(f"Found {len(frame)} students with overall MSE result {status}.", frame)

    roll_match = re.search(r"\b[a-z]*\d+[a-z0-9]*\b", text)
    tokens = [
        token for token in re.findall(r"[a-z]+", text)
        if len(token) > 2 and token not in {
            "who", "what", "show", "list", "give", "find", "tell", "about", "student",
            "students", "result", "results", "analysis", "summary", "overall", "class",
            "passed", "pass", "failed", "fail", "absent", "absentees", "count", "many",
            "number", "marks", "mark", "mse", "the", "for", "with", "from", "and",
            "how", "did", "perform", "please", "can", "you",
        }
    ]
    if roll_match or tokens:
        mask = pd.Series(True, index=students.index)
        if roll_match:
            mask &= students["Enrollment"].astype(str).str.casefold().eq(roll_match.group(0).casefold())
        if tokens:
            names = students["Student"].astype(str).str.casefold()
            for token in tokens:
                names_match = names.str.contains(re.escape(token), regex=True)
                mask &= names_match
        frame = students[mask].reset_index(drop=True)
        if len(frame):
            return Answer("Overall MSE outcome for the matching student(s).", frame)

    if re.search(r"\b(summary|overall|count|counts|how many|statistics|performance)\b", text):
        return Answer(
            f"Overall MSE results for {analysis.totals['students']} students.",
            pd.DataFrame([analysis.totals]),
        )
    return Answer("No matching MSE results were found. Ask for an overall summary, a subject outcome, or a student by name/enrollment.")


def answer_mse_request(analysis, question, answer_function):
    """Only forward uploaded requests that match an existing deterministic MSE query."""
    text = question.casefold().strip()
    if not text:
        return Answer("No request text was provided. Review the recognized text and try again.")
    if re.search(
        r"\b(predict\w*|forecast\w*|correlat\w*|regression|cluster\w*|"
        r"trend\w*|compar\w*|rank\w*|topper\w*|recommend\w*|causal\w*)\b",
        text,
    ):
        return Answer(
            "I can't safely produce that analysis from the current MSE workbook. "
            "This assistant supports MSE summaries, pass/fail/absent outcomes, "
            "subject summaries, and student lookups only."
        )

    if hasattr(analysis, "records"):
        records = analysis.records
        subjects = records["Course"].dropna().astype(str).unique()
        names = records["Student Name"].dropna().astype(str).unique()
        identifiers = records["Roll No"].dropna().astype(str).unique()
    else:
        subjects_frame = analysis.subject_summary
        subjects = set(subjects_frame["Subject"].dropna().astype(str))
        if "Course code" in subjects_frame:
            subjects.update(subjects_frame["Course code"].dropna().astype(str))
        names = analysis.student_results["Student"].dropna().astype(str).unique()
        identifiers = analysis.student_results["Enrollment"].dropna().astype(str).unique()

    summary = re.search(r"\b(summary|overall|count|counts|how many|statistics)\b", text)
    outcome = re.search(r"\b(pass(?:ed|es)?|fail(?:ed|s|ing)?|failure|backlogs?|absent|absentees?)\b", text)
    mentions_subject = any(_matches(text, value) for value in subjects)
    mentions_identifier = any(_matches(text, value) for value in identifiers)
    mentions_student = bool(re.search(r"\b(student|students|who|result|results)\b", text)) and any(
        tokens and all(_matches(text, token) for token in tokens)
        for name in names
        for tokens in [[token for token in re.findall(r"[a-z0-9]+", name.casefold()) if len(token) > 1]]
    )
    if not (summary or outcome or mentions_subject or mentions_identifier or mentions_student):
        return Answer(
            "I couldn't confidently match this request to an available MSE analysis. "
            "Try an overall summary, pass/fail/absent results, a subject summary, "
            "or a student lookup by name or roll number."
        )
    return answer_function(analysis, question)
