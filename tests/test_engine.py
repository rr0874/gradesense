"""Engine regression tests.  Run from the project folder:  python -m tests.test_engine"""
import tempfile, warnings
from pathlib import Path
warnings.filterwarnings("ignore")
from backend import Dataset
from backend.engine import Engine

ds = Dataset.from_workbook("data/AIML_4.xlsx", "test")
en = Engine(ds, log_path=tempfile.mkstemp(suffix=".csv")[1])
def ask(q): a = en.ask(q); return a, (0 if a.table is None else len(a.table))
passed = 0
def ok(cond, label):
    global passed
    assert cond, "FAILED: " + label
    passed += 1

# overall results agree with the sheet's own status column
ok(ask("ATKT students")[1] == 14, "14 ATKT students")
ok("Pass 133" in ask("class summary")[0].text, "133 students with final status Pass")
ok(ask("who failed")[1] + ask("who passed")[1] == 153, "failed + cleared = all 153 students")
ok(ask("students with no backlogs")[1] == ask("who passed")[1], "'no backlogs' equals 'who passed'")

# combined conditions are real set logic
ca, os_, both, either = ask("who failed in CA"), ask("who failed in OS"), ask("failed in CA and OS"), ask("failed in CA or OS")
ok(ca[1] == 9 and os_[1] == 0, "9 fail CA, nobody fails OS")
ok(both[1] == 0 and either[1] == ca[1], "AND is the intersection, OR is the union")
ok(ask("failed CA but passed OS")[1] == 9, "failed CA but passed OS")
ok(ask("did Sukhada Done pass OS?")[0].text.startswith("**Yes**"), "single-student yes/no")
ca_code = en.subj.query("term == 'IV' and label == 'CA'").code.iloc[0]
ok(ask("absent in CA")[1] == int(((ds.wb.marks.status == "AB") & (ds.wb.marks.code == ca_code)).sum()), "absent count matches raw marks")

# filters honour their numbers
t = ask("students with O grade in WT"); ok(t[1] == 19 and all(v.endswith("O") for v in t[0].table["WT"]), "19 students with O in WT")
t = ask("scored above 80 in DCN"); ok(t[1] > 0 and all(float(v.split(" · ")[0]) > 80 for v in t[0].table["DCN"]), "marks filter > 80")
t = ask("students with SGPA above 9"); ok(t[1] == 12 and (t[0].table.iloc[:, 2] > 9).all(), "SGPA filter > 9")
t = ask("top 5 students"); ok(t[1] == 5 and t[0].table["SGPA"].is_monotonic_decreasing, "top 5 sorted by SGPA")
t = ask("students with 2 backlogs"); ok(t[1] == 5 and (t[0].table["Backlogs"] == 2).all(), "exactly 2 backlogs")

# terms, names, comparisons, errors
ok("Term II — 138 students" in ask("summary of sem 2")[0].text, "term override by words")
ok("Term III" in ask("who improved from sem 3 to sem 4")[0].text, "term change question")
ok("Sukhada" in ask("result of EN23209642")[0].text, "lookup by enrollment number")
ok("Sukhada" in ask("analysis of Sukhda Don")[0].text, "typo-tolerant name match")
ok("couldn't find a student" in ask("analysis of Zzqx")[0].text, "unknown student is reported, not guessed")
ok("isn't in it" in ask("summary of sem 7")[0].text, "missing term is reported")
c1, c2 = ask("compare Sukhada Done and Prathmesh Patil")[0], ask("compare Prathmesh Patil and Sukhada Done")[0]
ok("higher SGPA" in c1.text and "higher SGPA" in c2.text and c1.text.split(" has")[0] == c2.text.split(" has")[0], "comparison is symmetric")
ok(ask("compare CA and OS")[1] == 2, "compare two subjects")
ok(len(ask("pass percentage of CA and OS")[0].table) == 2, "pass % for two subjects")
ok(ask("failed in AC sem 2")[0].text.startswith("Nobody failed AC"), "Term II duplicate abbreviation (theory+practical) handled")
ok(ask("xyzzy foo")[0].text.startswith("I didn't understand"), "gibberish gets help, not a guess")
ok("Term IV — 153 students" in en.ask_request("class summary").text,
   "uploaded request routes supported class analysis")
ok("can't safely produce that analysis" in en.ask_request("predict next semester results").text,
   "unsupported predictive request is declined")
ok("couldn't confidently match" in en.ask_request("make a very special analysis").text,
   "unclear uploaded request asks for clarification")
request_log = Path(en.log_path).read_text()
en.ask_request("compare")
ok(Path(en.log_path).read_text() == request_log, "uploaded request text is not written to the unmatched-query log")

# nothing crashes on a broad sweep
for q in ["help", "toppers in DCN", "grade distribution in ES", "average marks in WT", "list of subjects", "credits of Operating System", "sgpa trend",
          "sgpa trend of Sukhada Done", "absent students", "who failed in sem 1", "top 3 in sem 3", "students with CGPA above 8", "hardest subject"]:
    a, _ = ask(q); ok(len(a.text) > 5, f"answers: {q}")

# ---------------------------------------------------------------- second layout: single-term sheet with max marks / pass rules printed in the header
import pandas as pd
F2 = "data/Sample_Student_Results_Gen_AI.xlsx"
raw = pd.read_excel(F2, header=None); body = raw.iloc[5:]
gr = [c for c in range(raw.shape[1]) if str(raw.iat[4, c]).strip() == "Gr."]
name_of = lambda c: next(str(raw.iat[1, k]).split(" (")[0].strip() for k in range(c, -1, -1) if pd.notna(raw.iat[1, k]))
truth_fail = {name_of(c): int((body[c + 4] != "Pass").sum()) for c in gr}                 # computed straight from the cells
truth_O = {name_of(c): int((body[c] == "O").sum()) for c in gr}
e2 = Engine(Dataset.from_workbook(F2, "gen-ai"), log_path=tempfile.mkstemp(suffix=".csv")[1])
def ask2(q): a = e2.ask(q); return a, (0 if a.table is None else len(a.table))
ok(e2.terms == ["II"] and len(e2.S) == 50, "second layout: 50 students, one term")
for sub, n in truth_fail.items(): ok(ask2(f"who failed in {sub}")[1] == n, f"second layout: fails in {sub} = {n}")
for sub in ("Gen-AI-LAB", "IOT"): ok(ask2(f"students with O grade in {sub}")[1] == truth_O[sub], f"second layout: O grades in {sub}")
any_fail = int(pd.concat([body[c + 4] != "Pass" for c in gr], axis=1).any(axis=1).sum())
ok(ask2("who failed")[1] == any_fail, "second layout: students with any failed subject")
ok(ask2("students with SGPA above 8")[1] == int((pd.to_numeric(body[raw.shape[1] - 2]) > 8).sum()), "second layout: SGPA filter")
ok(ask2("failed gen ai lab")[0].text.startswith("Nobody failed Gen-AI-LAB") == (truth_fail["Gen-AI-LAB"] == 0), "subject typed with spaces matches hyphenated label")
t = ask2("scored above 70% in Gen AI"); ok(t[1] > 0 and all(float(v.split(" · ")[0]) > 70 for v in t[0].table["Gen AI"]), "percent filter uses the max marks printed in the sheet")
ok(ask2("list of subjects")[0].table["Max marks"].iloc[0] == "100", "max marks come from the sheet header")
for q in ["class summary", "top 5 students", "pass percentage", "analysis of Tejas Mohan Joshi", "compare Gen AI and RM", "average marks in PR", "grade distribution in NLP", "ATKT students", "who improved"]:
    ok(len(ask2(q)[0].text) > 5, f"second layout answers: {q}")

# Repeated source student rows are rejected; identical cumulative rows are safe to collapse.
duplicate_ds = Dataset.from_workbook("data/AIML_4.xlsx", "duplicate-cumulative")
duplicate_ds.wb.cumulative = pd.concat(
    [duplicate_ds.wb.cumulative, duplicate_ds.wb.cumulative.iloc[[0]]],
    ignore_index=True,
)
duplicate_engine = Engine(duplicate_ds, log_path=tempfile.mkstemp(suffix=".csv")[1])
ok(len(duplicate_engine.base(duplicate_engine.last)) == len(duplicate_ds.wb.students),
   "identical duplicate cumulative rows do not duplicate semester results")
duplicate_student_ds = Dataset.from_workbook("data/AIML_4.xlsx", "duplicate-students")
duplicate_student_ds.wb.students = pd.concat(
    [duplicate_student_ds.wb.students, duplicate_student_ds.wb.students.iloc[[0]]],
    ignore_index=True,
)
try:
    Engine(duplicate_student_ds, log_path=tempfile.mkstemp(suffix=".csv")[1])
except ValueError as ex:
    ok("duplicate enrollment number(s)" in str(ex) and "Remove the repeated student rows" in str(ex),
       "duplicate student rows are rejected with a source-workbook correction message")
else:
    raise AssertionError("FAILED: duplicate student rows must not be silently analyzed")
conflict_ds = Dataset.from_workbook("data/AIML_4.xlsx", "conflicting-cumulative")
conflict_row = conflict_ds.wb.cumulative.iloc[[0]].copy()
conflict_row["status"] = "Conflicting status"
conflict_ds.wb.cumulative = pd.concat([conflict_ds.wb.cumulative, conflict_row], ignore_index=True)
try:
    Engine(conflict_ds, log_path=tempfile.mkstemp(suffix=".csv")[1])
except ValueError as ex:
    ok("conflicting rows for enrollment number" in str(ex),
       "conflicting cumulative rows are rejected with an actionable error")
else:
    raise AssertionError("FAILED: conflicting cumulative rows must not be silently selected")

print(f"engine tests incl. second layout: {passed} checks passed")
