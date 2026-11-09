"""Stage 4 - analysis engine v2 on the universal backend. No API, no LLM:
a rule router + TF-IDF/LogReg intent classifier + regex entities + vetted pandas analytics.
Works for any term / subject set the backend ingested; answers come only from the sheet."""
import csv, difflib, re
from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace
import numpy as np, pandas as pd
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.linear_model import LogisticRegression
from sklearn.pipeline import make_pipeline
from .catalog import all_catalogs, apply_catalog

GRADES = ["O", "A+", "A", "B+", "B", "C", "P", "F", "AB"]
ROMAN = ["I", "II", "III", "IV", "V", "VI", "VII", "VIII"]
ORD = {"first": "I", "second": "II", "third": "III", "fourth": "IV", "fifth": "V", "sixth": "VI", "seventh": "VII", "eighth": "VIII"}
TERM_RX = re.compile(r"\b(?:sem(?:ester)?|term)\s*[-:.]?\s*(viii|vii|vi|iv|v|iii|ii|i|[1-8])\b|\b(first|second|third|fourth|fifth|sixth|seventh|eighth)\s+(?:sem(?:ester)?|term)\b", re.I)
CMP_RX = re.compile(r"(above|over|more than|greater than|at least|not less than|below|under|less than|at most|not more than|exactly|equal to|>=|<=|==|>|<)\s*(\d+(?:\.\d+)?)\s*(%|percent)?", re.I)
CMP_OP = {"above": ">", "over": ">", "more than": ">", "greater than": ">", ">": ">", "at least": ">=", "not less than": ">=", ">=": ">=",
          "below": "<", "under": "<", "less than": "<", "<": "<", "at most": "<=", "not more than": "<=", "<=": "<=", "exactly": "==", "equal to": "==", "==": "=="}
OP_WORD = {">": "above", ">=": "at least", "<": "below", "<=": "at most", "==": "exactly"}
FAIL_RX = re.compile(r"\b(?:fail(?:ed|s|ing)?|did(?:n't|\s+not)\s+(?:pass|clear)|could(?:n't|\s+not)\s+(?:pass|clear)|not\s+(?:pass(?:ed)?|clear(?:ed)?)|unable\s+to\s+(?:pass|clear)|backlogs?\s+in)\b(?!\s*(?:percent|%|rate|list))", re.I)
PASS_RX = re.compile(r"\b(?:pass(?:ed|es)?|clear(?:ed|s)?|succeed(?:ed)?)\b(?!\s*(?:percent|%|rate|list|marks))", re.I)
ABS_RX = re.compile(r"\b(?:absent(?:ee|ees)?)\b", re.I)
GUARD = re.compile(r"\b(percent(?:age)?|rate|distribution|average|mean|median|top|toppers?|highest|lowest|credits?|scheme|details|compare|vs|versus|rank|summary|statistics)\b", re.I)
CHANGE = re.compile(r"\b(improv\w*|drop\w*|declin\w*|increas\w*|decreas\w*|progress\w*|rose|fell|worse\w*|better|regress\w*)\b", re.I)
UNSUPPORTED_REQUEST = re.compile(
    r"\b(predict\w*|forecast\w*|correlat\w*|regression|cluster\w*|demograph\w*|"
    r"gender|age group|causal\w*|custom formula|combine datasets?)\b",
    re.I,
)

TEMPLATES = {
 "fail_list": ["who failed", "list of failed students", "show fail list", "which students are failing", "students with backlog", "students who did not pass", "fail students", "failures this semester", "who has backlogs"],
 "pass_list": ["who passed", "list of passed students", "show pass list", "students who cleared all subjects", "pass students", "who passed everything", "list students who passed"],
 "grade_list": ["students with xgrade grade", "who got xgrade", "list students having xgrade grade", "show xgrade grade students", "who scored xgrade grade", "students who got xgrade", "grade xgrade students", "xgrade grade holders"],
 "topper_list": ["top xnum students", "who are the toppers", "toppers in xsubj", "highest scorer in xsubj", "top xnum in xsubj", "rank list", "best performing students", "who scored highest in xsubj", "top scorers", "merit list", "top xnum students by sgpa", "top xnum students by cgpa", "top students by cgpa", "who has the highest cgpa", "highest cgpa students", "who is the topper"],
 "atkt_list": ["atkt students", "who has atkt", "list of atkt students", "students with atkt", "show atkt list", "atkt list", "how many atkt students"],
 "pass_pct": ["pass percentage of xsubj", "what is the pass percentage", "pass rate in xsubj", "subject wise pass percentage", "how many percent passed xsubj", "result percentage of xsubj", "pass percent", "success rate in xsubj", "which subject has lowest pass percentage", "hardest subject", "toughest subject"],
 "grade_dist": ["grade distribution in xsubj", "how many students got each grade", "grade wise count", "distribution of grades", "grade breakup of xsubj", "count of students per grade in xsubj", "show grade distribution"],
 "class_summary": ["class summary", "overall result", "summary of the semester", "overall performance of the class", "give me class analysis", "semester analysis", "result summary", "overall analysis", "how did the class perform", "average sgpa of class", "average cgpa of class", "class cgpa average", "how many students passed and failed", "total number of students", "summary of xterm", "class summary for xterm", "how many students are there"],
 "subject_info": ["credits of xsubj", "marking scheme of xsubj", "details of xsubj", "course code of xsubj", "what are the subjects in this semester", "list of subjects", "passing marks for xsubj", "tell me about xsubj", "teaching scheme", "how many credits is xsubj", "course details", "subjects in xterm", "which subjects are there"],
 "subject_stats": ["average marks in xsubj", "mean score of xsubj", "how did the class do in xsubj", "statistics of xsubj", "subject analysis of xsubj", "highest and lowest marks in xsubj", "average of xsubj", "class average in xsubj"],
 "student_report": ["analysis of xname", "show result of xname", "how did xname perform", "report card of xname", "xname marks", "student analysis xname", "performance of xname", "details of student xname", "xname result", "xname", "marks of xname in all subjects", "weak subjects of xname", "strong subjects of xname", "result of xname in xterm"],
 "sgpa_trend": ["sgpa trend of xname", "sgpa history of xname", "term wise sgpa", "sgpa of all semesters", "how has xname progressed", "improvement of xname over semesters", "cgpa trend", "semester wise sgpa of class", "average sgpa per term", "what is the sgpa of xname"],
 "absent_list": ["absent students", "list of absentees", "students marked ab", "who is absent", "students absent for exam", "who was absent"],
 "help": ["help", "hi", "hello", "what can you do", "how to use", "what can i ask", "commands", "examples", "hey"],
}
PREFIXES = ["", "please ", "can you ", "show me ", "tell me ", "i want to know "]
HELP = ("Ask me things like:\n- *class summary* / *summary of sem 3*\n- *who failed in OS?* / *students with O grade in WT*\n- *failed CA but passed OS* / *scored above 80 in DCN*\n"
        "- *top 5 students* / *toppers in CA* / *pass percentage* / *average marks in ES*\n- *analysis of <student name or enrollment no.>* / *compare <student> and <student>*\n"
        "- *students with SGPA above 9* / *students with more than 1 backlog* / *who improved from sem 3 to sem 4*\n- *ATKT students*, *absent in CAL*, *list of subjects*, *credits of <subject>*")

@dataclass
class Answer:
    text: str
    table: pd.DataFrame = None
    chart: pd.DataFrame = None

def _cmp(s, op, v): return {">": s > v, ">=": s >= v, "<": s < v, "<=": s <= v, "==": s == v}[op]
def _tab(df): return df if df is not None and len(df) else None


def _unique_enrollment_rows(frame, label, allow_identical_duplicates=False):
    duplicate_ids = frame.loc[frame.enroll.duplicated(keep=False), "enroll"].drop_duplicates()
    if duplicate_ids.empty:
        return frame.copy()

    examples = ", ".join(map(str, duplicate_ids.head(3)))
    unique_rows = frame.drop_duplicates()
    conflicting_ids = unique_rows.loc[
        unique_rows.enroll.duplicated(keep=False), "enroll"
    ].drop_duplicates()
    if not conflicting_ids.empty:
        conflicts = ", ".join(map(str, conflicting_ids.head(3)))
        raise ValueError(f"{label} contains conflicting rows for enrollment number(s): {conflicts}.")
    if not allow_identical_duplicates:
        raise ValueError(f"{label} contains duplicate enrollment number(s): {examples}. "
                         "Remove the repeated student rows from the source workbook and upload it again.")
    return unique_rows.drop_duplicates("enroll").copy()


class Engine:
    def __init__(self, ds, catalogs=None, log_path=None):
        self.ds, wb = ds, ds.wb
        self.log_path = log_path or (Path.cwd() / "unmatched_queries.csv")
        subj = wb.subjects.copy()
        for cat in (all_catalogs() if catalogs is None else catalogs).values(): subj, _ = apply_catalog(subj, cat)
        if "aliases" not in subj: subj["aliases"] = ""
        if "in_catalog" not in subj: subj["in_catalog"] = False
        self.subj = subj
        self.terms = list(wb.terms.drop_duplicates("term").sort_values("term_n").term)
        self.last = self.default = self.terms[-1]
        students = _unique_enrollment_rows(wb.students, "Student records")
        cumulative = _unique_enrollment_rows(
            wb.cumulative, "Cumulative records", allow_identical_duplicates=True
        )
        self.S = students.merge(cumulative, on="enroll", how="left", validate="one_to_one").set_index("enroll", drop=False)
        student_names = students[["enroll", "name"]]
        self.T = wb.terms.merge(student_names, on="enroll", validate="many_to_one")
        keep = ["term", "code", "label", "name", "abbr", "part", "type", "is_common"]
        self.M = (wb.marks.merge(subj[keep].rename(columns={"name": "subject"}), on=["term", "code"])
                  .merge(student_names, on="enroll", validate="many_to_one"))
        self.smap = {(r.term, r.code): r for r in subj.itertuples()}
        self._base = {}
        self._alias = {t: self._aliases(t) for t in self.terms}
        pref = sorted(set(self.S.prefix.dropna()), key=len, reverse=True)
        self.enroll_rx = re.compile(r"\b(?:" + "|".join(pref) + r")\s?\d{5,}\b", re.I) if pref else re.compile(r"$^")
        self.names = {en: set(re.findall(r"[a-z]+", n.lower())) for en, n in self.S.name.items()}
        self.vocab = set().union(*self.names.values())
        X, y = [], []
        for intent, ts in TEMPLATES.items():
            for t in ts:
                for p in PREFIXES: X.append(p + t); y.append(intent)
        self.clf = make_pipeline(TfidfVectorizer(ngram_range=(1, 2)), LogisticRegression(C=20, max_iter=3000)).fit(X, y)
        self.stop = {w for t in X for w in re.findall(r"[a-z]+", t)} | set(
            "the and for are was were has have that this from with not you your students student sem semester marks result results please compare between versus term who what which how".split())

    # ------------------------------------------------------------ helpers
    def _aliases(self, term):
        a = {}
        for r in self.subj[self.subj.term == term].itertuples():
            base = re.sub(r"\s*\(lab\)$", "", r.name, flags=re.I).lower()
            extra = [x for x in str(r.aliases or "").split("|") if x]
            names = {r.label.lower(), r.abbr.lower()}
            if r.part == "Practical": names |= {f"{base} lab", f"{base} practical", f"{base} (lab)"} | {f"{x} lab" for x in extra}
            else: names |= {base, r.name.lower()} | set(extra)
            for n in list(names):
                for v in {n, re.sub(r"[-_/]+", " ", n)}:
                    v = " ".join(v.split())
                    if v: a.setdefault(v, set()).add(r.code)
        return sorted(a.items(), key=lambda kv: -len(kv[0]))

    def disp(self, term, code):
        r = self.smap[(term, code)]
        return r.label if r.name == r.label else f"{r.name} [{r.label}]"
    def lab(self, term, code): return self.smap[(term, code)].label
    def _mt(self, term): return self.M[self.M.term == term]

    def _cgpa_available(self, term):
        return term == self.last and "cgpa" in self.S and self.S.cgpa.notna().any()

    def _requested_metric(self, e):
        if "cgpa" in e.low and self._cgpa_available(e.term):
            return "cgpa"
        return "sgpa"

    def base(self, term):
        if term not in self._base:
            t = self.T[self.T.term == term][["enroll", "name", "sgpa", "pct", "cr_reg", "cr_earned"]].copy()
            mt = self._mt(term); bad = mt[mt.status != "Pass"].groupby("enroll").agg(backlogs=("code", "count"), failed=("label", ", ".join))
            t = t.merge(bad, on="enroll", how="left"); t["backlogs"] = t.backlogs.fillna(0).astype(int); t["failed"] = t.failed.fillna("")
            t["status"] = t.enroll.map(self.S.status) if term == self.last and "status" in self.S else np.where(t.backlogs > 0, "Backlog", "Pass")
            self._base[term] = t.reset_index(drop=True)
        return self._base[term]

    def examples(self):
        b = self.base(self.last); nm = b.sort_values("sgpa", ascending=False).name.iloc[0] if len(b) else "a student"
        c = [r for r in self.subj[(self.subj.term == self.last) & self.subj.is_common].itertuples()]
        ex = ["Class summary", "Who failed?", "Top 5 students", "Pass percentage"]
        if c: ex.append(f"Who failed in {c[0].label}?")
        pair = next(((a, other) for i, a in enumerate(c) for other in c[i + 1:]
                     if a.code != other.code), None)
        if pair: ex.append(f"Failed {pair[0].label} but passed {pair[1].label}")
        ex.append(f"Analysis of {nm}")
        if len(self.terms) > 1: ex.append(f"Who improved from sem {self.terms[-2]} to sem {self.last}?")
        return list(dict.fromkeys(ex))

    # ------------------------------------------------------------ parsing
    def _grade_tokens(self, text, taken):
        out = []
        def ok(a, b): return not any(a < y and x < b for x, y, _ in taken + out)
        for rx, flags in [(r"(?<![\w+])(a\+|b\+|ab|o|a|b|c|p|f)\s+grades?\b", re.I), (r"\bgrades?\s+(?:of\s+|is\s+|=\s*)?(a\+|b\+|ab|o|a|b|c|p|f)(?![\w+])", re.I),
                          (r"(?<![\w+])(A\+|B\+|AB|O|A|B|C|P|F)(?![\w+])", 0), (r"(?<![\w+])([ab]\+)(?![\w+])", re.I)]:
            for m in re.finditer(rx, text, flags):
                if ok(*m.span(1)): out.append((m.start(1), m.end(1), m.group(1).upper()))
        if out:
            for m in re.finditer(r"\b(?:or|and)\s+(o|A\+|B\+|A|B|C|P|F)(?![\w+])", text):
                if ok(*m.span(1)): out.append((m.start(1), m.end(1), m.group(1).upper()))
        return out

    def parse(self, text, default=None):
        d = default if default in self.terms else self.default
        e = SimpleNamespace(text=text, term=d, explicit=False, bad_term=None, terms=[], groups=[], subjects=[], enroll=[], grades=[], cmps=[],
                            states=[], n=None, students=[], tokens=[], low=text.lower())
        low, spans = e.low, []
        free = lambda a, b: not any(a < y and x < b for x, y, _ in spans)
        for m in TERM_RX.finditer(low):
            raw = m.group(1) or ORD[m.group(2)]
            lab = ROMAN[int(raw) - 1] if raw.isdigit() else raw.upper()
            spans.append((m.start(), m.end(), "xterm")); e.terms.append(lab)
        good = [t for t in e.terms if t in self.terms]
        if e.terms and not good: e.bad_term = e.terms[0]
        if good: e.term, e.explicit = good[0], True
        order = [e.term] if e.explicit else [d] + [t for t in reversed(self.terms) if t != d]
        for t in order:                                        # subject mentions: chosen term first, else the latest term that has them
            hits = []
            for alias, codes in self._alias[t]:
                for m in re.finditer(r"(?<![\w+-])" + re.escape(alias) + r"(?![\w+-])", low):
                    if free(m.start(), m.end()) and not any(m.start() < y and x < m.end() for x, y, _ in [(h[0], h[1], 0) for h in hits]):
                        hits.append((m.start(), m.end(), sorted(codes)))
            if hits:
                e.term = t; e.groups = sorted(hits); spans += [(a, b, "xsubj") for a, b, _ in hits]
                e.subjects = list(dict.fromkeys(c for _, _, cs in hits for c in cs)); break
        for m in self.enroll_rx.finditer(low):
            if free(*m.span()): spans.append((m.start(), m.end(), "xname")); e.enroll.append(re.sub(r"\s", "", m.group(0)).upper())
        for m in CMP_RX.finditer(low):
            if free(*m.span()):
                op, v, pct = CMP_OP[m.group(1).lower()], float(m.group(2)), bool(m.group(3))
                spans.append((m.start(), m.end(), "xnum")); e.cmps.append((op, v, pct)); e.states.append((m.start(), "cmp", (op, v, pct)))
        for a, b, g in self._grade_tokens(text, spans):
            spans.append((a, b, "xgrade")); e.grades.append(g); e.states.append((a, "grade", g))
        for rx, kind in ((FAIL_RX, "fail"), (PASS_RX, "pass"), (ABS_RX, "absent")):
            for m in rx.finditer(low):
                if free(*m.span()): e.states.append((m.start(), kind, None)); spans.append((m.start(), m.end(), "xstate"))
        for m in re.finditer(r"\b\d{1,3}\b", low):
            if free(*m.span()):
                e.n = int(m.group(0)); spans.append((m.start(), m.end(), "xnum")); break
        e.grades = list(dict.fromkeys(e.grades))
        masked = low
        for a, b, tok in sorted(spans, reverse=True): masked = masked[:a] + f" {tok if tok != 'xstate' else low[a:b]} " + masked[b:]
        e.tokens = [w for w in re.findall(r"[a-z]+", masked) if len(w) >= 3 and w not in self.stop and not w.startswith("x")]
        e.students, used = self._match(e.tokens)
        for w in used: masked = re.sub(r"\b" + w + r"\b", " xname ", masked)
        if e.enroll:
            e.students = [x for x in e.enroll if x in self.names]
        return e, " ".join(masked.split())

    def _match(self, Q, fuzzy=False):
        if fuzzy: Q = [g[0] for g in (difflib.get_close_matches(w, self.vocab, 1, 0.82) for w in Q) if g]
        hit = {en: {w for w in Q if w in tk} for en, tk in self.names.items()}
        hit = {en: h for en, h in hit.items() if h}
        if not hit: return [], []
        top = max(len(h) for h in hit.values())
        return [en for en, h in hit.items() if len(h) == top], list(set().union(*hit.values()))

    def classify(self, masked):
        p = self.clf.predict_proba([masked])[0]; i = p.argmax()
        return self.clf.classes_[i], float(p[i])

    # ------------------------------------------------------------ routing
    def ask(self, text, default=None, log_unmatched=True):
        if not text.strip(): return Answer(HELP)
        e, masked = self.parse(text, default)
        if e.bad_term: return Answer(f"This sheet has terms {', '.join(self.terms)}. Term {e.bad_term} isn't in it.")
        if e.enroll and not e.students: return Answer(f"No student with enrollment no. {e.enroll[0]} in this sheet.")
        low = e.low
        if re.search(r"\b(compare|versus|vs\.?|difference between|comparison)", low):
            r = self.a_compare(e)
            if r: return r
        if e.cmps and not e.subjects:
            if re.search(r"backlog|failed subjects|arrears", low): return self.a_backlog_filter(e)
            if not e.students: return self.a_metric_filter(e)
        if re.search(r"\b(no|zero|without|nil)\b.*\bbacklogs?\b|\bbacklogs?\b.*\b(count|wise)\b|(\d+)\s+backlogs?", low) and not e.subjects: return self.a_backlog_filter(e)
        if e.subjects and e.states and not GUARD.search(low): return self.a_conditions(e)
        if CHANGE.search(low) and not e.students and len(self.terms) > 1 and re.search(r"sgpa|result|perform|student|who|marks|sem|term", low): return self.a_sgpa_change(e)
        if e.subjects and re.search(r"\b(average|mean|median|avg)\b", low): return self.a_subject_stats(e)
        intent, conf = self.classify(masked)
        if e.grades and intent in {"pass_list", "fail_list", "topper_list", "help"}: intent = "grade_list"
        if not e.students and e.tokens and (conf < 0.6 or intent in {"student_report", "sgpa_trend"}):
            e.students = self._match(e.tokens, fuzzy=True)[0]
            if e.students and intent not in {"student_report", "sgpa_trend"} and conf < 0.6: intent = "student_report"
        if conf < 0.25 and e.students: intent = "student_report"
        elif conf < 0.25 and intent != "grade_list":
            if log_unmatched: self._log(text, intent, conf)
            return Answer("I didn't understand that.\n\n" + HELP)
        if not e.students and not e.subjects and e.tokens and re.search(r"\b(analysis|analyse|analyze|report|result|marks|performance|perform|details|progress|sgpa)\b", e.low) \
                and intent in {"student_report", "class_summary", "sgpa_trend", "subject_info"}:
            return Answer(f"I couldn't find a student matching **{' '.join(e.tokens)}**. Check the spelling, or use the enrollment number.")
        if intent == "student_report" and not e.students: return Answer("Which student? Give a name or an enrollment number (e.g. *analysis of Sukhada Done*).")
        return getattr(self, "a_" + intent)(e)

    def ask_request(self, text, default=None):
        """Route OCR/user-authored requests conservatively through existing analyses."""
        if not text.strip():
            return Answer("No request text was provided. Review the recognized text and try again.")
        if UNSUPPORTED_REQUEST.search(text):
            return Answer(
                "I can't safely produce that analysis from the current result workbook. "
                "This assistant supports the existing student, subject, grade, pass/fail, "
                "summary, comparison, and semester-trend analyses only."
            )

        e, masked = self.parse(text, default)
        low = e.low
        supported_rule = (
            bool(re.search(r"\b(compare|versus|vs\.?|difference between|comparison)\b", low))
            or (bool(e.cmps) and not e.subjects)
            or (bool(e.subjects) and bool(e.states))
            or (bool(e.subjects) and bool(re.search(r"\b(average|mean|median|avg)\b", low)))
            or (bool(CHANGE.search(low)) and len(self.terms) > 1)
        )
        intent, confidence = self.classify(masked)
        if not e.students:
            residual = {
                word for word in e.tokens
                if word not in {"and", "or", "but", "also", "student", "students", "class", "results"}
                and not CHANGE.fullmatch(word)
            }
            if residual:
                return Answer(
                    "I couldn't confidently match this request to an available analysis. "
                    "Please rephrase it using a supported question, such as *class summary*, "
                    "*who failed in <subject>*, *average marks in <subject>*, or *top 5 students*."
                )
        if not supported_rule and confidence < 0.45:
            return Answer(
                "I couldn't confidently match this request to an available analysis. "
                "Please rephrase it using a supported question, such as *class summary*, "
                "*who failed in <subject>*, *average marks in <subject>*, or *top 5 students*."
            )
        return self.ask(text, default, log_unmatched=False)

    def _log(self, text, intent, conf):
        try:
            new = not Path(self.log_path).exists()
            with open(self.log_path, "a", newline="") as f:
                w = csv.writer(f)
                if new: w.writerow(["query", "predicted", "confidence"])
                w.writerow([text, intent, round(conf, 3)])
        except OSError: pass

    def _ctx(self, e): return f" (Term {e.term})" if len(self.terms) > 1 else ""

    # ------------------------------------------------------------ answers
    def a_help(self, e): return Answer(HELP)

    def a_fail_list(self, e, absent=False):
        b = self.base(e.term); mt = self._mt(e.term)
        if absent:
            ab = mt[mt.status == "AB"].groupby(["enroll", "name"]).label.agg(", ".join).reset_index()
            ab = ab.rename(columns={"name": "Name", "enroll": "Enrollment", "label": "Absent in"})
            return Answer(f"{len(ab)} student(s) absent in at least one subject{self._ctx(e)}." if len(ab) else f"Nobody was marked absent{self._ctx(e)}.", _tab(ab[["Name", "Enrollment", "Absent in"]]))
        d = b[b.backlogs > 0].sort_values(["backlogs", "name"], ascending=[False, True])
        t = d.rename(columns={"name": "Name", "enroll": "Enrollment", "failed": "Failed / absent in", "backlogs": "Count", "sgpa": "SGPA", "status": "Status"})[["Name", "Enrollment", "Failed / absent in", "Count", "SGPA", "Status"]]
        return Answer(f"{len(d)} of {len(b)} students failed or were absent in at least one subject{self._ctx(e)}." if len(d) else f"Every student cleared all subjects{self._ctx(e)}.", _tab(t))

    def a_absent_list(self, e): return self.a_fail_list(e, absent=True)

    def a_pass_list(self, e):
        b = self.base(e.term); d = b[b.backlogs == 0].sort_values("sgpa", ascending=False)
        t = d.rename(columns={"name": "Name", "enroll": "Enrollment", "sgpa": "SGPA", "pct": "%"})[["Name", "Enrollment", "SGPA", "%"]]
        return Answer(f"{len(d)} of {len(b)} students cleared every subject{self._ctx(e)}.", _tab(t))

    def a_grade_list(self, e):
        if not e.grades: return Answer("Which grade? e.g. *students with O grade in OS*.")
        if e.term != self.last or "grade" not in self.S: return Answer("An overall grade is only printed for the latest term. Name a subject, e.g. *students with O grade in <subject>*.")
        d = self.S[self.S.grade.isin(e.grades)].sort_values("cgpa", ascending=False)
        t = d.rename(columns={"name": "Name", "enroll": "Enrollment", "cgpa": "CGPA", "grade": "Grade"})[["Name", "Enrollment", "CGPA", "Grade"]]
        return Answer(f"{len(d)} student(s) with overall grade {'/'.join(e.grades)} (CGPA-based).", _tab(t))

    def a_topper_list(self, e):
        n = e.n or 10
        if e.subjects:
            c = e.subjects[0]; r = self._mt(e.term); r = r[r.code == c].sort_values(["total", "gpoints"], ascending=False).head(n)
            t = r.rename(columns={"name": "Name", "enroll": "Enrollment", "total": "Total", "grade": "Grade"})[["Name", "Enrollment", "Total", "Grade"]]
            return Answer(f"Top {len(t)} in {self.disp(e.term, c)}{self._ctx(e)}.", _tab(t))
        metric = self._requested_metric(e)
        if metric == "cgpa":
            d = self.S.dropna(subset=["cgpa"]).sort_values(["cgpa", "name"], ascending=[False, True]).head(n)
            t = d.rename(columns={"name": "Name", "enroll": "Enrollment", "cgpa": "CGPA"})[["Name", "Enrollment", "CGPA"]]
        else:
            d = self.base(e.term).sort_values(["sgpa", "pct"], ascending=False).head(n)
            t = d.rename(columns={"name": "Name", "enroll": "Enrollment", "sgpa": "SGPA", "pct": "%"})[["Name", "Enrollment", "SGPA", "%"]]
        note = " (CGPA is not available for this term; ranked by SGPA)" if "cgpa" in e.low and metric == "sgpa" else ""
        return Answer(f"Top {len(t)} students by {metric.upper()}{self._ctx(e)}{note}.", _tab(t))

    def a_atkt_list(self, e):
        if e.term != self.last or "status" not in self.S:
            return Answer("ATKT status is only printed for the latest term in this sheet. Showing who has backlogs instead.\n", self.a_fail_list(e).table)
        d = self.base(e.term); d = d[d.status == "ATKT"]
        t = d.rename(columns={"name": "Name", "enroll": "Enrollment", "failed": "Failed / absent in", "sgpa": "SGPA", "cr_earned": "Credits earned", "cr_reg": "Credits registered"})[["Name", "Enrollment", "Failed / absent in", "SGPA", "Credits earned", "Credits registered"]]
        return Answer(f"{len(d)} student(s) have ATKT status.", _tab(t))

    def _stats(self, term, codes):
        mt, rows = self._mt(term), []
        for c in codes:
            r = mt[mt.code == c]
            if r.empty: continue
            mx = self.smap[(term, c)].max_total
            rows.append({"Subject": self.lab(term, c), "Enrolled": len(r), "Passed": int((r.status == "Pass").sum()), "Failed": int((r.status == "Fail").sum()), "Absent": int((r.status == "AB").sum()),
                         "Pass %": round(100 * (r.status == "Pass").mean(), 1), "Avg marks": round(r.total.mean(), 1), "Avg %": round(100 * r.total.mean() / mx, 1) if pd.notna(mx) and mx else np.nan, "Highest": r.total.max(), "Lowest": r.total.min(),
                         "O / A+ %": round(100 * r.grade.isin(["O", "A+"]).mean(), 1)})
        df = pd.DataFrame(rows)
        return df.drop(columns="Avg %") if len(df) and df["Avg %"].isna().all() else df

    def a_pass_pct(self, e):
        codes = e.subjects or list(self.subj[(self.subj.term == e.term)].code)
        df = self._stats(e.term, codes).sort_values("Pass %")
        if e.subjects: return Answer("Pass % — " + "; ".join(f"{r.Subject}: {r._6}%" for r in df.itertuples()) + f"{self._ctx(e)}.", df, df.set_index("Subject")[["Pass %"]] if len(df) > 1 else None)
        big = df[df.Enrolled >= max(15, 0.1 * len(self.base(e.term)))]; low = (big if len(big) else df).iloc[0]
        b = self.base(e.term); ov = round(100 * (b.backlogs == 0).mean(), 1)
        tie = int((df["Pass %"] == low["Pass %"]).sum()) - 1
        return Answer(f"{ov}% of students cleared every subject{self._ctx(e)}. Lowest pass rate: {low.Subject} ({low['Pass %']}%)" + (f", tied with {tie} other subject(s)." if tie > 0 else "."), df, df.set_index("Subject")[["Pass %"]])

    def a_subject_stats(self, e):
        codes = e.subjects or list(self.subj[(self.subj.term == e.term) & self.subj.is_common].code)
        df = self._stats(e.term, codes)
        if len(df) == 1:
            r = df.iloc[0]; return Answer(f"{r.Subject}{self._ctx(e)}: {r.Enrolled} students, average {r['Avg marks']}, highest {r.Highest:g}, lowest {r.Lowest:g}, pass rate {r['Pass %']}%.", df)
        return Answer(f"Subject statistics{self._ctx(e)}.", df, df.set_index("Subject")[["Avg marks"]])

    def a_grade_dist(self, e):
        if e.subjects:
            r = self._mt(e.term); r = r[r.code == e.subjects[0]]; g = r.grade.value_counts(); label = self.disp(e.term, e.subjects[0])
        elif e.term == self.last and "grade" in self.S: g = self.S.grade.value_counts(); label = "overall grade (CGPA-based)"
        else: return Answer("Name a subject, e.g. *grade distribution in <subject>*.")
        df = g.rename_axis("Grade").reset_index(name="Students"); df["k"] = df.Grade.map(lambda x: GRADES.index(x) if x in GRADES else 99)
        df = df.sort_values("k").drop(columns="k")
        return Answer(f"Grade distribution — {label}{self._ctx(e)}.", df, df.set_index("Grade"))

    def a_class_summary(self, e):
        b = self.base(e.term); n = len(b); ok = int((b.backlogs == 0).sum())
        metric = self._requested_metric(e)
        if metric == "cgpa":
            ranked = self.S.dropna(subset=["cgpa"]).sort_values("cgpa", ascending=False)
            top = ranked.iloc[0]
            mean = ranked.cgpa.mean()
            range_text = f" (range {ranked.cgpa.min():g}–{ranked.cgpa.max():g})"
            metric_line = f"- Average CGPA {mean:.2f}{range_text}"
            topper_line = f"- Topper: {top['name']} (CGPA {top.cgpa:g})"
        else:
            top = b.sort_values("sgpa", ascending=False).iloc[0]
            metric_line = f"- Average SGPA {b.sgpa.mean():.2f} (range {b.sgpa.min():g}–{b.sgpa.max():g})"
            topper_line = f"- Topper: {top['name']} (SGPA {top.sgpa:g})"
        st = self._stats(e.term, list(self.subj[(self.subj.term == e.term)].code)); big = st[st.Enrolled >= max(15, 0.1 * n)]
        low = (big if len(big) else st).sort_values("Pass %").iloc[0]
        lines = [f"**Term {e.term} — {n} students**", f"- Cleared every subject: {ok} ({round(100 * ok / n, 1)}%); with backlogs: {n - ok}",
                 f"{metric_line}; average % {b.pct.mean():.1f}", topper_line,
                 f"- Lowest pass rate: {low.Subject} ({low['Pass %']}%)"]
        if "cgpa" in e.low and metric == "sgpa":
            lines.insert(1, "- CGPA is not present for this term; using SGPA for class ranking.")
        stt = b.status.value_counts().rename_axis("Status").reset_index(name="Students")
        if e.term == self.last: lines.append("- Final status: " + ", ".join(f"{r.Status} {r.Students}" for r in stt.itertuples()))
        return Answer("\n".join(lines), stt)

    def a_subject_info(self, e):
        s = self.subj[self.subj.term == e.term]
        if e.subjects: s = s[s.code.isin(e.subjects)]
        rows = []
        for r in s.itertuples():
            mx = [f"{k.upper()} {getattr(r, 'max_' + k):g}" for k in ("ise", "mse", "ese", "int", "oe_poe") if pd.notna(getattr(r, "max_" + k))]
            rows.append({"Code": r.code, "Subject": r.name if r.name != r.label else r.label, "Sheet label": r.label, "Type": r.type, "Part": r.part, "Credits": r.credits,
                         "Students": r.enrolled, "Max marks": f"{r.max_total:g}" if pd.notna(r.max_total) else "—", "Scheme": ", ".join(mx) or "—",
                         "Pass rule": r.pass_rule or (f"internal ≥ {r.int_pass}, external ≥ {r.ext_pass}" if pd.notna(r.int_pass) else "—")})
        df = pd.DataFrame(rows)
        note = "" if s.in_catalog.any() else (" Marking schemes come from the sheet; upload the syllabus for full subject names." if s.max_total.notna().any()
                                              else " Upload the syllabus image for full names and marking schemes.")
        return Answer(f"{len(df)} subject(s) in Term {e.term}.{note}", _tab(df))

    def _student_term(self, en, term):
        return self.base(term).query("enroll == @en")

    def a_student_report(self, e):
        if len(e.students) > 1:
            d = self.S.loc[e.students][["name", "enroll"]].rename(columns={"name": "Name", "enroll": "Enrollment"})
            return Answer(f"{len(e.students)} students match — which one? Use the enrollment number.", d.sort_values("Name"))
        en = e.students[0]; nm = self.S.at[en, "name"]; b = self.base(e.term); row = b[b.enroll == en]
        if row.empty: return Answer(f"**{nm}** ({en}) has no results in Term {e.term}. Available terms: {', '.join(t for t in self.terms if not self.base(t).query('enroll == @en').empty)}.")
        s = row.iloc[0]; m = self._mt(e.term); m = m[m.enroll == en].sort_values(["gpoints", "total"], ascending=False)
        metric = self._requested_metric(e)
        if metric == "cgpa" and pd.isna(self.S.at[en, "cgpa"]):
            metric = "sgpa"
        if metric == "cgpa":
            population = self.S.dropna(subset=["cgpa"])
            value = self.S.at[en, "cgpa"]
            rank = int((population.cgpa > value).sum()) + 1
            avg = population.cgpa.mean()
            rank_line = f"Rank by CGPA: {rank} of {len(population)}; class average {avg:.2f}."
        else:
            rank = int((b.sgpa > s.sgpa).sum()) + 1; avg = b.sgpa.mean()
            rank_line = f"Rank by SGPA: {rank} of {len(b)}; class average {avg:.2f} ({'above' if s.sgpa > avg else 'below' if s.sgpa < avg else 'at'} average)."
            if "cgpa" in e.low:
                rank_line += " CGPA is unavailable for this student/term, so SGPA was used."
        fin = f", CGPA {self.S.at[en, 'cgpa']:g}, overall grade {self.S.at[en, 'grade']}" if e.term == self.last and "cgpa" in self.S and pd.notna(self.S.at[en, "cgpa"]) else ""
        tx = [f"**{nm}** ({en}) — Term {e.term} SGPA **{s.sgpa:g}** ({s.pct:g}%){fin}, status **{s.status}**.",
              rank_line,
              f"Credits earned: {s.cr_earned:g} of {s.cr_reg:g}."]
        bad = m[m.status != "Pass"]
        tx.append("Backlogs: " + ", ".join(f"{r.label} ({r.status})" for r in bad.itertuples()) + "." if len(bad) else "No backlogs this term.")
        sc = m[(m.credits > 0) & m.gpoints.notna()]
        if len(sc) >= 4: tx.append("Strongest: " + ", ".join(f"{r.label} ({r.grade})" for r in sc.head(2).itertuples()) + ". Weakest: " + ", ".join(f"{r.label} ({r.grade})" for r in sc.tail(2).iloc[::-1].itertuples()) + ".")
        h = self.T[self.T.enroll == en].sort_values("term_n")
        if len(h) > 1: tx.append("SGPA by term: " + " → ".join(f"{r.term} {r.sgpa:g}" for r in h.itertuples() if pd.notna(r.sgpa)) + ".")
        t = m.rename(columns={"label": "Subject", "ise": "ISE", "mse": "MSE", "internal": "Internal", "int": "INT", "ese": "ESE", "poe": "POE", "other": "Marks", "total": "Total", "grade": "Grade", "status": "Status"})
        t = t[["Subject", "ISE", "MSE", "Internal", "INT", "ESE", "POE", "Marks", "Total", "Grade", "Status"]].dropna(axis=1, how="all")
        return Answer("\n\n".join(tx), t)

    def a_sgpa_trend(self, e):
        if len(e.students) == 1:
            en = e.students[0]; h = self.T[self.T.enroll == en].sort_values("term_n")[["term", "sgpa"]].dropna().rename(columns={"term": "Term", "sgpa": "SGPA"})
            d = "" if len(h) < 2 else f" Change vs previous term: {h.SGPA.iloc[-1] - h.SGPA.iloc[-2]:+.2f}."
            return Answer(f"SGPA trend for {self.S.at[en, 'name']}.{d}", h, h.set_index("Term"))
        h = self.T.dropna(subset=["sgpa"]).groupby(["term_n", "term"]).sgpa.agg(["mean", "count"]).reset_index().rename(columns={"term": "Term", "mean": "Avg SGPA", "count": "Students"})
        h["Avg SGPA"] = h["Avg SGPA"].round(2)
        return Answer("Class average SGPA by term.", h[["Term", "Avg SGPA", "Students"]], h.set_index("Term")[["Avg SGPA"]])

    # ---- Stage 4 additions
    def _who(self, term, codes, st):
        mt = self._mt(term); out = set()
        for c in codes:
            r = mt[mt.code == c]; kind, val = st
            if kind == "pass": ok = r.status == "Pass"
            elif kind == "fail": ok = r.status != "Pass"
            elif kind == "absent": ok = r.status == "AB"
            elif kind == "grade": ok = r.grade.isin(val if isinstance(val, (list, tuple)) else [val])
            else:
                op, v, pct = val; mx = self.smap[(term, c)].max_total
                if pct and pd.isna(mx): raise ValueError("A percentage needs the subject's maximum marks - upload the syllabus, or use marks (e.g. *above 40*).")
                ok = _cmp(r.total / mx * 100 if pct else r.total, op, v)
            out |= set(r[ok].enroll)
        return out

    def _phrase(self, term, codes, st):
        kind, val = st; lab = " / ".join(self.lab(term, c) for c in codes)
        if kind == "pass": return f"passed {lab}"
        if kind == "fail":
            ab = self._mt(term).query("status == 'AB' and code in @codes")
            return f"failed {lab}" + (" (absent counts as failed)" if len(ab) else "")
        if kind == "absent": return f"absent in {lab}"
        if kind == "grade": return f"got {'/'.join(val if isinstance(val, (list, tuple)) else [val])} in {lab}"
        op, v, pct = val; return f"scored {OP_WORD[op]} {v:g}{'%' if pct else ''} in {lab}"

    def a_conditions(self, e):
        G, srt, prev, conds = e.groups, sorted(e.states), None, []
        for i, (a, b, codes) in enumerate(G):
            lo = G[i - 1][1] if i else 0; hi = G[i + 1][0] if i + 1 < len(G) else len(e.low)
            win = [s for s in srt if lo <= s[0] < a] or [s for s in srt if b <= s[0] < hi]
            if win:
                kinds = {s[1] for s in win}
                st = ("grade", tuple(dict.fromkeys(s[2] for s in win))) if kinds == {"grade"} else win[-1][1:]
            else: st = prev
            prev = st; conds.append((codes, st))
        if any(st is None for _, st in conds): return Answer("I found the subject but not what to check (pass, fail, a grade, or marks). Try *who failed in <subject>*.")
        try: sets = [self._who(e.term, c, st) for c, st in conds]
        except ValueError as ex: return Answer(str(ex))
        groups, cur = [], [sets[0]]; desc = [self._phrase(e.term, *conds[0])]
        for i in range(1, len(conds)):
            joiner = "OR" if re.search(r"\bor\b", e.low[G[i - 1][1]:G[i][0]]) else "AND"
            if joiner == "OR": groups.append(cur); cur = [sets[i]]
            else: cur.append(sets[i])
            desc.append(joiner); desc.append(self._phrase(e.term, *conds[i]))
        groups.append(cur)
        res = set().union(*[set.intersection(*g) for g in groups]); text = " ".join(desc)
        cols = list(dict.fromkeys(c for cs, _ in conds for c in cs)); mt = self._mt(e.term)
        if len(e.students) == 1:
            en = e.students[0]; nm = self.S.at[en, "name"]; lines = []
            for c in cols:
                r = mt[(mt.code == c) & (mt.enroll == en)]
                lines.append(f"{self.lab(e.term, c)}: " + (f"{r.total.iloc[0]:g} marks, grade {r.grade.iloc[0]}, {r.status.iloc[0]}" if len(r) else "not taken"))
            return Answer(f"**{'Yes' if en in res else 'No'}** — {nm} {'did' if en in res else 'did not'} match: {text}.\n\n" + "\n".join(f"- {l}" for l in lines))
        rows = []
        for en in sorted(res, key=lambda x: self.S.at[x, "name"]):
            row = {"Name": self.S.at[en, "name"], "Enrollment": en}
            for c in cols:
                r = mt[(mt.code == c) & (mt.enroll == en)]
                row[self.lab(e.term, c)] = f"{r.total.iloc[0]:g} · {r.grade.iloc[0]}" if len(r) else "–"
            rows.append(row)
        df = pd.DataFrame(rows)
        if not len(df): return Answer(f"Nobody {text}{self._ctx(e)}.")
        extra = " Say *or* instead of *and* for either." if len(e.groups) > 1 and " AND " in f" {text} " and not re.search(r"\bbut\b", e.low) else ""
        return Answer(f"{len(df)} student(s) {text}{self._ctx(e)}.{extra}", df)

    def a_compare(self, e):
        if len(e.groups) >= 2 and not e.students:
            df = self._stats(e.term, e.subjects)
            if len(df) < 2: return None
            hard = df.sort_values("Pass %").iloc[0]
            txt = f"{hard.Subject} has the lower pass rate ({hard['Pass %']}%)." if df["Pass %"].nunique() > 1 else f"Same pass rate ({hard['Pass %']}%)."
            if "Avg %" in df: txt += f" {df.sort_values('Avg %', ascending=False).iloc[0].Subject} has the higher average score ({df['Avg %'].max()}% of max marks)."
            else: txt += " Average marks aren't compared because the subjects are out of different totals (upload the syllabus to compare as %)."
            return Answer(txt, df, df.set_index("Subject")[["Pass %"]])
        if len(e.enroll) >= 2: picks = [[x] for x in e.enroll[:2] if x in self.names]
        else:
            body = re.sub(r"\b(compare|comparison|difference between|between)\b", " ", e.low)
            picks = []
            for part in re.split(r"\b(?:and|vs\.?|versus|with|&)\b", body):
                toks = [w for w in re.findall(r"[a-z]+", part) if len(w) >= 3 and w not in self.stop]
                got = self._match(toks, fuzzy=True)[0] if toks else []
                if got: picks.append(got)
        if len(picks) < 2: return None
        for p in picks:
            if len(p) > 1:
                d = self.S.loc[p][["name", "enroll"]].rename(columns={"name": "Name", "enroll": "Enrollment"})
                return Answer("One of those names matches several students — use the enrollment number for each.", d.sort_values("Name"))
        a, b = picks[0][0], picks[1][0]
        if a == b: return Answer("Those match the same student. Give two different students.")
        A, B = self.S.at[a, "name"], self.S.at[b, "name"]; bt = self.base(e.term); ra, rb = bt[bt.enroll == a], bt[bt.enroll == b]
        if ra.empty or rb.empty: return Answer(f"Both students need results in Term {e.term} to compare.")
        ra, rb = ra.iloc[0], rb.iloc[0]; mt = self._mt(e.term)
        ma, mb = mt[mt.enroll == a].set_index("code"), mt[mt.enroll == b].set_index("code"); common = [c for c in ma.index if c in mb.index]
        w = sum(ma.at[c, "total"] > mb.at[c, "total"] for c in common); l = sum(ma.at[c, "total"] < mb.at[c, "total"] for c in common)
        metric = self._requested_metric(e)
        if metric == "cgpa" and (pd.isna(self.S.at[a, "cgpa"]) or pd.isna(self.S.at[b, "cgpa"])):
            metric = "sgpa"
        if metric == "cgpa":
            ca, cb = self.S.at[a, "cgpa"], self.S.at[b, "cgpa"]
            population = self.S.dropna(subset=["cgpa"])
            rows = [{"Metric": "CGPA", A: ca, B: cb},
                    {"Metric": "Rank in class by CGPA", A: int((population.cgpa > ca).sum()) + 1,
                     B: int((population.cgpa > cb).sum()) + 1},
                    {"Metric": "SGPA", A: ra.sgpa, B: rb.sgpa},
                    {"Metric": "Percentage", A: ra.pct, B: rb.pct},
                    {"Metric": "Backlogs", A: ra.backlogs, B: rb.backlogs}]
        else:
            rows = [{"Metric": "SGPA", A: ra.sgpa, B: rb.sgpa}, {"Metric": "Percentage", A: ra.pct, B: rb.pct},
                    {"Metric": "Rank in class by SGPA", A: int((bt.sgpa > ra.sgpa).sum()) + 1,
                     B: int((bt.sgpa > rb.sgpa).sum()) + 1}, {"Metric": "Backlogs", A: ra.backlogs, B: rb.backlogs}]
        rows += [{"Metric": self.lab(e.term, c), A: f"{ma.at[c, 'total']:g} · {ma.at[c, 'grade']}", B: f"{mb.at[c, 'total']:g} · {mb.at[c, 'grade']}"} for c in common]
        df = pd.DataFrame(rows)
        for c in (A, B): df[c] = df[c].map(lambda v: f"{v:g}" if isinstance(v, (int, float, np.integer, np.floating)) else str(v))
        va, vb = (self.S.at[a, "cgpa"], self.S.at[b, "cgpa"]) if metric == "cgpa" else (ra.sgpa, rb.sgpa)
        lead = A if va > vb else B if vb > va else None
        head = f"{lead} has the higher {metric.upper()} ({max(va, vb):g} vs {min(va, vb):g})." if lead else f"Both have {metric.upper()} {va:g}."
        if metric == "sgpa" and "cgpa" in e.low:
            head += " CGPA is unavailable for this comparison, so SGPA was used."
        return Answer(f"{head} In {len(common)} common subjects: {A} scored higher in {w}, {B} in {l}, {len(common) - w - l} tied{self._ctx(e)}.", df)

    def a_metric_filter(self, e):
        op, v, pct = e.cmps[0]; low = e.low; b = self.base(e.term)
        if "cgpa" in low and self._cgpa_available(e.term): d = self.S.rename(columns={"cgpa": "val"})[["name", "enroll", "val"]].dropna(subset=["val"]); metric = "CGPA (cumulative)"
        elif pct or re.search(r"percent|%", low): d = b.rename(columns={"pct": "val"})[["name", "enroll", "val"]]; metric = "percentage"
        else: d = b.rename(columns={"sgpa": "val"})[["name", "enroll", "val"]]; metric = "SGPA"
        d = d[_cmp(d.val, op, v)].sort_values("val", ascending=op in ("<", "<="))
        t = d.rename(columns={"name": "Name", "enroll": "Enrollment", "val": metric})
        note = " (CGPA is not available for this term; used SGPA)" if "cgpa" in low and metric == "SGPA" else ""
        return Answer(f"{len(t)} student(s) with {metric} {OP_WORD[op]} {v:g}{self._ctx(e)}{note}.", _tab(t))

    def a_backlog_filter(self, e):
        b = self.base(e.term).copy(); low = e.low
        if e.cmps: op, v = e.cmps[0][0], e.cmps[0][1]
        elif re.search(r"\b(no|zero|without|nil)\b", low): op, v = "==", 0
        elif e.n is not None: op, v = "==", e.n
        else: op, v = ">", 0
        d = b[_cmp(b.backlogs, op, v)].sort_values(["backlogs", "name"], ascending=[False, True])
        t = d.rename(columns={"name": "Name", "enroll": "Enrollment", "backlogs": "Backlogs", "failed": "Subjects", "sgpa": "SGPA"})[["Name", "Enrollment", "Backlogs", "Subjects", "SGPA"]]
        what = "no backlogs" if (op, v) == ("==", 0) else f"exactly {v:g} backlog(s)" if op == "==" else f"backlogs {OP_WORD[op]} {v:g}"
        return Answer(f"{len(t)} student(s) with {what}{self._ctx(e)}.", _tab(t))

    def a_sgpa_change(self, e):
        ts = [t for t in dict.fromkeys(e.terms) if t in self.terms]
        if len(ts) < 2: ts = self.terms[-2:]
        t0, t1 = sorted(ts[:2], key=self.terms.index)
        p = self.T[self.T.term.isin([t0, t1])].pivot_table(index=["enroll", "name"], columns="term", values="sgpa").dropna().reset_index()
        p["Change"] = (p[t1] - p[t0]).round(2)
        p = p.rename(columns={"name": "Name", "enroll": "Enrollment", t0: f"SGPA {t0}", t1: f"SGPA {t1}"})
        up = bool(re.search(r"improv|increas|progress|rose|better", e.low)); dn = bool(re.search(r"drop|declin|decreas|fell|worse|regress", e.low))
        n = e.n or 10
        if up and not dn: d, what = p[p.Change > 0].sort_values("Change", ascending=False), "improved"
        elif dn and not up: d, what = p[p.Change < 0].sort_values("Change"), "dropped"
        else: d, what = pd.concat([p.sort_values("Change", ascending=False).head(n), p.sort_values("Change").head(n)]).drop_duplicates("Enrollment"), "changed most"
        d = d[["Name", "Enrollment", f"SGPA {t0}", f"SGPA {t1}", "Change"]]
        return Answer(f"{len(d)} of {len(p)} students {what} between Term {t0} and Term {t1}. Class average moved {p[f'SGPA {t1}'].mean() - p[f'SGPA {t0}'].mean():+.2f}.", _tab(d.head(max(n, 10)) if what == "changed most" else d))
