"""Step 0 - shared constants and tiny helpers (synonyms, regexes, cell classifiers)."""
import re
from difflib import get_close_matches
import pandas as pd
from .ingest import _norm

SYN = {"id": ["student unique no", "student unique number", "student id", "enrollment no", "enrollment number", "enrolment no",
              "prn", "prn no", "erp id", "reg no", "registration no", "roll no", "roll number", "roll", "id", "seat no", "seat number"],
       "name": ["student name", "name of student", "name of the student", "candidate name", "name", "full name"]}
SKIP = re.compile(r"%|\b(total|tot|percent\w*|percentage|result|grade|gr|rank|sgpa|cgpa|status|remark\w*|sr|s no|serial|sign\w*|mother|father|avg|average)\b", re.I)
GENERIC = re.compile(r"^(marks?|mse|ise|ese|int|internal|poe|oral|score|obtained|out of \d+|max\.?|\(?\d+\)?)(\s*\(?\d+\)?)?$", re.I)
COMP = {"ise": "ise", "mse": "mse", "ese": "ese", "int": "int", "internal": "int", "poe": "poe", "oral": "poe", "total": "total", "tot": "total"}
HDRWORD = re.compile(r"\b(name|student|roll|marks?|total|sr|no|division|year|semester|term|sheet|class)\b", re.I)
ABS = {"ab", "abs", "absent", "-", "—", "–"}
STD_MAX = (10, 15, 20, 25, 30, 40, 50, 60, 70, 75, 80, 100)

def _t(v):
    if pd.isna(v): return ""
    if isinstance(v, float) and v.is_integer(): v = int(v)
    return re.sub(r"\s+", " ", str(v)).strip()
def _is_name(s): return bool(re.fullmatch(r"[A-Za-z][A-Za-z.' ]{2,}", s)) and " " in s.strip() and not HDRWORD.search(s)
def _is_mark(s): return s.casefold().replace(".", "") in ABS or bool(re.fullmatch(r"\d{1,3}(\.\d+)?", s))
def _syn_score(label, kind):
    k = _norm(label); hit = get_close_matches(k, SYN[kind], 1, 0.82)
    return (len(SYN[kind]) - SYN[kind].index(hit[0])) / len(SYN[kind]) if hit else 0.0

