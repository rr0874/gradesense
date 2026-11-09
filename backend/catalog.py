"""Syllabus catalogs (Stage 2 fills these automatically; today they are JSON files).
A catalog only ENRICHES subjects (full name, marking scheme). Analysis never depends on it."""
import json, re
import pandas as pd
from pathlib import Path
CAT_DIR = Path(__file__).parent / "catalogs"

def load(path): return json.load(open(path))
def all_catalogs(): return {p.stem: load(p) for p in sorted(CAT_DIR.glob("*.json"))}
def _key(code): return re.sub(r"^(\d{3}[A-Z]*?(?:PCC|HSSM|MDM|OEC|VSEC|VEC|AEC|CEP|CCA|EEM|MC|FY))[LPM]?(\d{3})$", r"\1\2", str(code))

_STOP = {"and", "of", "with", "the", "for", "to", "in"}
def _initials(name): return "".join(w[0] for w in re.findall(r"[A-Za-z]+", name) if w.lower() not in _STOP).upper()

def match(row, cat):
    """Catalog course for a subjects-table row: explicit abbr map, else same code ignoring the L/P part, else name initials == abbr."""
    for c in cat["courses"]:
        if row["abbr"] in c.get("sheet", {}): return c
    for c in cat["courses"]:
        if _key(c["code"]) == _key(row["code"]): return c
    ab = re.sub(r"[^A-Z]", "", str(row["abbr"]).upper())
    if len(ab) >= 3:
        for c in cat["courses"]:
            if _initials(c["name"]) == ab: return c
    return None

def apply_catalog(subjects, cat):
    """Add name / scheme / alias columns from a catalog (call repeatedly: one catalog per semester). Returns (subjects, unmatched labels)."""
    s = subjects.copy(); miss = []; hit = {}
    if "in_catalog" not in s: s["in_catalog"], s["aliases"] = False, ""
    for k in ("ISE", "MSE", "ESE", "INT", "OE_POE", "total"):
        if "max_" + k.lower() not in s: s["max_" + k.lower()] = float("nan")
    for i, r in s.iterrows():
        if cat.get("semester") and r["term"] != cat["semester"]: continue
        c = match(r, cat)
        if c is None: miss.append(r["label"]); continue
        hit[i] = c
    per_course = {}
    for c in hit.values(): per_course[c["code"]] = per_course.get(c["code"], 0) + 1
    for i, c in hit.items():
        lab = per_course[c["code"]] > 1 and s.at[i, "part"] == "Practical"     # a course with theory + practical sheet columns
        s.at[i, "name"] = c["name"] + (" (Lab)" if lab else "")
        s.at[i, "in_catalog"] = True
        s.at[i, "aliases"] = "|".join(c.get("aliases", []))
        for k in ("ISE", "MSE", "ESE", "INT", "OE_POE", "total"):
            if isinstance(c.get(k), (int, float)) and pd.isna(s.at[i, "max_" + k.lower()]): s.at[i, "max_" + k.lower()] = c[k]
    return s, miss
