"""Step 1 - profile the sheet: name column, ID column, first student row, mark columns. Content-based, offline.
Every decision can be overridden by the user (see u4_layouts.py), so an odd sheet never needs a code change."""
import pandas as pd
from .ingest import UnsupportedLayout
from .u0_common import *
from .u0_common import _t, _is_name, _is_mark, _syn_score

def profile(raw, overrides=None):
    """Detect id/name/mark columns, header band and student rows. `overrides` (from the UI / a saved layout) wins over detection."""
    o = overrides or {}; t = raw.map(_t); R, C = t.shape
    name_c = o.get("name_col", int(t.map(_is_name).sum().idxmax()))
    nm = t[name_c]
    ok = nm.map(lambda s: bool(s) and not HDRWORD.search(s) and not re.fullmatch(r"[\d.\-]+", s)) if "name_col" in o else nm.map(_is_name)
    if ok.sum() < 3: raise UnsupportedLayout("No column of student names found.")
    first = o.get("first_row", next((r for r in range(R - 2) if ok.iloc[r] and ok.iloc[r:r + 5].mean() >= 0.8), None))
    if first is None: raise UnsupportedLayout("Could not find where the student rows start.")
    rows = [r for r in range(first, R) if ok.iloc[r]]
    band = list(o.get("header_rows", range(max(0, first - 8), first)))
    hdr = {c: " ".join(t.iat[r, c] for r in band if t.iat[r, c]) for c in range(C)}
    body = t.iloc[rows]
    def seq(c): v = pd.to_numeric(body[c], errors="coerce"); return v.notna().all() and (v.diff().dropna() == 1).all()
    cands = {}
    for c in range(C):
        if c == name_c or c in o.get("skip_cols", ()): continue
        v = body[c][body[c] != ""]
        if len(v) < 0.9 * len(body) or v.nunique() < 0.98 * len(v) or seq(c) or SKIP.search(hdr[c]) and not re.search("roll|seat|prn|id|enrol", hdr[c], re.I): continue
        idlike = v.map(lambda s: bool(re.fullmatch(r"[A-Za-z]{0,4}[\s-]?\d+[A-Za-z0-9]*", s))).mean()
        cands[c] = 3 * _syn_score(hdr[c], "id") + idlike
    if "id_col" in o: id_c = o["id_col"]
    elif cands: id_c = max(cands, key=cands.get)
    else: raise UnsupportedLayout("No unique student ID / roll-number column found.")
    used = {name_c, id_c}
    if "mark_cols" in o: marks = {c: 0 for c in o["mark_cols"] if c not in used}
    else:
        marks = {}
        for c in range(C):
            if c in used: continue
            v = body[c][body[c] != ""]
            if len(v) and v.map(_is_mark).mean() >= 0.9 and not SKIP.search(hdr[c]) and not (v.map(lambda s: s.isdigit()).all() and v.nunique() > 0.9 * len(v)):
                marks[c] = len(v)
        if marks:
            for c in range(min(marks), max(marks) + 1):             # a blank column inside the mark block (nobody entered it) stays a subject
                if c not in used and c not in marks and (body[c] == "").all() and hdr[c] and not SKIP.search(hdr[c]): marks[c] = 0
    if not marks: raise UnsupportedLayout("No mark columns found.")
    if len(marks) > 60 and "mark_cols" not in o: raise UnsupportedLayout("Too many mark columns - this looks like a multi-term ERP sheet, not a single register.")
    conf = round(min(1, cands.get(id_c, 0) / 3 + 0.2), 2) if "id_col" not in o else 1.0
    return dict(name_col=name_c, id_col=id_c, first_row=first, rows=rows, band=band, header=hdr,
                mark_cols=sorted(marks), confidence=conf, t=t)
