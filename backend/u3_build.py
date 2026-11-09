"""Step 3 - build the Workbook (students, subjects, marks with source cell, flags, mapping) from steps 1-2."""
import numpy as np, pandas as pd
from .ingest import Workbook, UnsupportedLayout, ALLF, _mse_status
from .u0_common import *
from .u1_detect import profile
from .u2_headers import _labels

def read_generic(raw, overrides=None) -> Workbook:
    p = profile(raw, overrides); t, cols = p["t"], p["mark_cols"]
    lab = _labels(t, p["band"], cols, overrides)
    for c, (s, _, _) in lab.items():
        if not s: raise UnsupportedLayout(f"Column {c + 1}: could not find a subject heading above the marks.")
    ids = [t.iat[r, p["id_col"]].upper() for r in p["rows"]]
    if len(set(ids)) != len(ids): raise UnsupportedLayout("Duplicate student IDs; refusing to combine or misassign marks.")
    subj_cols = {}
    for c, (s, comp, mx) in lab.items(): subj_cols.setdefault(s, []).append(c)
    stu = pd.DataFrame({"enroll": ids, "name": [t.iat[r, p["name_col"]] for r in p["rows"]]})
    stu["prefix"] = stu.enroll.str.extract(r"^([A-Z]+)")[0].fillna(""); stu["batch"] = float("nan")
    recs, subs, flags = [], [], []
    for k, (s, cs) in enumerate(subj_cols.items(), 1):
        code = f"REG-C{k}"; mxs = []
        for c in cs:
            vals = [raw.iat[r, c] for r in p["rows"]]
            nums = pd.to_numeric(pd.Series(vals), errors="coerce"); mx = lab[c][2]
            if np.isnan(mx):
                top = nums.max(); mx = next((m for m in STD_MAX if m >= top), float("nan")) if pd.notna(top) else float("nan")
            mxs.append(mx)
            for r, v, i in zip(p["rows"], vals, ids):
                if pd.notna(mx) and pd.notna(pd.to_numeric(v, errors="coerce")) and not 0 <= float(v) <= mx:
                    flags.append((i, s, v, f"outside 0-{mx:g}", f"R{r + 1}C{c + 1}"))
        comps = [lab[c][1] for c in cs]
        for j, i in enumerate(ids):
            row = {"enroll": i, "term": "MSE", "code": code, "src": ";".join(f"R{p['rows'][j] + 1}C{c + 1}" for c in cs)}
            vals = {lab[c][1]: raw.iat[p["rows"][j], c] for c in cs}
            st = [_mse_status(v) for v in vals.values()]
            for f in ALLF: row[f] = float("nan") if f not in ("grade", "status") else ""
            for comp, v in vals.items():
                if comp in row and comp != "total" and pd.notna(pd.to_numeric(v, errors="coerce")): row[comp] = float(v)
            tot = pd.to_numeric(vals.get("total", np.nan), errors="coerce")
            parts = [pd.to_numeric(v, errors="coerce") for c, v in vals.items() if c != "total"]
            row["total"] = tot if pd.notna(tot) else (np.nansum(parts) if any(pd.notna(parts)) else float("nan"))
            row["mse_status"] = "AB" if "AB" in st else "Marked" if "Marked" in st else next((x for x in st if x.startswith("Invalid")), "Unreported")
            row["status"] = "AB" if row["mse_status"] == "AB" else "Unreported"
            recs.append(row)
        subs.append({"term": "MSE", "term_n": 1, "abbr": s, "label": s, "name": s, "code": code, "dept": "", "type": "Other", "part": "Theory",
                     "int_pass": None, "ext_pass": None, "enrolled": len(ids), "credits": float("nan"), "max_total": float(np.nansum(mxs)) if not np.isnan(mxs).all() else float("nan"),
                     "pass_rule": "", "has_mse": "mse" in comps, "share": 1.0, "is_common": True,
                     **{f"max_{n}": float("nan") for n in ("ise", "mse", "ese", "int", "oe_poe")}})
    marks = pd.DataFrame(recs); marks["mse_status"] = marks.pop("mse_status")
    subjects = pd.DataFrame(subs)
    terms = pd.DataFrame({"enroll": ids, "term": "MSE", "term_n": 1, **{k: float("nan") for k in ("sgpa", "pct", "cr_reg", "cr_earned", "grade_total")}})
    warn = [f"Auto-detected layout (confidence {p['confidence']}): ID='{p['header'][p['id_col']] or 'col ' + str(p['id_col'] + 1)}', "
            f"name col {p['name_col'] + 1}, {len(subj_cols)} subjects. Confirm before relying on it."]
    if flags: warn.append(f"{len(flags)} marks fall outside their maximum; see info['flags'].")
    info = {"students": len(ids), "terms": ["MSE"], "subjects": len(subs), "rows": len(raw), "mse_available": True, "term_max_marks": {},
            "warnings": warn, "flags": pd.DataFrame(flags, columns=["enroll", "subject", "value", "reason", "cell"]),
            "mapping": {"id_col": p["id_col"], "name_col": p["name_col"], "subjects": {s: cs for s, cs in subj_cols.items()}}}
    return Workbook(stu, terms, subjects, marks, pd.DataFrame({"enroll": ids}), info)
