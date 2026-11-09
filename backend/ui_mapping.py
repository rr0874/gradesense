"""Streamlit panel for sheets the automatic reader is unsure about: shows what was detected, lets staff fix it, saves the layout
(layouts/*.json) so the same sheet shape is read automatically next time. No code update needed for new department formats."""
import pandas as pd, streamlit as st
from .u1_detect import profile
from .u2_headers import _labels
from .u3_build import read_generic
from .u4_layouts import load_layout, save_layout, delete_layout

def layout_wizard(raw, key="wiz"):
    """Call with the sheet as pd.read_excel(..., header=None). Returns True after a layout is saved -> caller clears caches and reruns."""
    saved = load_layout(raw)
    try: p = profile(raw, saved); lab = _labels(p["t"], p["band"], p["mark_cols"], saved); err = ""
    except Exception as ex: p, lab, err = None, {}, str(ex)
    st.warning("Automatic detection is unsure: " + err) if err else st.info("Detected columns below - check them, fix if wrong, then save.")
    st.caption("Raw sheet (columns are numbered; rows are Excel row numbers)")
    view = raw.head(15).copy(); view.columns = [f"{i + 1}" for i in range(raw.shape[1])]; view.index = range(1, len(view) + 1)
    st.dataframe(view.astype(str), height=260)
    C = raw.shape[1]; name_of = lambda c: f"{c + 1}: {str(raw.iat[min(len(raw) - 1, (p['first_row'] - 1) if p and p['first_row'] else 0), c])[:22]}"
    pick = lambda label, default, k: st.selectbox(label, range(C), index=min(default, C - 1), format_func=lambda c: f"column {c + 1}", key=f"{key}_{k}")
    c1, c2, c3 = st.columns(3)
    with c1: id_c = pick("Student ID column", p["id_col"] if p else 0, "id")
    with c2: name_c = pick("Student name column", p["name_col"] if p else min(1, C - 1), "name")
    with c3: first = st.number_input("First student row (Excel row no.)", 1, len(raw), (p["first_row"] + 1) if p else 4, key=f"{key}_first") - 1
    band = st.multiselect("Header rows (subject names / max marks)", range(len(raw)), default=[r for r in (p["band"] if p else range(max(0, first - 3), first)) if r < first],
                          format_func=lambda r: f"row {r + 1}", key=f"{key}_band")
    marks = st.multiselect("Mark columns", range(C), default=p["mark_cols"] if p else [], format_func=lambda c: f"column {c + 1}", key=f"{key}_marks")
    names, maxes = {}, {}
    for c in marks:
        a, b = st.columns([3, 1]); dflt = lab.get(c, ("", "", float("nan")))
        names[c] = a.text_input(f"Subject for column {c + 1}", dflt[0], key=f"{key}_n{c}")
        maxes[c] = b.number_input("Max", 0.0, 1000.0, 0.0 if pd.isna(dflt[2]) else float(dflt[2]), key=f"{key}_m{c}")
    ov = {"id_col": id_c, "name_col": name_c, "first_row": int(first), "header_rows": sorted(band), "mark_cols": sorted(marks),
          "subject_names": {c: n for c, n in names.items() if n.strip()}, "max_marks": {c: m for c, m in maxes.items() if m > 0}}
    try:
        wb = read_generic(raw, ov); st.success(f"Preview: {wb.info['students']} students x {wb.info['subjects']} subjects")
        pv = wb.marks.merge(wb.students[["enroll", "name"]], on="enroll").merge(wb.subjects[["code", "label"]], on="code")
        st.dataframe(pv.assign(mark=pv.mse.where(pv.mse_status.eq("Marked"), pv.mse_status)).pivot(index=["enroll", "name"], columns="label", values="mark").head(8).astype(str))
        if len(wb.info["flags"]): st.error(f"{len(wb.info['flags'])} marks are above their maximum - fix the Max values or the sheet."); st.dataframe(wb.info["flags"])
        ready = True
    except Exception as ex: st.error(f"Cannot read with these settings: {ex}"); ready = False
    d1, d2 = st.columns(2)
    if d1.button("Save layout and analyse", disabled=not ready, key=f"{key}_save"): save_layout(raw, ov); return True
    if saved and d2.button("Forget saved layout", key=f"{key}_del"): delete_layout(raw); st.rerun()
    return False
