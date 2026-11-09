"""Stage 3 seed - reconcile the parsed data against the sheet's own totals so wrong numbers are caught, not shown."""
from dataclasses import dataclass, field
import numpy as np
import pandas as pd

@dataclass
class Check:
    name: str
    level: str            # "error" = data likely wrong, "info" = worth knowing
    checked: int
    failed: int
    examples: list = field(default_factory=list)
    note: str = ""
    @property
    def ok(self): return self.failed == 0

def _mk(name, level, df, bad, fmt, note=""):
    ex = [fmt(r) for r in df[bad].head(3).itertuples()]
    return Check(name, level, int(len(df)), int(bad.sum()), ex, note)

def run_checks(ds):
    m, t, s, st, cum = ds.wb.marks, ds.wb.terms, ds.wb.subjects, ds.wb.students, ds.wb.cumulative
    out = []
    out.append(Check("Unique enrollment numbers", "error", len(st), int(st.enroll.duplicated().sum()),
                     list(st.enroll[st.enroll.duplicated()].head(3))))
    sub = np.where(m.internal.notna(), m.internal.fillna(0), m.ise.fillna(0) + m.mse.fillna(0))
    parts = pd.Series(sub, index=m.index) + m[["int", "ese", "poe", "other"]].fillna(0).sum(axis=1)
    x = m[m[["ise", "mse", "internal", "int", "ese", "poe", "other"]].notna().any(axis=1) & m.total.notna()]
    out.append(_mk("Total = sum of its mark components", "error", x, (x.total - parts[x.index]).abs() > 0.51,
                   lambda r: f"{r.enroll} {r.code}: components add to {parts[r.Index]:g} but total is {r.total:g}"))
    grade = m.grade.astype(str).str.strip().str.replace(r"\s*\*$", "", regex=True)
    known_grade = grade.isin(["O", "A+", "A", "B+", "B", "C", "P", "F", "AB"])
    bad_grade_status = known_grade & (((grade == "F") & m.status.ne("Fail")) |
                                      ((grade == "AB") & ~m.status.isin(["AB", "Fail"])) |
                                      (~grade.isin(["F", "AB"]) & m.status.ne("Pass")))
    out.append(_mk("Grade agrees with pass/fail status", "error", m, bad_grade_status,
                   lambda r: f"{r.enroll} {r.code}: grade {r.grade} but status {r.status}"))
    g = m[m.credits > 0].groupby("grade").gpoints.nunique()
    out.append(Check("One grade-point value per grade", "error", len(g), int((g > 1).sum()), [f"grade {k}" for k in g[g > 1].index[:3]]))
    m2 = m.assign(gp=m.credits.fillna(0) * m.gpoints.fillna(0), ce=m.earned.fillna(0))
    agg = m2.groupby(["enroll", "term"]).agg(gp=("gp", "sum"), ce=("ce", "sum")).reset_index()
    j = t.merge(agg, on=["enroll", "term"], how="inner").dropna(subset=["grade_total"])
    out.append(_mk("Grade total = sum(credits x grade points)", "error", j, (j.gp - j.grade_total).abs() > 0.51,
                   lambda r: f"{r.enroll} term {r.term}: computed {r.gp:g}, sheet {r.grade_total:g}"))
    j2 = j.dropna(subset=["sgpa", "cr_reg"]); j2 = j2[j2.cr_reg > 0]
    out.append(_mk("SGPA = grade total / credits registered", "error", j2, (j2.grade_total / j2.cr_reg - j2.sgpa).abs() > 0.011,
                   lambda r: f"{r.enroll} term {r.term}: computed {r.grade_total / r.cr_reg:.2f}, sheet {r.sgpa}"))
    j3 = j.dropna(subset=["cr_earned"])
    out.append(_mk("Credits earned = sum of subject credits earned", "error", j3, (j3.ce - j3.cr_earned).abs() > 0.51,
                   lambda r: f"{r.enroll} term {r.term}: computed {r.ce:g}, sheet {r.cr_earned:g}"))
    last = t.term_n.max(); lm = m[m.term == t[t.term_n == last].term.iloc[0]]
    bad = set(lm[lm.status != "Pass"].enroll)
    if "status" in cum:
        c = cum.dropna(subset=["status"])
        out.append(_mk("Final 'Pass' students have no failed subject in last term", "error", c,
                       (c.status == "Pass") & c.enroll.isin(bad), lambda r: f"{r.enroll} marked Pass but has a failed subject"))
    d = s[s.credits.isna() | (s.enrolled == 0)]
    out.append(Check("Every subject has credits and enrolled students", "info", len(s), len(d), list(d.label.head(3))))
    return out

def render_report(checks):
    lines = []
    for c in checks:
        mark = "OK  " if c.ok else ("FAIL" if c.level == "error" else "NOTE")
        lines.append(f"[{mark}] {c.name}: {c.failed} of {c.checked} rows off" + (f"  e.g. {'; '.join(map(str, c.examples))}" if c.examples else ""))
    return "\n".join(lines)
