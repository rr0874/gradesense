"""python -m backend <workbook.xlsx> [catalog.json]  -> prints what was extracted and the validation report."""
import sys, warnings
warnings.filterwarnings("ignore")
from . import Dataset, run_checks, render_report, UnsupportedLayout
from .catalog import load, apply_catalog

path = sys.argv[1] if len(sys.argv) > 1 else "data/AIML_4.xlsx"
try:
    ds = Dataset.from_workbook(path, path)
except UnsupportedLayout as e:
    sys.exit(f"Cannot read this workbook: {e}")
i = ds.wb.info
print(f"{path}: {i['students']} students, terms {i['terms']}, {i['subjects']} subject columns\n")
subs = ds.wb.subjects
if len(sys.argv) > 2 and sys.argv[2].lower().endswith((".png", ".jpg", ".jpeg", ".pdf")):
    import json
    from .syllabus import extract_scheme
    sch = extract_scheme(open(sys.argv[2], "rb").read()); cat = sch.to_catalog()
    print(f"syllabus: semester {sch.semester}, {len(sch.courses)} courses read, passing rule {sch.pass_pct}%")
    for p in sch.problems: print("  REVIEW:", p)
    out = f"backend/catalogs/sem{sch.semester}.json"; json.dump(cat, open(out, "w"), indent=1); print("  saved", out, "\n")
    subs, miss = apply_catalog(subs, cat); print(f"catalog matched {subs.in_catalog.sum()} subject columns; unmatched (e.g.): {miss[:6]}\n")
elif len(sys.argv) > 2:
    subs, miss = apply_catalog(subs, load(sys.argv[2])); print(f"catalog matched {subs.in_catalog.sum()} subject columns; unmatched (e.g.): {miss[:6]}\n")
for term, g in subs.groupby("term_n", sort=True):
    common = g[g.is_common]; opt = g[~g.is_common]
    print(f"Term {g.term.iloc[0]}: {len(common)} common subjects ({', '.join(common.label)}); {len(opt)} optional/elective ({', '.join(opt.label.head(6))}{'...' if len(opt) > 6 else ''})")
print("\nValidation:\n" + render_report(run_checks(ds)))
