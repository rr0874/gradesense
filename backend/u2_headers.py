"""Step 2 - read the header band: subject name, component (ISE/MSE/ESE...) and max marks for every mark column."""
from .ingest import _norm
from .u0_common import *

def _labels(t, band, cols, overrides=None):
    """Subject / component / max for every mark column from the header band; merged group headers are forward-filled."""
    grid = {r: {} for r in band}
    for r in band:
        last = ""
        for c in range(cols[0], cols[-1] + 1):
            last = t.iat[r, c] or last
            grid[r][c] = t.iat[r, c] or last
    out = {}
    for c in cols:
        texts = [grid[r][c] for r in band if grid[r][c]]            # top -> bottom, nearest data last
        subj = next((x for x in reversed(texts) if not GENERIC.match(x) and not SKIP.search(x) and not re.fullmatch(r"[\W\d]+", x)), "")
        comp = next((COMP[_norm(x).split()[0]] for x in reversed(texts) if _norm(x) and _norm(x).split()[0] in COMP), "mse")
        mx = next((float(m.group(1)) for x in reversed(texts) for m in [re.search(r"\(\s*(\d+)\s*\)|out of (\d+)", x)] if m for m in [m] if m.group(1)), float("nan"))
        out[c] = (re.sub(r"\s*\(\s*\d+\s*\)\s*$", "", subj).strip(), comp, mx)
    o = overrides or {}                                              # user-typed subject names / max marks win
    for c, (s_, k_, m_) in list(out.items()):
        out[c] = (o.get("subject_names", {}).get(c, s_), o.get("components", {}).get(c, k_), o.get("max_marks", {}).get(c, m_))
    return out

