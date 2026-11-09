"""Step 4 - saved layouts: a staff member confirms an unfamiliar sheet once, the choice is stored as a small JSON file
(next to the app, in layouts/) and reused automatically whenever a sheet with the same header shape is uploaded.
No code change or update is needed for new department formats."""
import hashlib, json, os, re
from pathlib import Path
from .u0_common import _t, _is_name

LAYOUT_DIR = Path(os.environ.get("GRADESENSE_LAYOUTS", Path(__file__).resolve().parent.parent / "layouts"))
_INTKEYS = ("subject_names", "max_marks", "components")

def signature(raw, k):
    """Fingerprint of the first k rows (the header rows: positions + texts) and the column count; independent of student data."""
    cells = sorted(f"{r},{c}:{re.sub(chr(32) + '+', chr(32), _t(v)).lower()}" for r, row in enumerate(raw.iloc[:k].to_numpy()) for c, v in enumerate(row) if _t(v))
    return hashlib.sha1((f"{raw.shape[1]}|{k}|" + "|".join(cells)).encode()).hexdigest()[:12]

def _file(raw, k): return LAYOUT_DIR / f"{signature(raw, k)}.json"
def save_layout(raw, overrides, name=""):
    LAYOUT_DIR.mkdir(parents=True, exist_ok=True)
    k = int(overrides.get("first_row", 0)); doc = {"name": name, "header_rows": k, "signature": signature(raw, k), "overrides": {k: ({str(a): b for a, b in v.items()} if k in _INTKEYS else v) for k, v in overrides.items()}}
    _file(raw, k).write_text(json.dumps(doc, indent=1, default=int))
    return doc["signature"]

def load_layout(raw):
    """Most specific saved layout whose header rows match this sheet exactly (tries 15 header rows down to none)."""
    for k in range(min(15, len(raw)), -1, -1):
        f = _file(raw, k)
        if f.exists():
            o = json.loads(f.read_text())["overrides"]
            return {kk: ({int(a): b for a, b in v.items()} if kk in _INTKEYS else v) for kk, v in o.items()}
    return None

def delete_layout(raw):
    for k in range(min(15, len(raw)), -1, -1):
        if _file(raw, k).exists(): _file(raw, k).unlink(); return True
    return False
