"""Stage 2 - offline syllabus reader (no API, no LLM): teaching-scheme table image/PDF -> catalog draft.
Pipeline: find table grid lines (OpenCV) -> OCR each cell (Tesseract) -> parse rows -> verify against the
table's own printed Total row. Anything that doesn't reconcile is flagged for the review step."""
import io, os, re, shutil
from dataclasses import dataclass, field
import cv2, numpy as np, pytesseract
from .ingest import UnsupportedLayout, parse_code

def _find_tesseract():
    """Use the tesseract on PATH, else $TESSERACT_CMD, else the usual Windows/macOS install folders."""
    if shutil.which("tesseract"): return
    for c in (os.environ.get("TESSERACT_CMD", ""), r"C:\Program Files\Tesseract-OCR\tesseract.exe",
              r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe", "/opt/homebrew/bin/tesseract", "/usr/local/bin/tesseract"):
        if c and os.path.exists(c): pytesseract.pytesseract.tesseract_cmd = c; return
_find_tesseract()

COLS = ["code", "type", "name", "credits", "L", "P", "T", "ISE", "MSE", "ESE", "INT", "OE_POE", "total"]
NUM = {"L", "P", "T", "ISE", "MSE", "ESE", "INT", "OE_POE"}
PARTS = ["ISE", "MSE", "ESE", "INT", "OE_POE"]
SUMS = ["credits", "L", "P", "T", "ISE", "MSE", "ESE", "INT", "OE_POE", "total"]

@dataclass
class Scheme:
    semester: str
    courses: list
    printed_totals: dict
    pass_pct: int
    problems: list = field(default_factory=list)      # human-readable review items
    def to_catalog(self):
        return {"semester": self.semester, "min_pass_rule": f"{self.pass_pct}% of total marks of the individual course" if self.pass_pct else "",
                "total_credits": self.printed_totals.get("credits"), "total_marks": self.printed_totals.get("total"), "courses": self.courses}

def _to_gray(data):
    if data[:4] == b"%PDF":
        import pdfplumber
        with pdfplumber.open(io.BytesIO(data)) as pdf:
            arr = np.array(pdf.pages[0].to_image(resolution=220).original.convert("L"))
        return arr
    img = cv2.imdecode(np.frombuffer(data, np.uint8), cv2.IMREAD_GRAYSCALE)
    if img is None: raise UnsupportedLayout("Could not decode the syllabus file as an image or PDF.")
    return img

def _clusters(vals, gap=6):
    if len(vals) == 0: return []
    vals = sorted(vals); out = [[vals[0]]]
    for v in vals[1:]: (out[-1].append(v) if v - out[-1][-1] <= gap else out.append([v]))
    return [int(np.mean(c)) for c in out]

def _lines(g):
    H, W = g.shape
    _, bw = cv2.threshold(g, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)
    hor = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (max(20, W // 40), 1)))
    ver = cv2.morphologyEx(bw, cv2.MORPH_OPEN, cv2.getStructuringElement(cv2.MORPH_RECT, (1, max(12, H // 40))))
    return hor, ver

def _deskew(g, force=False):
    hor, _ = _lines(g)
    ln = cv2.HoughLinesP(hor, 1, np.pi / 1440, threshold=80, minLineLength=g.shape[1] // 3, maxLineGap=25)
    if ln is None or len(ln) < 3: return g
    ln = np.asarray(ln).reshape(-1, 4)
    a = np.degrees(np.arctan2(ln[:, 3] - ln[:, 1], ln[:, 2] - ln[:, 0]))
    ang = float(np.median(a))
    if (a.std() > 0.4 and not force) or abs(ang) < (0.05 if force else 0.3) or abs(ang) > 8: return g          # noisy estimate or already straight
    H, W = g.shape
    return cv2.warpAffine(g, cv2.getRotationMatrix2D((W / 2, H / 2), ang, 1.0), (W, H), flags=cv2.INTER_CUBIC, borderValue=255)

def _bbox(hor, ver):
    cs, _ = cv2.findContours(cv2.dilate(cv2.bitwise_or(hor, ver), np.ones((3, 3), np.uint8)), cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
    if not cs: raise UnsupportedLayout("No table grid found in the syllabus image.")
    return cv2.boundingRect(max(cs, key=cv2.contourArea))

def _is_horizontal_dash(binary):
    points = np.column_stack(np.where(binary < 128))
    if not len(points): return False
    width = int(points[:, 1].max() - points[:, 1].min() + 1)
    height = int(points[:, 0].max() - points[:, 0].min() + 1)
    return 4 <= width <= binary.shape[1] * 0.5 and height <= max(3, width * 0.4)

def _measure(g):
    hor, ver = _lines(g); x, y, w, h = _bbox(hor, ver)
    vx = _clusters(np.where(ver[y:y + h, x:x + w].sum(axis=0) > 255 * h * 0.12)[0] + x)
    hy = _clusters(np.where(hor[y:y + h, x:x + w].sum(axis=1) > 255 * w * 0.12)[0] + y)
    return (g, hor, vx, hy, (x, y, w, h)), len(vx) - 1

def _grid(g, force=False):
    if g.shape[1] < 1600: g = cv2.resize(g, None, fx=1600 / g.shape[1], fy=1600 / g.shape[1], interpolation=cv2.INTER_CUBIC)
    g = _deskew(g, force)
    res, seen = _measure(g)                                  # attempt 1: as given (best for clean screenshots)
    if seen in (len(COLS), len(COLS) + 1): return res
    x, y, w, h = res[4]                                      # attempt 2: crop to the table and normalise its scale
    if w < 300 or h < 150: raise UnsupportedLayout("No table grid found in the syllabus image.")
    g = g[max(0, y - 8):y + h + int(0.16 * h), max(0, x - 8):x + w + 8]
    for width in (1150, 1350, 1000):
        res, seen = _measure(cv2.resize(g, None, fx=width / g.shape[1], fy=width / g.shape[1], interpolation=cv2.INTER_CUBIC))
        if seen in (len(COLS), len(COLS) + 1): return res
    raise UnsupportedLayout(f"Found {seen} table columns, expected {len(COLS)} (Code, Type, Name, Credits, L, P, T, ISE, MSE, ESE, INT, OE/POE, Total). "
                            "Use a straight, sharp image of the full table.")

def _ocr(g, x0, x1, y0, y1, col, retried):
    crop = g[y0 + 3:y1 - 2, x0 + 3:x1 - 2]
    if crop.size == 0: return None
    bin_ = cv2.threshold(crop, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)[1]
    if (bin_ < 128).mean() < 0.0015: return None                                    # empty cell: skip OCR
    if col in NUM and _is_horizontal_dash(bin_): return "-"
    if np.median(crop) < 215: crop = bin_                                           # shaded cell (e.g. blue Total row): flatten it
    big = cv2.copyMakeBorder(cv2.resize(crop, None, fx=3, fy=3, interpolation=cv2.INTER_CUBIC), 12, 12, 12, 12, cv2.BORDER_CONSTANT, value=255)
    wl = lambda s: f"-c tessedit_char_whitelist={s}"
    cfg = {"code": f"--psm 7 {wl('ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789')}", "credits": "--psm 7", "total": "--psm 7"}.get(col)
    if col in NUM: cfg = f"--psm 7 {wl('0123456789-*')}"
    txt = " ".join(pytesseract.image_to_string(big, config=cfg or "--psm 6").split())
    if not txt and col in NUM:            # a lone thin '1' is read as l | [ ! unless unrestricted; map those back
        t = "".join(pytesseract.image_to_string(big, config="--psm 10").split()).translate(str.maketrans("lI|![]i/", "11111111"))
        if re.fullmatch(r"[\d\-*]+", t): txt = t; retried.add(t)
    return txt

def _ocr_cells(g, cells, col):
    from concurrent.futures import ThreadPoolExecutor

    def read(cell):
        retries = set()
        text = _ocr(g, *cell, col, retries)
        return text, retries

    with ThreadPoolExecutor(max_workers=4) as pool:
        output = list(pool.map(read, cells))
    return [text for text, _ in output], set().union(*(retries for _, retries in output))

def _num(s):
    s = (s or "").strip()
    if s in ("", "-"): return None, False
    m = re.fullmatch(r"(\d+)(\*?)", s)
    return (int(m.group(1)), bool(m.group(2))) if m else (s, False)

def _fix_code(s):
    s = re.sub(r"[^A-Z0-9]", "", s.upper())
    if len(s) < 8: return s
    d = {"O": "0", "I": "1", "L": "1", "S": "5", "B": "8", "Z": "2"}
    return "".join(d.get(c, c) for c in s[:3]) + s[3:-3] + "".join(d.get(c, c) for c in s[-3:])

def _is_course_code(code):
    return bool(re.fullmatch(r"\d{3}[A-Z][A-Z0-9]{5,}", code))

def _is_leading_serial_column(g, columns, rows, retried):
    widths = [b - a for a, b in columns]
    if len(widths) != len(COLS) + 1 or widths[0] >= widths[1] * 0.5:
        return False
    code_cells = [(columns[1][0], columns[1][1], y0, y1) for y0, y1 in rows]
    serial_cells = [(columns[0][0], columns[0][1], y0, y1) for y0, y1 in rows]
    code_values, _ = _ocr_cells(g, code_cells, "code")
    serial_values, _ = _ocr_cells(g, serial_cells, "name")
    codes = [_fix_code(text or "") for text in code_values]
    serials = []
    for code, text in zip(codes, serial_values):
        if not _is_course_code(code):
            continue
        match = re.search(r"\d+", text or "")
        serials.append(int(match.group()) if match else None)
    values = [value for value in serials if value is not None]
    consecutive = sum(right == left + 1 for left, right in zip(values, values[1:]))
    return len(values) >= 4 and consecutive >= 2 and consecutive / (len(values) - 1) >= 0.5

def extract_scheme(data: bytes) -> Scheme:
    gray = _to_gray(data)
    clahe = cv2.createCLAHE(clipLimit=2.4, tileGridSize=(8, 8)).apply(gray)
    adaptive = cv2.adaptiveThreshold(
        cv2.GaussianBlur(gray, (3, 3), 0), 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
        cv2.THRESH_BINARY, 35, 11)
    sharpened = cv2.addWeighted(clahe, 1.6, cv2.GaussianBlur(clahe, (0, 0), 1.4), -0.6, 0)
    candidates = [gray, clahe, adaptive, sharpened]
    best, best_score, failures = None, None, []
    for candidate in candidates:
        try:
            result = _extract_image(candidate, False)
        except UnsupportedLayout as ex:
            failures.append(str(ex))
            continue
        missing = sum(not isinstance(course.get(field), int) for course in result.courses for field in SUMS if field != "credits")
        score = (len(result.problems), missing, -len(result.courses))
        if best_score is None or score < best_score:
            best, best_score = result, score
        if score[:2] == (0, 0) or not best.problems:
            return best
    if best is not None:
        if best.problems:
            best.problems.insert(0, "Review the recognized values before saving; OCR may be uncertain.")
        return best
    reason = failures[-1] if failures else "Could not recognize a syllabus table."
    raise UnsupportedLayout(f"{reason} Try a clearer or closer image, or continue analysis using the result workbook without a syllabus.")

def _extract_image(image, force):
    g, hor, vx, hy, (tx, ty, tw, th) = _grid(image, force)
    cb, rb, retried = list(zip(vx[:-1], vx[1:])), list(zip(hy[:-1], hy[1:])), set()
    offset = 0
    if len(cb) == len(COLS) + 1:
        if not _is_leading_serial_column(g, cb, rb, retried):
            raise UnsupportedLayout("Found an extra leading table column, but it could not be validated as sequential row numbers.")
        offset = 1
    elif len(cb) != len(COLS):
        raise UnsupportedLayout(f"Found {len(cb)} table columns; expected {len(COLS)} or those columns plus a leading row-number column.")
    def has_line(ci, y):
        a, b = cb[ci]; seg = hor[y - 2:y + 3, a + 3:b - 3]
        return seg.size > 0 and seg.max(axis=0).mean() > 115
    rows, span, unread = {}, {}, set()
    for ci, col in enumerate(COLS):                     # OCR each cell; cells merged across rows are read once
        ci += offset
        i = 0
        cells, groups = [], []
        while i < len(rb):
            j = i
            while j + 1 < len(rb) and not has_line(ci, rb[j][1]): j += 1
            cells.append((cb[ci][0], cb[ci][1], rb[i][0], rb[j][1]))
            groups.append((i, j))
            i = j + 1
        values, retried_cells = _ocr_cells(g, cells, col)
        retried.update(retried_cells)
        for (i, j), txt in zip(groups, values):
            if not txt and col in NUM and i == j: unread.add((i, col))
            for k in range(i, j + 1): rows.setdefault(k, {})[col] = txt or ""; span[(k, col)] = i
    head = pytesseract.image_to_string(g[max(0, ty - 4):ty + 30, tx:tx + tw], config="--psm 7")
    sem = re.search(r"SEMESTER[-\s]*([IVX]+|\d+)", head.upper().replace("|", "I"))
    foot = pytesseract.image_to_string(g[ty + th:ty + th + 90, tx:], config="--psm 6")
    pct = re.search(r"(\d{2})\s*%", foot)
    courses, printed, problems = [], {}, []
    for k in sorted(rows):
        r = rows[k]; code = _fix_code(r["code"])
        name = re.sub(r"[\'\"`|]", "", re.sub(r"^\s*\d+[.)]\s*", "", r["name"])).strip()
        if _is_course_code(code):
            c = {"code": code, "name": name, "_k": k, "_g": span[(k, "credits")]}
            typ = re.sub(r"\s*\(.*", "", r["type"]).strip()
            c["type"] = typ if 0 < len(typ) <= 10 else parse_code(code)[1]
            for f in COLS[3:]:
                v, star = _num(r[f])
                if v is not None: c[f] = v
                if star: c.setdefault("not_in_total", []).append(f)
            courses.append(c)
        elif name.lower().startswith("total"):
            printed = {f: _num(r[f])[0] for f in SUMS if isinstance(_num(r[f])[0], int)}
    if not courses: raise UnsupportedLayout("Table found, but no course-code rows (e.g. 231AIMLPCC204) could be read.")
    for c in courses:                                    # each row must add up
        if isinstance(c.get("total"), int):
            parts = sum(c.get(p, 0) for p in PARTS if isinstance(c.get(p), int))
            if parts != c["total"]: problems.append(f"{c['code']} {c['name']}: components add to {parts} but total says {c['total']}")
    def colsum(f):
        got, groups = 0, set()
        for c in courses:
            v = c.get(f)
            if not isinstance(v, int) or f in c.get("not_in_total", []): continue
            if f == "credits":                           # a merged credits cell is counted once
                if c["_g"] in groups: continue
                groups.add(c["_g"])
            got += v
        return got
    gap = {f: printed[f] - colsum(f) for f in SUMS if f in printed}
    off = {f: d for f, d in gap.items() if d}
    repaired = []
    if len(off) == 1 and not problems:                   # safe repair: ONE column is off, every row adds up, ONE unreadable cell
        (f, d), = off.items()
        cands = [c for c in courses if (c["_k"], f) in unread and f not in c]
        if f in NUM and 1 <= d <= 9 and len(cands) == 1:
            cands[0][f] = d; repaired.append(f"{cands[0]['code']} {f}={d}"); off = {}
    for f, d in off.items():
        problems.append(f"Column {f}: rows add to {printed[f] - d} but the printed Total is {printed[f]}")
    if not printed: problems.append("No printed Total row found, so numbers could not be cross-checked.")
    if repaired: problems.append("Filled from the printed Total row (please confirm): " + ", ".join(repaired))
    if retried: problems.append(f"{len(retried)} thin-digit cell(s) needed a second OCR pass (values {sorted(retried)}); confirm them.")
    shared = {}
    for c in courses: shared[c["_g"]] = shared.get(c["_g"], 0) + 1
    for c in courses:
        if shared[c["_g"]] > 1 and isinstance(c.get("credits"), int): c["credits_shared_by"] = shared[c["_g"]]
        c.pop("_k"); c.pop("_g")
    return Scheme(sem.group(1) if sem else "?", courses, printed, int(pct.group(1)) if pct else 0, problems)
