"""GradeSense - chat UI (Stage 5).  Run:  streamlit run app.py
Fully offline: no API or external AI service is called anywhere."""
import hashlib, io, json, re, time, traceback
from importlib import reload
from pathlib import Path
import pandas as pd
import streamlit as st
import backend.mse_chat as mse_chat
from backend import Dataset, UnsupportedLayout, run_checks
from backend.catalog import CAT_DIR
from backend.engine import Answer, Engine
from backend.analysis_request import AnalysisRequestError, extract_analysis_request
from backend.mse import analyze_mse, validate_mse_dataset
from backend.mse_chat import answer_division_mse, answer_single_mse
from backend.mse_divisions import read_division_mse_workbook
from backend.syllabus import extract_scheme
from backend.report import build_generic_report, build_mse_analysis_ii, build_mse_report, build_report
from backend.ui_mapping import layout_wizard

ROOT = Path(__file__).parent; DATA = ROOT / "data"
SAMPLE_XL, SAMPLE_IMG = DATA / "AIML_4.xlsx", DATA / "scheme_sem4.png"
SCHEMES = ROOT / "backend" / "report_schemes"        # one sem<ROMAN>.json per semester for the Excel result analysis
S = st.session_state
for k, v in {"msgs": [], "data": None, "digest": None, "fname": None, "error": None, "default_term": None, "sy_bytes": None,
             "chat_epoch": 0, "workbook_type": "semester", "show_mse_analysis": False,
             "upload_workbook_type": "semester", "mse_kind": "single",
             "home_view": False, "wizard": None}.items(): S.setdefault(k, v)

st.set_page_config(page_title="GradeSense | Academic Results", page_icon="🎓", layout="wide", initial_sidebar_state="expanded")
st.markdown("""<style>
#MainMenu, footer {visibility: hidden}
.block-container {max-width: 1240px; padding-top: 2rem; padding-bottom: 7rem}
[data-testid="stSidebar"] {border-right: 1px solid #e1e7ef}
[data-testid="stSidebar"] [data-testid="stVerticalBlock"] {gap: .75rem}
.gs-brand {display:flex; align-items:center; gap:.8rem; margin: .2rem 0 1.6rem}
.gs-mark {width: 42px; height: 42px; border-radius: 12px; display:grid; place-items:center; background:#17365d; color:#fff; font-size:1.2rem}
.gs-brand-title {font-size:1.05rem; font-weight:700; color:#17365d; letter-spacing:.01em}
.gs-brand-sub {font-size:.72rem; color:#64748b; margin-top:.08rem}
.gs-hero {border-bottom:1px solid #e2e8f0; padding:.15rem 0 1.25rem; margin-bottom:1.35rem}
.gs-eyebrow {color:#a77c24; text-transform:uppercase; font-size:.72rem; font-weight:700; letter-spacing:.13em}
.gs-title {font-size:2rem; line-height:1.2; font-weight:700; color:#17365d; margin:.3rem 0}
.gs-sub {color:#64748b; font-size:.98rem; line-height:1.55}
.gs-card {border:1px solid #e1e7ef; border-radius:12px; padding:.8rem 1rem; background:#fff; font-size:.88rem; line-height:1.5}
.gs-metric {border:1px solid #e1e7ef; border-radius:12px; padding:.95rem 1.05rem; background:#fff; min-height:98px}
.gs-metric-label {font-size:.76rem; font-weight:600; color:#64748b; text-transform:uppercase; letter-spacing:.055em}
.gs-metric-value {font-size:1.48rem; font-weight:700; color:#17365d; margin-top:.32rem}
.gs-section {font-size:1.12rem; font-weight:700; color:#17365d; margin:.4rem 0 .2rem}
.gs-chat-intro {font-size:1rem; font-weight:650; color:#17365d; margin:.7rem 0 .15rem}
[data-testid="stChatMessage"] {max-width:940px; border-radius:12px; padding:.9rem 1rem; margin:0 auto .65rem; border:1px solid #e1e7ef}
[data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) {max-width:76%; margin-right:max(0px, calc((100% - 940px) / 2)); margin-left:auto; background:#f2f6fb; border-color:#e1e7ef}
[data-testid="stChatInput"] {max-width:940px; margin-left:auto; margin-right:auto; border-radius:12px}
.gs-workspace-title {font-size:1.12rem; font-weight:700; color:#17365d}
.gs-workspace-context {font-size:.8rem; color:#64748b; margin-top:.18rem}
div.stButton > button {border-radius:9px; min-height:2.65rem; border:1px solid #cbd5e1; font-weight:600; transition:all .15s ease}
div.stButton > button:hover {border-color:#17365d; color:#17365d; background:#f2f6fb}
div[data-testid="stDownloadButton"] > button {border-radius:9px; font-weight:600}
div[data-testid="stExpander"] {border-color:#e1e7ef; border-radius:10px}
@media (max-width: 700px) {
 .block-container {padding-top:1.2rem}
 .gs-title {font-size:1.65rem}
 .gs-metric {min-height:84px}
 [data-testid="stChatMessage"]:has([data-testid="stChatMessageAvatarUser"]) {max-width:90%}
}
</style>""", unsafe_allow_html=True)

# ---------------------------------------------------------------- cached backend
@st.cache_resource(show_spinner="Reading the result sheet…")
def get_dataset(digest: str, workbook_type: str, _data: bytes):
    return Dataset.from_workbook(_data, digest, include_mse_rows=workbook_type == "mse")
@st.cache_resource(show_spinner="Reading division MSE sheets…")
def get_division_mse(digest: str, filename: str, _data: bytes):
    return read_division_mse_workbook(_data, filename)
@st.cache_resource(show_spinner="Preparing the assistant…")
def get_engine(digest: str, sig: str, workbook_type: str, _ds): return Engine(_ds, log_path=DATA / "unmatched_queries.csv")
@st.cache_resource(show_spinner=False)
def get_checks(digest: str, workbook_type: str, _ds): return run_checks(_ds)
@st.cache_data(show_spinner=False)
def read_scheme(data: bytes):
    try: return extract_scheme(data), None
    except UnsupportedLayout as ex: return None, str(ex)
    except Exception as ex:
        if ex.__class__.__name__ == "TesseractNotFoundError":
            return None, "Local Tesseract OCR is not installed or configured. Install Tesseract to read syllabus images; result-workbook analysis remains available without it."
        return None, f"Could not read this file ({ex.__class__.__name__})."
@st.cache_data(show_spinner="Building the analysis workbook…")
def make_report(digest: str, scheme_path: str, mtime: int, _ds):
    b = io.BytesIO(); build_report(_ds, json.loads(Path(scheme_path).read_text()), b); return b.getvalue()
@st.cache_data(show_spinner="Preparing the validated analysis workbook…")
def make_generic_report(digest: str, workbook_type: str, term: str, _ds):
    b = io.BytesIO(); build_generic_report(_ds, term, b); return b.getvalue()
@st.cache_data(show_spinner="Building the MSE analysis workbook…")
def make_mse_report(digest: str, term: str, _ds):
    b = io.BytesIO(); build_mse_report(_ds, analyze_mse(_ds, term), b); return b.getvalue()
@st.cache_data(show_spinner="Building the MSE Analysis II workbook…")
def make_mse_analysis_ii_report(digest: str, filename: str, source_data: bytes, _analysis):
    b = io.BytesIO(); build_mse_analysis_ii(_analysis, b, source_data); return b.getvalue()
def cat_sig(): return "|".join(f"{p.name}:{p.stat().st_mtime_ns}" for p in sorted(CAT_DIR.glob("*.json")))

WORKBOOK_TYPES = {"semester": "Whole-semester result", "mse": "MSE result"}

def load(data: bytes, name: str, workbook_type: str = "semester"):
    digest = hashlib.sha1(data).hexdigest()[:12]
    if digest == S.digest and workbook_type == S.workbook_type:
        if S.home_view:
            S.home_view = False
            st.rerun()
        return
    mse_kind = "single"
    try:
        if workbook_type == "mse":
            try:
                get_division_mse(digest, name, data)
                mse_kind = "divisions"
            except UnsupportedLayout:
                ds = get_dataset(digest, workbook_type, data)
                validate_mse_dataset(ds)
        else:
            get_dataset(digest, workbook_type, data)
    except UnsupportedLayout as ex:
        hint = (" If this is an MSE-only register, select **MSE result** and press **Analyze workbook**."
                if workbook_type == "semester" and "No 'Term - X' markers" in str(ex) else "")
        S.error = f"**{name}** can't be read yet: {ex}{hint}"
        if workbook_type == "mse": S.wizard = {"data": data, "name": name}     # let staff confirm the columns once
        return
    except Exception as ex:
        traceback.print_exc()
        S.error = f"**{name}** couldn't be opened ({ex.__class__.__name__}: {ex}). Check the terminal for details."
        return
    S.update(data=data, digest=digest, fname=name, workbook_type=workbook_type, mse_kind=mse_kind, error=None,
             wizard=None, msgs=[], default_term=None, show_mse_analysis=False, home_view=False)
    st.rerun()


def ask_cb(q): S.pending = q

def start_new_conversation():
    S.msgs = []
    S.pop("pending", None)
    S.chat_epoch += 1

def analysis_request_panel():
    query = None
    with st.expander("Analyze a request from an image or PDF", expanded=bool(S.get("analysis_request_digest"))):
        st.caption(
            "Upload a PNG, JPG, JPEG, or PDF containing the analysis you want. "
            "OCR and analysis run locally; review and correct the recognized text before running it."
        )
        request_file = st.file_uploader(
            "Analysis request (.png, .jpg, .jpeg, or .pdf)",
            type=["png", "jpg", "jpeg", "pdf"],
            key="analysis_request_upload",
        )
        if not request_file:
            old_digest = S.pop("analysis_request_digest", None)
            if old_digest:
                S.pop(f"analysis_request_text_{old_digest}", None)
            return query

        request_bytes = request_file.getvalue()
        digest = hashlib.sha1(request_bytes).hexdigest()[:12]
        previous_digest = S.get("analysis_request_digest")
        if previous_digest and previous_digest != digest:
            S.pop(f"analysis_request_text_{previous_digest}", None)
            S.pop("analysis_request_digest", None)
            S.pop("analysis_request_error", None)

        if st.button("Read request text", key=f"read_analysis_request_{digest}"):
            try:
                S[f"analysis_request_text_{digest}"] = extract_analysis_request(
                    request_bytes, request_file.name
                )
                S.analysis_request_digest = digest
                S.pop("analysis_request_error", None)
            except AnalysisRequestError as ex:
                S.analysis_request_error = str(ex)
            except Exception as ex:
                S.analysis_request_error = f"Could not read this request file ({ex.__class__.__name__})."
        if S.get("analysis_request_error"):
            st.error(S.analysis_request_error)
        if S.get("analysis_request_digest") == digest:
            text_key = f"analysis_request_text_{digest}"
            st.text_area(
                "Recognized text — review or correct before analysis",
                key=text_key,
                height=170,
            )
            if st.button("Run analysis request", type="primary", key=f"run_analysis_request_{digest}"):
                query = S.get(text_key, "").strip()
                if not query:
                    st.warning("Enter or correct the request text before running it.")
    return query

def workspace_header(workspace_name):
    title, actions = st.columns([5, 3])
    title.markdown(
        f'<div class="gs-workspace-title">GradeSense <span style="color:#94a3b8;font-weight:400">/</span> '
        f'{workspace_name}</div><div class="gs-workspace-context">{S.fname} &nbsp;·&nbsp; '
        f'{"MSE" if S.workbook_type == "mse" else "Results"}</div>',
        unsafe_allow_html=True,
    )
    home, reset = actions.columns(2)
    home.button("Home", key="workspace_home", on_click=lambda: S.update(home_view=True))
    reset.button("New conversation", key="new_conversation", on_click=start_new_conversation)

def mse_chat_view(analysis, answer_function, examples):
    st.markdown(
        '<div class="gs-hero"><div class="gs-eyebrow">MSE assistant</div>'
        '<div class="gs-title">Chat about these results</div>'
        '<div class="gs-sub">Ask about pass/fail/absent results, subjects, divisions, or individual students. '
        'Answers are calculated from the active workbook.</div></div>',
        unsafe_allow_html=True,
    )
    if not S.msgs:
        st.markdown('<div class="gs-chat-intro">Try an MSE analysis</div>', unsafe_allow_html=True)
        prompt_cols = st.columns(2)
        for index, prompt in enumerate(examples):
            prompt_cols[index % 2].button(
                prompt, key=f"mse_prompt_{index}", on_click=ask_cb, args=(prompt,), width="stretch"
            )

    request_query = analysis_request_panel()
    chat_query = st.chat_input("Ask a question about these MSE results…", key=f"mse_chat_input_{S.chat_epoch}")
    pending_query = S.pop("pending", None)
    question = chat_query or pending_query or request_query
    if question:
        S.msgs.append({"role": "user", "text": question})
        try:
            if request_query is not None and not chat_query and not pending_query:
                reload(mse_chat)
                answer = mse_chat.answer_mse_request(analysis, question, answer_function)
            else:
                answer = answer_function(analysis, question)
        except (ValueError, KeyError) as ex:
            answer = Answer(f"Couldn't answer that MSE question: {ex}")
        S.msgs.append({"role": "assistant", "text": answer.text, "table": answer.table, "chart": None})

    for index, message in enumerate(S.msgs):
        with st.chat_message(message["role"], avatar="🎓" if message["role"] == "assistant" else None):
            st.markdown(message["text"])
            table = message.get("table")
            if table is not None:
                st.dataframe(table, hide_index=True, width="stretch")
                excel, csv, _ = st.columns([1, 1, 5])
                excel_data = io.BytesIO()
                table.to_excel(excel_data, index=False)
                excel.download_button(
                    "Download Excel table", excel_data.getvalue(), f"mse_results_{index}.xlsx",
                    key=f"mse_chat_excel_{S.chat_epoch}_{index}",
                )
                csv.download_button(
                    "Download CSV table", table.to_csv(index=False).encode(),
                    f"mse_results_{index}.csv", key=f"mse_chat_csv_{S.chat_epoch}_{index}",
                )

def mse_workspace(ds):
    workspace_header("MSE results workspace")
    terms = list(ds.wb.terms.drop_duplicates("term").sort_values("term_n").term)
    change_type_key = f"change_workbook_type_{S.digest}"
    with st.sidebar:
        st.markdown('<div class="gs-brand"><div class="gs-mark">GS</div><div><div class="gs-brand-title">GradeSense</div>'
                    '<div class="gs-brand-sub">ACADEMIC RESULTS INTELLIGENCE</div></div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="gs-card"><b>ACTIVE WORKBOOK</b><br>{S.fname}<br>{ds.wb.info["students"]:,} students · '
                    f'{len(ds.wb.subjects)} subjects · MSE mode</div>', unsafe_allow_html=True)
        with st.expander("✅ MSE data check passed"):
            st.write(f"{ds.wb.info['students']:,} students and {len(ds.wb.subjects)} subjects loaded.")
            st.write("MSE results are validated for the 0–30 range and supported absence markers.")
            st.write("Pass threshold: 12/30.")
        with st.expander("Change workbook", expanded=False):
            st.selectbox("Workbook type", list(WORKBOOK_TYPES), format_func=WORKBOOK_TYPES.get,
                         index=list(WORKBOOK_TYPES).index(S.get(change_type_key, S.workbook_type)),
                         key=change_type_key)
            up = st.file_uploader("Another result sheet (.xlsx)", type=["xlsx"], key="mse_xl_up")
            if up: load(up.getvalue(), up.name, S[change_type_key])
            if S.error: st.error(S.error)
        st.caption("MSE marks are analysed locally using a 30-mark maximum and a 12-mark pass threshold.")

    st.markdown('<div class="gs-brand"><div class="gs-mark">GS</div><div><div class="gs-brand-title">GradeSense</div>'
                '<div class="gs-brand-sub">ACADEMIC RESULTS INTELLIGENCE</div></div></div>'
                '<div class="gs-hero"><div class="gs-eyebrow">MSE results</div><div class="gs-title">MSE result analysis</div>'
                '<div class="gs-sub">Numeric marks below 12 are Fail; 12–30 are Pass. AB and Absent are reported separately.</div></div>',
                unsafe_allow_html=True)
    selected_term = st.selectbox("MSE result period", terms, index=len(terms) - 1, key="mse_term_selector")
    try:
        analysis = analyze_mse(ds, selected_term)
    except UnsupportedLayout as ex:
        st.error(str(ex))
        return
    with st.sidebar.expander("Download MSE result analysis", expanded=False):
        try:
            report = make_mse_report(S.digest, selected_term, ds)
            st.download_button(
                "Download validated analysis workbook", report,
                f"mse_result_analysis_{selected_term}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", width="stretch",
                key=f"download_mse_analysis_{selected_term}",
            )
        except ValueError as ex:
            st.error(str(ex))
        st.caption("Includes Result_Analysis, subjectwise, auditable Data, and Notes sheets.")
    if st.button("MSE RESULT analysis", type="primary", key="run_mse_analysis"):
        S.show_mse_analysis = True

    t = analysis.totals
    metric_cols = st.columns(4)
    for col, (label, value) in zip(metric_cols, [
        ("Students", t["students"]), ("Overall pass", t["pass"]),
        ("Overall fail", t["fail"]), ("Overall absent", t["absent"]),
    ]):
        col.markdown(f'<div class="gs-metric"><div class="gs-metric-label">{label}</div>'
                     f'<div class="gs-metric-value">{value:,}</div></div>', unsafe_allow_html=True)
    st.caption("Per-subject pass rates exclude absent and unreported results. Overall result is Fail if any subject fails; "
               "otherwise Absent if any subject is absent, Unreported if any result is missing, and Pass only when all are passed.")
    st.markdown('<div class="gs-section">Per-subject MSE outcomes</div>', unsafe_allow_html=True)
    st.dataframe(analysis.subject_summary, hide_index=True, width="stretch")
    if S.show_mse_analysis:
        st.markdown('<div class="gs-section">Overall student outcomes</div>', unsafe_allow_html=True)
        st.dataframe(analysis.student_results, hide_index=True, width="stretch")
    else:
        st.info("Select **MSE RESULT analysis** to show per-student overall outcomes.")
    mse_chat_view(
        analysis, answer_single_mse,
        ["Overall MSE summary", "Who failed?", "Who was absent?", "Show the student results"],
    )
def mse_analysis_ii_workspace(analysis):
    workspace_header("MSE results workspace")
    change_type_key = f"change_workbook_type_{S.digest}"
    with st.sidebar:
        st.markdown('<div class="gs-brand"><div class="gs-mark">GS</div><div><div class="gs-brand-title">GradeSense</div>'
                    '<div class="gs-brand-sub">ACADEMIC RESULTS INTELLIGENCE</div></div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="gs-card"><b>ACTIVE WORKBOOK</b><br>{S.fname}<br>'
                    f'{len(analysis.students):,} students · {len(analysis.subjects)} class/division subjects · MSE mode</div>',
                    unsafe_allow_html=True)
        with st.expander("✅ MSE data check passed"):
            st.write(f"{len(analysis.students):,} students and {len(analysis.records):,} subject results loaded.")
            st.write("Student identities, source marks, subject columns, and 0–30 numeric mark ranges were validated.")
            st.write("Pass threshold: 12/30. Blank and absence markers are counted as absent.")
        with st.expander("Change workbook", expanded=False):
            st.selectbox("Workbook type", list(WORKBOOK_TYPES), format_func=WORKBOOK_TYPES.get,
                         index=list(WORKBOOK_TYPES).index(S.get(change_type_key, S.workbook_type)),
                         key=change_type_key)
            up = st.file_uploader("Another result sheet (.xlsx)", type=["xlsx"], key="mse_xl_up")
            if up:
                load(up.getvalue(), up.name, S[change_type_key])
            if S.error:
                st.error(S.error)
    with st.sidebar.expander("Download MSE Analysis II", expanded=False):
        try:
            report = make_mse_analysis_ii_report(S.digest, S.fname, S.data, analysis)
            st.download_button(
                "Download MSE Analysis II workbook", report,
                f"MSE_Analysis_II_{analysis.academic_year.replace('-', '_')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary", width="stretch", key="download_mse_analysis_ii",
            )
        except ValueError as ex:
            st.error(str(ex))
        st.caption("Class/division subject summary, student results, cell-level source audit, and notes.")

    st.markdown(
        '<div class="gs-brand"><div class="gs-mark">GS</div><div><div class="gs-brand-title">GradeSense</div>'
        '<div class="gs-brand-sub">ACADEMIC RESULTS INTELLIGENCE</div></div></div>'
        '<div class="gs-hero"><div class="gs-eyebrow">MSE results · multiple divisions</div>'
        '<div class="gs-title">MSE Performance Analysis</div>'
        '<div class="gs-sub">Marks below 12 fail; marks of 12 or above pass. AB, blank, or dash marks count as absent, matching the downloadable report.</div></div>',
        unsafe_allow_html=True,
    )
    students = analysis.students
    subjects = analysis.subjects
    metrics = st.columns(4)
    for col, (label, value) in zip(metrics, [
        ("Students", len(students)),
        ("Marks passed", int(analysis.summary["Students Passed"].sum())),
        ("Marks failed", int(analysis.summary["Students Failed"].sum())),
        ("Marks absent", int(analysis.summary["Student Absent"].sum())),
    ]):
        col.markdown(
            f'<div class="gs-metric"><div class="gs-metric-label">{label}</div>'
            f'<div class="gs-metric-value">{value:,}</div></div>',
            unsafe_allow_html=True,
        )
    st.markdown('<div class="gs-section">Class and division subject summary</div>', unsafe_allow_html=True)
    st.dataframe(analysis.summary, hide_index=True, width="stretch")
    st.caption("Appeared counts numeric marks only (Passed + Failed). AB, blank, and dash marks count as absent in this report.")
    detail = analysis.records[[
        "Class", "Division", "Roll No", "Student Name", "Course",
        "MSE Marks", "Outcome", "Source Sheet", "Source Row", "Source Column",
    ]].copy()
    with st.expander("Source student marks and outcomes"):
        st.dataframe(detail, hide_index=True, width="stretch", height=min(560, 42 + len(detail) * 28))
    st.markdown('<div class="gs-section">Explore these MSE results</div>', unsafe_allow_html=True)
    mse_chat_view(
        analysis, answer_division_mse,
        ["Overall MSE summary", "Who failed in FAI?", "Who was absent in division A?", "Show results for student A01"],
    )

# ---------------------------------------------------------------- new sheet layout (confirm columns once, then remembered)
if S.wizard:
    st.markdown('<div class="gs-hero"><div class="gs-eyebrow">New sheet layout</div><div class="gs-title">Confirm the columns of this sheet</div>'
                '<div class="gs-sub">GradeSense could not read this sheet automatically. Check the detected columns, fix any that are wrong, and save. '
                'The same sheet layout will be read automatically from now on.</div></div>', unsafe_allow_html=True)
    raw = pd.read_excel(io.BytesIO(S.wizard["data"]), header=None)
    if layout_wizard(raw, "wiz"):
        d, n = S.wizard["data"], S.wizard["name"]
        S.wizard, S.error = None, None
        load(d, n, "mse")
    if st.button("Cancel", key="wiz_cancel"):
        S.wizard = None
        st.rerun()
    st.stop()

# ---------------------------------------------------------------- home and workbook upload
if S.data is None or S.home_view:
    st.markdown('<div class="gs-brand"><div class="gs-mark">GS</div><div><div class="gs-brand-title">GradeSense</div>'
                '<div class="gs-brand-sub">ACADEMIC RESULTS INTELLIGENCE</div></div></div>'
                '<div class="gs-hero"><div class="gs-eyebrow">Academic analytics</div><div class="gs-title">Student results, clearly understood.</div>'
                '<div class="gs-sub">Upload an official semester result workbook to review outcomes, identify student support needs, and prepare analysis reports.</div></div>',
                unsafe_allow_html=True)
    st.markdown('<div class="gs-section">Start an analysis</div><div class="gs-sub">Choose a result workbook in .xlsx format. Your data is processed locally on this computer.</div>',
                unsafe_allow_html=True)
    if S.data is not None:
        st.info(f"Active workbook is still loaded: **{S.fname}**. Returning to it will not re-upload or re-read the file.")
        st.button("Return to active workbook", on_click=lambda: S.update(home_view=False),
                  type="primary", key="return_to_active_workbook")
    if S.error: st.error(S.error)
    with st.form("upload_result_workbook"):
        st.radio("Workbook type", list(WORKBOOK_TYPES), format_func=WORKBOOK_TYPES.get,
                 key="upload_workbook_type", horizontal=True,
                 help="Choose how to interpret the workbook before uploading it.")
        up = st.file_uploader("Student results (.xlsx)", type=["xlsx"],
                              help="Upload a whole-semester result or an MSE workbook in a supported ERP layout.")
        submitted = st.form_submit_button("Analyze workbook", type="primary")
    if submitted:
        if up is None:
            S.error = "Choose an .xlsx workbook before continuing."
            st.rerun()
        load(up.getvalue(), up.name, S.upload_workbook_type)
        if S.wizard: st.rerun()                                   # unreadable MSE sheet -> show the column-confirmation screen
    if SAMPLE_XL.exists():
        st.button("Open demonstration workbook · AIML, 153 students", on_click=lambda: load(SAMPLE_XL.read_bytes(), SAMPLE_XL.name, "semester"),
                  type="primary", key="open_demo_workbook")
    st.markdown('<div class="gs-card"><b>Privacy and data handling</b><br>Workbooks are analysed locally; results are not sent to an external AI service. '
                'A syllabus image or PDF can optionally be added after loading results to provide full subject names and marking schemes.</div>',
                unsafe_allow_html=True)
    st.stop()

if S.workbook_type == "mse":
    if S.mse_kind == "divisions":
        division_analysis = get_division_mse(S.digest, S.fname, S.data)
        mse_analysis_ii_workspace(division_analysis)
        st.stop()
    ds = get_dataset(S.digest, S.workbook_type, S.data)
    mse_workspace(ds)
    st.stop()
ds = get_dataset(S.digest, S.workbook_type, S.data)
try:
    eng = get_engine(S.digest, cat_sig(), S.workbook_type, ds)
except ValueError as ex:
    st.error(f"This workbook can't be analyzed safely: {ex}")
    st.stop()
chk = get_checks(S.digest, S.workbook_type, ds)
errs = [c for c in chk if c.level == "error" and not c.ok]
selected_term = S.get("default_term_selector") or S.default_term
if selected_term not in eng.terms: selected_term = eng.last
term_students = eng.base(selected_term)
term_subjects = int(eng.subj[eng.subj.term == selected_term].code.nunique())

workspace_header(f"Results workspace · Semester {selected_term}")

# ---------------------------------------------------------------- sidebar
def syllabus_panel():
    up = st.file_uploader("Syllabus table (image or PDF)", type=["png", "jpg", "jpeg", "pdf"], key="sy_up")
    if SAMPLE_IMG.exists() and not up: st.button("Use the sample syllabus (Sem IV)", on_click=lambda: S.update(sy_bytes=SAMPLE_IMG.read_bytes()))
    if up:
        S.sy_bytes = up.getvalue()
    saved = sorted(p.stem for p in CAT_DIR.glob("*.json"))
    st.caption("Saved syllabi: " + (", ".join(saved) if saved else "none yet"))

def syllabus_draft_panel(data):
    with st.container(border=True):
        st.markdown('<div class="gs-section">Review syllabus draft</div>'
                    '<div class="gs-sub">Edit any OCR mistakes directly in the table, then save the reviewed values to the semester catalogue.</div>',
                    unsafe_allow_html=True)
        with st.spinner("Reading the syllabus locally…"):
            sch, err = read_scheme(data)
        if err:
            st.error(err + " Use a straight, sharp photo of the whole teaching-scheme table.")
            return
        for problem in sch.problems:
            st.warning(problem)
        if not sch.problems:
            st.success("Every row and column adds up to the table's printed totals.")
        st.caption("Click a cell to edit an OCR value. Review course codes, marks, credits, and totals before saving.")
        cols = ["code", "name", "type", "credits", "L", "P", "T", "ISE", "MSE", "ESE", "INT", "OE_POE", "total"]
        df = pd.DataFrame(sch.courses).reindex(columns=cols).astype(object).where(lambda d: d.notna(), None)
        editor_key = "ed_" + hashlib.sha1(data).hexdigest()[:8]
        edited = st.data_editor(df, num_rows="dynamic", width="stretch", key=editor_key)
        sem = st.text_input("Semester (Roman numeral)", sch.semester if sch.semester != "?" else "",
                            key="syllabus_semester").strip().upper()
        if st.button("Save reviewed syllabus", type="primary", disabled=not sem,
                     key="save_syllabus_catalogue", help="Save the corrected course table for result analysis."):
            rows = []
            for record in edited.to_dict("records"):
                course = {key: (int(value) if isinstance(value, float) and value == int(value) else value)
                          for key, value in record.items()
                          if value is not None and not (isinstance(value, float) and pd.isna(value)) and value != ""}
                if course.get("code"):
                    rows.append(course)
            cat = {"semester": sem,
                   "min_pass_rule": f"{sch.pass_pct}% of total marks of the individual course" if sch.pass_pct else "",
                   "total_credits": sch.printed_totals.get("credits"),
                   "total_marks": sch.printed_totals.get("total"), "courses": rows}
            try:
                (CAT_DIR / f"sem{sem}.json").write_text(json.dumps(cat, indent=1))
                S.sy_bytes = None
                st.success(f"Reviewed syllabus saved for Semester {sem}.")
                st.rerun()
            except OSError as ex:
                st.error(f"Couldn't save the syllabus ({ex}).")

def report_panel():
    if errs:
        st.error("The workbook has unresolved validation errors. Review the Data check before exporting an analysis workbook.")
        return
    try:
        report = make_generic_report(S.digest, S.workbook_type, selected_term, ds)
        st.download_button("Download validated analysis workbook", report,
                           f"result_analysis_sem{selected_term}.xlsx", type="primary", width="stretch",
                           key=f"download_analysis_{selected_term}")
    except ValueError as ex: st.error(str(ex))
    st.caption("Includes an overview, student results, subject analysis, source rows, and methodology notes. "
               "Headers and student identifiers stay visible while scrolling.")

with st.sidebar:
    st.markdown('<div class="gs-brand"><div class="gs-mark">GS</div><div><div class="gs-brand-title">GradeSense</div>'
                '<div class="gs-brand-sub">ACADEMIC RESULTS INTELLIGENCE</div></div></div>', unsafe_allow_html=True)
    st.markdown(f'<div class="gs-card"><b>ACTIVE WORKBOOK</b><br>{S.fname}<br>{ds.wb.info["students"]:,} students · '
                f'{len(eng.terms)} semester(s)</div>', unsafe_allow_html=True)
    if len(eng.terms) > 1:
        S.default_term = st.selectbox("Default term", eng.terms, index=len(eng.terms) - 1, key="default_term_selector",
                                      help="Used when a question doesn't name a term. Say 'sem 3' in a question to override.")
    else: S.default_term = eng.terms[0]
    with st.expander(("⚠️ " if errs else "✅ ") + ("Data check: %d issue(s)" % len(errs) if errs else "Data check passed")):
        for c in chk:
            st.markdown(f"{'✅' if c.ok else '❌' if c.level == 'error' else 'ℹ️'} {c.name} — {c.failed} of {c.checked}" + (f"  \n<small>e.g. {'; '.join(map(str, c.examples))}</small>" if c.examples else ""), unsafe_allow_html=True)
    with st.expander("Syllabus catalogue", expanded=bool(S.sy_bytes)): syllabus_panel()
    with st.expander("Download result analysis", expanded=False): report_panel()
    with st.expander("Change workbook", expanded=False):
        change_type_key = f"change_workbook_type_{S.digest}"
        st.selectbox("Workbook type", list(WORKBOOK_TYPES), format_func=WORKBOOK_TYPES.get,
                     index=list(WORKBOOK_TYPES).index(S.get(change_type_key, S.workbook_type)),
                     key=change_type_key)
        up2 = st.file_uploader("Another result sheet (.xlsx)", type=["xlsx"], key="xl_up2")
        if up2: load(up2.getvalue(), up2.name, S[change_type_key])
        if S.error: st.error(S.error)
    recent = list(dict.fromkeys(m["text"] for m in reversed(S.msgs) if m["role"] == "user"))[:5]
    if recent:
        st.markdown("**Run a previous query**")
        for j, q in enumerate(recent):
            st.button(q[:38] + ("…" if len(q) > 38 else ""), key=f"replay_query_{j}", on_click=ask_cb, args=(q,),
                      width="stretch", help=f"Run this analysis again: {q}")
    st.caption("Results are derived from the active workbook. Processing remains on this computer.")

if S.sy_bytes:
    syllabus_draft_panel(S.sy_bytes)

analysis_request_query = analysis_request_panel()

# ---------------------------------------------------------------- chat
def stream(text):
    for w in re.split(r"(\s+)", text): yield w; time.sleep(0.01)

def xl_bytes(df):
    b = io.BytesIO(); df.to_excel(b, index=False); return b.getvalue()

def show(m, i):
    with st.chat_message(m["role"], avatar="🎓" if m["role"] == "assistant" else None):
        if m.pop("fresh", False): st.write_stream(stream(m["text"]))
        else: st.markdown(m["text"])
        t = m.get("table")
        if t is not None:
            t = t.copy()
            for c in t.columns:                                       # mixed text/number columns can't be drawn; show as text
                if t[c].dtype == object: t[c] = t[c].astype(str)
            st.dataframe(t, hide_index=True, width="stretch", height=min(430, 38 + 35 * len(t)))
            a, b, _ = st.columns([1, 1, 5])
            a.download_button("Download Excel table", xl_bytes(t), f"gradesense_{i}.xlsx", key=f"download_excel_{i}")
            b.download_button("Download CSV table", t.to_csv(index=False).encode(), f"gradesense_{i}.csv", key=f"download_csv_{i}")
        if m.get("chart") is not None: st.bar_chart(m["chart"])

chat_query = st.chat_input("Ask a question about these results…", key=f"analysis_chat_{S.chat_epoch}")
pending_query = S.pop("pending", None)
q = chat_query or pending_query or analysis_request_query
if q:
    S.msgs.append({"role": "user", "text": q})
    try:
        is_request = not chat_query and not pending_query and analysis_request_query is not None
        with st.spinner("Analysing…"):
            a = eng.ask_request(q, S.default_term) if is_request else eng.ask(q, S.default_term)
    except Exception:
        (DATA / "errors.log").open("a").write(f"{q!r}\n{traceback.format_exc()}\n")
        a = Answer("Sorry, that question hit an error. Try rephrasing it, or ask *help* to see what I can do.")
    S.msgs.append({"role": "assistant", "text": a.text, "table": a.table, "chart": a.chart, "fresh": True})

if not S.msgs:
    if errs: st.warning("The data check found problems in this sheet, so some numbers may be off. See *Data check* in the sidebar.")
    b = term_students
    cleared = int((b.backlogs == 0).sum())
    pass_rate = 100 * cleared / len(b) if len(b) else 0
    st.markdown(f'<div class="gs-hero"><div class="gs-eyebrow">Semester {selected_term}</div>'
                '<div class="gs-title">Results overview</div>'
                '<div class="gs-sub">A concise view of the active workbook. Ask a question below to explore student and subject outcomes.</div></div>',
                unsafe_allow_html=True)
    metric_cols = st.columns(4)
    metrics = [("Students", f"{len(b):,}"), ("Subjects", f"{term_subjects:,}"),
               ("Class average SGPA", f"{b.sgpa.mean():.2f}" if len(b) else "—"),
               ("Cleared all subjects", f"{pass_rate:.1f}%")]
    for col, (label, value) in zip(metric_cols, metrics):
        col.markdown(f'<div class="gs-metric"><div class="gs-metric-label">{label}</div>'
                     f'<div class="gs-metric-value">{value}</div></div>', unsafe_allow_html=True)
    st.markdown('<div class="gs-chat-intro">Explore this result sheet</div>'
                '<div class="gs-sub">Select a focused analysis or enter your own question in the chat box.</div>',
                unsafe_allow_html=True)
    cols = st.columns(2)
    for j, ex in enumerate(eng.examples()):
        cols[j % 2].button(ex, key=f"analysis_prompt_{j}", on_click=ask_cb, args=(ex,), width="stretch",
                            help=f"Run this analysis for semester {selected_term}: {ex}")
for i, m in enumerate(S.msgs): show(m, i)
