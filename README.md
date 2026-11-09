# GradeSense

Chat with a university result sheet. Upload the ERP result `.xlsx` (and optionally the semester syllabus image/PDF),
then ask in plain English: *who failed in OS?*, *failed CA but passed OS*, *top 5 students*, *compare X and Y*.

**Fully offline.** No API, no cloud AI. Answers come only from your sheet.

## Folder structure

```
gradesense/
├─ app.py                     chat UI (Streamlit)
├─ requirements.txt
├─ README.md
├─ .streamlit/
│  └─ config.toml             theme + settings
├─ data/
│  ├─ AIML_4.xlsx             sample result sheet (Terms I-IV, 153 students)
│  ├─ Sample_Student_Results_Gen_AI.xlsx   second sample layout (Term II, 50 students)
│  ├─ scheme_sem4.png         sample syllabus image (Semester IV)
│  └─ unmatched_queries.csv   created automatically: questions the bot didn't understand
├─ backend/
│  ├─ ingest.py               reads any term/subject layout of the ERP sheet -> clean tables
│  ├─ validate.py             re-checks the sheet's own arithmetic (SGPA, totals, grades)
│  ├─ syllabus.py             offline OCR: syllabus image/PDF -> course catalog
│  ├─ catalog.py              attaches full subject names / marking schemes
│  ├─ store.py                SQLite store for one workbook
│  ├─ engine.py               question understanding + analysis
│  ├─ mse.py                  MSE threshold and attendance analysis
│  ├─ __main__.py             command-line check:  python -m backend data/AIML_4.xlsx
│  └─ catalogs/
│     └─ semIV.json           saved syllabus (one file per semester)
└─ tests/
   ├─ test_backend.py         ingestion, validation, syllabus reading
   ├─ test_engine.py          63 checks on answers (both sample layouts)
   └─ test_mse.py             MSE parsing, thresholds, absences, and semester regression
```

## Setup (VS Code, Windows)

1. Install **Python 3.10+** (tick "Add Python to PATH") and **Tesseract OCR** (UB Mannheim Windows build).
   Tesseract is only needed to read syllabus images; the chatbot itself works without it.
2. VS Code: *File > Open Folder* > `gradesense`. Install the **Python** extension.
3. Terminal > New Terminal, then:
   ```
   python -m venv venv
   venv\Scripts\activate
   pip install -r requirements.txt
   ```
   (PowerShell blocks activation? Run `Set-ExecutionPolicy -Scope Process Bypass` first.)
   Mac/Linux: `source venv/bin/activate`.
4. Run:
   ```
   streamlit run app.py
   ```
   It opens at http://localhost:8501. Stop with Ctrl+C.

## Using it

1. Choose **Whole-semester result** or **MSE result**, upload the workbook, then press **Analyze workbook** (or press *Try the sample sheet* for the existing whole-semester demo).
   Whole-semester uploads retain the existing chat, checks, and validated analysis workbook.
   MSE mode supports ERP subject-block workbooks with explicit per-subject `MSE` columns, normalized long tables with student ID/name, subject, and an explicit MSE marks column, and wide MSE registers with a labeled student ID/name header, one `Marks (30)` column per subject, and each subject named above its marks column. It validates marks against the 0–30 range,
   classifies marks below 12 as Fail and 12 or above as Pass, and reports `AB`/`Absent` separately; blanks and dashes are unreported.
   For workbooks with one class/division marksheet per worksheet (Roll No, Name, a Marks heading, and subject names above the marks columns),
   the sidebar also offers **Download MSE Analysis II workbook**. It preserves the uploaded division marksheets and adds an `ISE MSE Analysis`
   summary and `Checks` sheet, matching the supplied college workbook’s merged table layout and formatting. The summary uses live Excel formulas:
   numeric marks are appeared, marks `>=12` pass, marks `<12` fail, and `AB`, blank, or dash marks count as absent. Numeric-only formulas ensure blank
   cells are not counted as failures. Courses with no marks entered are shown as `-`; the checks sheet reconciles appeared + absent against
   the named student count and passed + failed against appeared.
   MSE workspaces use the same connected results-overview layout, including summary cards, suggested analyses, and an always-available chatbot for
   workbook-based summary, subject, student, pass, fail, and absence questions. The **Home** button returns to the upload page without discarding
   the active workbook; **Return to active workbook** resumes the same analysis session. Use **MSE RESULT analysis** for per-subject and overall
   student outcomes. Unsupported or ambiguous MSE layouts are rejected with an explanation
   rather than being guessed. Per-subject pass rates exclude absent/unreported results; the report includes both mutually exclusive overall
   statuses and a separate count of students absent in any subject. Overall status is Fail if any subject fails, otherwise Absent if any subject
   is absent, Unreported if a result is missing, and Pass only if all MSE subjects pass. The MSE download creates the same four report tabs
   as the semester export: `Result_Analysis`, `subjectwise`, `Data`, and `Notes`.
2. *(Optional)* Sidebar > **Syllabus catalogue**: upload the teaching-scheme image/PDF. It is read offline.
   Edit any uncertain OCR values in the visible **Review syllabus draft** table, type the semester (e.g. `IV`), and press **Save reviewed syllabus**.
   It is saved in `backend/catalogs/` and reused for every later sheet of that semester.
3. **Ask.** Name a term to override the default: *summary of sem 3*, *who failed in CA in sem 2*.
4. Every table has **Excel / CSV** download buttons.

Sidebar > **Data check** shows whether the sheet's own totals reconcile. A red item means some numbers in the sheet look wrong.

The validated **analysis workbook** has a filterable `Result_Analysis` student-by-subject table, `subjectwise` outcome summary, and auditable `Data` and `Notes` tabs. Student identifiers and headers stay frozen while scrolling; report columns have explicit widths, and blank source results remain blank rather than being filled with guessed values. Ranking uses SGPA by default; the workbook does not invent institution-specific class bands or unknown maximum marks. Export is blocked when source validation reports errors.

## What you can ask

| Kind | Examples |
|---|---|
| Overall | *class summary*, *summary of sem 2*, *ATKT students*, *students with no backlogs* |
| Pass / fail | *who failed*, *who failed in OS*, *absent in CAL*, *pass percentage of CA* |
| Combined | *failed CA but passed OS*, *failed CA or OS*, *scored above 80 in DCN and below 40 in OS* |
| Grades | *students with O grade in WT*, *grade distribution in ES* |
| Rankings | *top 5 students*, *toppers in DCN*, *hardest subject* |
| Filters | *students with SGPA above 9*, *students with more than 1 backlog* |
| Student | *analysis of Sukhada Done*, *result of EN23209642*, *did Sukhada pass OS?*, *sgpa trend of ...* |
| Compare | *compare Sukhada and Prathmesh*, *compare CA and OS* |
| Progress | *who improved from sem 3 to sem 4*, *who dropped* |
| Scheme | *list of subjects*, *credits of Operating System* |

## Tests

```
python -m tests.test_backend data/scheme_sem4.png
python -m tests.test_engine
python -m tests.test_excel_compat
python -m tests.test_ui
python -m tests.test_mse
```

## Command line (no UI)

```
python -m backend data/AIML_4.xlsx data/scheme_sem4.png
```
Prints what was extracted, any items to review, and the data check.

## Known limits

- Reads the ERP **multi-level `.xlsx`** layouts exercised by the regression tests: the multi-term AIML sheet; a single-term sheet whose
  header prints *Max Marks, Credits, Pass rule* and has INT/POE or single-mark blocks; and a sheet where the Term/subject rows appear
  above the student header and subject names include parenthetical descriptors. Blocks of any width are fine. It also accepts normalized
  long tables with one student-course result per row and explicit student ID/name, subject, marks, and grade or result/status columns.
  For workbooks with staff analysis tabs, the first worksheet must be the ERP input sheet; later analysis tabs are not ingested.
  Unstructured sheets, wide tables with one column per course, and workbooks without explicit grade/status rules are not guessed.
- Chat rankings and summaries use **SGPA by default** and use **CGPA when the question explicitly asks for it** and the source provides it.
  Reports are generated from the parsed workbook, not a required semester-specific catalog. Unknown maxima and institution-specific class
  bands are left blank/not invented. Export is blocked when source validation reports errors.
- Syllabus OCR uses local image enhancement retries for low contrast or uneven lighting. It still requires local Tesseract OCR; uncertain
  readings are flagged for human review and must not be saved without checking. Result-sheet analysis does not depend on a syllabus.
- When the sheet prints max marks and credits (the single-term variant), no syllabus is needed for marking schemes or percentage questions; the syllabus
  only adds full subject names.
- `data/` contains optional demonstration and regression workbooks, not model-training data. No network API or remote model is used.
- Syllabus reading expects the 13-column teaching-scheme table (Code, Type, Name, Credits, L, P, T, ISE, MSE, ESE, INT, OE/POE, Total)
  in a straight, sharp image. Rows that don't add up to the printed totals are flagged for review. PDFs are read as images.
- It only understands the question types in the table above. Anything else returns the help list and is logged to
  `data/unmatched_queries.csv`; add phrasings to `TEMPLATES` in `backend/engine.py` to teach it.
- "Failed" counts absent students too (the answer says so when it applies).
- Overall grade and CGPA are only printed for the latest term in the sheet.

## Troubleshooting

- `streamlit` not recognised: the venv isn't active, or `pip install` didn't finish.
- "OCR unavailable" / syllabus won't read: install Tesseract, or set the `TESSERACT_CMD` environment variable to its `tesseract.exe`.
- "can't be read yet": the workbook layout isn't supported (see limits).
- Wrong Python in VS Code: Ctrl+Shift+P > *Python: Select Interpreter* > pick the `venv` one.
