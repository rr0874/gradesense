"""Streamlit UI smoke tests. Run from the project folder: python -m tests.test_ui"""
from pathlib import Path
import hashlib
import io
import warnings
from unittest.mock import patch

warnings.filterwarnings("ignore")

from openpyxl import Workbook
from streamlit.testing.v1 import AppTest

ROOT = Path(__file__).resolve().parents[1]
app = AppTest.from_file(str(ROOT / "app.py"), default_timeout=180).run()
assert not app.exception, [e.message for e in app.exception]
assert any("Student results, clearly understood." in item.value for item in app.markdown)
assert any(button.label == "Analyze workbook" for button in app.button)
assert any(button.label == "Open demonstration workbook · AIML, 153 students" for button in app.button)
assert app.radio[0].options == ["Whole-semester result", "MSE result"]
assert app.radio[0].value == "semester"

next(button for button in app.button if button.label == "Analyze workbook").click().run()
assert not app.exception, [e.message for e in app.exception]
assert any("Choose an .xlsx workbook" in item.value for item in app.error)
next(button for button in app.button if button.label == "Open demonstration workbook · AIML, 153 students").click().run()
assert not app.exception, [e.message for e in app.exception]
assert any("Results overview" in item.value for item in app.markdown)
assert any(item.label == "Analyze a request from an image or PDF" for item in app.expander)
assert [button.label for button in app.button].count("New conversation") == 1

# Changing the term must update the visible context before showing its metrics.
app.selectbox[0].select("III").run()
assert app.session_state["default_term"] == "III"
assert any("Semester III" in item.value for item in app.markdown), "The header and overview should reflect the selected term."

prompt_labels = [button.label for button in app.button
                 if button.label not in ("Home", "New conversation", "Use the sample syllabus (Sem IV)")]
assert prompt_labels and len(prompt_labels) == len(set(prompt_labels)), "Analysis actions should be clear and distinct."
visible_actions = [button.label for button in app.button]
assert len(visible_actions) == len(set(visible_actions)), "Visible buttons should not duplicate another action."

for prompt in prompt_labels:
    next(button for button in app.button if button.label == prompt).click().run()
    assert not app.exception, [f"{prompt}: {e.message}" for e in app.exception]
    assert len(app.chat_message) >= 2, f"{prompt} should produce an answer."
    next(button for button in app.button if button.label == "New conversation").click().run()

next(button for button in app.button if button.label == "Class summary").click().run()
assert not app.exception, [e.message for e in app.exception]
assert any("Term III" in item.value for item in app.markdown), "The analysis should use the selected term."
downloads = [button.label for button in app.get("download_button")]
assert downloads == ["Download Excel table", "Download CSV table", "Download validated analysis workbook"], downloads
next(button for button in app.get("download_button")
     if button.label == "Download validated analysis workbook").click().run()
assert not app.exception, [e.message for e in app.exception]

old_chat_id = app.chat_input[0].proto.id
next(button for button in app.button if button.label == "New conversation").click().run()
assert not app.exception, [e.message for e in app.exception]
assert len(app.chat_message) == 0, "Starting a conversation should clear the conversation only."
assert any("Results overview" in item.value for item in app.markdown), "The loaded workbook should remain active."
assert app.session_state["default_term"] == "III", "Starting a conversation should keep the selected semester."
assert app.chat_input[0].proto.id != old_chat_id, "Starting a conversation should reset the chat input."

app.session_state["sy_bytes"] = (ROOT / "data" / "scheme_sem4.png").read_bytes()
app.run()
assert not app.exception, [e.message for e in app.exception]
assert any("Review syllabus draft" in item.value for item in app.markdown), "The editable syllabus draft should be visible in the main workspace."
assert len(app.dataframe) == 1, "Users should be able to correct OCR values in the visible draft."
assert any(item.label == "Save reviewed syllabus" for item in app.button), "The visible draft should have an explicit save action."

# An image request is reviewed before it is passed to the existing result engine.
request_app = AppTest.from_file(str(ROOT / "app.py"), default_timeout=180).run()
next(button for button in request_app.button
     if button.label == "Open demonstration workbook · AIML, 153 students").click().run()
next(uploader for uploader in request_app.file_uploader
     if uploader.label == "Analysis request (.png, .jpg, .jpeg, or .pdf)").set_value(
    ("analysis.png", (ROOT / "data" / "scheme_sem4.png").read_bytes(), "image/png")
).run()
with patch("backend.analysis_request.pytesseract.image_to_string", return_value="class summary"):
    next(button for button in request_app.button if button.label == "Read request text").click().run()
assert any(widget.label == "Recognized text — review or correct before analysis"
           for widget in request_app.text_area)
next(button for button in request_app.button if button.label == "Run analysis request").click().run()
assert not request_app.exception, [e.message for e in request_app.exception]
assert any(message.value.startswith("**Term IV — 153 students**")
           for message in request_app.markdown)
assert any(button.label == "Download Excel table" for button in request_app.get("download_button"))

# MSE mode gets its own explicit analysis action and renders both output tables.
mse_app = AppTest.from_file(str(ROOT / "app.py"), default_timeout=180).run()
mse_book = Workbook()
mse_sheet = mse_book.active
mse_sheet.append(["Student Marks Register"])
mse_sheet.append([None, None, None, None, None, "DS-C", "IPR"])
mse_sheet.append([None, None, None, None, None, "Int Passing (20)", "Int Passing (20)"])
mse_sheet.append([None, None, None, None, None, "Ext Passing (20)", "Ext Passing (0)"])
mse_sheet.append(["Sr No", "Student Unique No", "Student Seat No", "Student Name",
                  "Student Mother Name", "Marks (30)", "Marks (30)"])
mse_sheet.append([1, "UID001", "T001", "Student One", "Parent One", 11, "Abs"])
mse_sheet.append([2, "UID002", "T002", "Student Two", "Parent Two", 12, 10])
mse_output = io.BytesIO()
mse_book.save(mse_output)
mse_data = mse_output.getvalue()
mse_app.session_state["data"] = mse_data
mse_app.session_state["digest"] = hashlib.sha1(mse_data).hexdigest()[:12]
mse_app.session_state["fname"] = "Student_Marks_MSE_test_data.xlsx"
mse_app.session_state["workbook_type"] = "mse"
mse_app.session_state["change_workbook_type"] = "mse"
mse_app.run()
assert not mse_app.exception, [e.message for e in mse_app.exception]
assert any(item.label == "Analyze a request from an image or PDF" for item in mse_app.expander)
next(button for button in mse_app.button if button.label == "MSE RESULT analysis").click().run()
assert not mse_app.exception, [e.message for e in mse_app.exception]
assert len(mse_app.dataframe) == 2, "MSE analysis should show subject and student outcome tables."
assert mse_app.dataframe[0].value.iloc[0]["Failed"] == 1
assert mse_app.dataframe[0].value.iloc[1]["Absent"] == 1
assert any(button.label == "Download validated analysis workbook" for button in mse_app.get("download_button"))
next(button for button in mse_app.get("download_button")
     if button.label == "Download validated analysis workbook").click().run()
assert not mse_app.exception, [e.message for e in mse_app.exception]
next(button for button in mse_app.button if button.label == "Who failed?").click().run()
assert not mse_app.exception, [e.message for e in mse_app.exception]
assert len(mse_app.chat_message) == 2
assert any("Found 2 students with overall MSE result Fail" in item.value
           for item in mse_app.markdown)
next(button for button in mse_app.button if button.label == "Home").click().run()
assert not mse_app.exception, [e.message for e in mse_app.exception]
assert any("Active workbook is still loaded" in item.value for item in mse_app.info)
assert any(button.label == "Return to active workbook" for button in mse_app.button)
next(button for button in mse_app.button if button.label == "Return to active workbook").click().run()
assert not mse_app.exception, [e.message for e in mse_app.exception]
assert mse_app.session_state["fname"] == "Student_Marks_MSE_test_data.xlsx"

# The MSE Analysis II export is separately available for workbooks containing
# one class/division sheet each.
division_book = Workbook()
for sheet, title, subjects, entries in [
    (division_book.active, "Second Year Division - A Marksheet", ["FAI"], [
        ["A01", "Student One", 11],
        ["A02", "Student Two", 12],
    ]),
]:
    sheet.title = "SY- A"
    sheet.append(["College"])
    sheet.append(["Department"])
    sheet.append([title])
    sheet.append([None, "Roll No", "Name", "Marks", "Total"])
    sheet.append([None, None, None, *subjects])
    for roll, name, mark in entries:
        sheet.append([None, roll, name, mark, None])
division_bytes = io.BytesIO()
division_book.save(division_bytes)
division_data = division_bytes.getvalue()
mse_app.session_state["data"] = division_data
mse_app.session_state["digest"] = hashlib.sha1(division_data).hexdigest()[:12]
mse_app.session_state["fname"] = "MSE MARKSHEET 2026-27 (SEM-I).xlsx"
mse_app.session_state["mse_kind"] = "divisions"
mse_app.session_state["msgs"] = []
mse_app.session_state["chat_epoch"] += 1
mse_app.run()
assert not mse_app.exception, [e.message for e in mse_app.exception]
assert any("MSE Performance Analysis" in item.value for item in mse_app.markdown)
assert any(button.label == "Download MSE Analysis II workbook"
           for button in mse_app.get("download_button"))
next(button for button in mse_app.get("download_button")
     if button.label == "Download MSE Analysis II workbook").click().run()
assert not mse_app.exception, [e.message for e in mse_app.exception]
assert any("Chat about these results" in item.value for item in mse_app.markdown)
next(button for button in mse_app.button if button.label == "Who failed in FAI?").click().run()
assert not mse_app.exception, [e.message for e in mse_app.exception]
assert len(mse_app.chat_message) == 2
assert any("Found 1 fail MSE result" in item.value for item in mse_app.markdown)
next(button for button in mse_app.button if button.label == "Home").click().run()
assert not mse_app.exception, [e.message for e in mse_app.exception]
assert any("Active workbook is still loaded" in item.value for item in mse_app.info)
next(button for button in mse_app.button if button.label == "Return to active workbook").click().run()
assert not mse_app.exception, [e.message for e in mse_app.exception]
assert any("MSE Performance Analysis" in item.value for item in mse_app.markdown)
assert mse_app.session_state["fname"] == "MSE MARKSHEET 2026-27 (SEM-I).xlsx"

print("UI tests: landing, whole-semester flow, MSE analysis, exports, and conversation reset passed")
