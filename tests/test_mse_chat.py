"""Focused checks for MSE chatbot answers over validated result data."""
import io

from openpyxl import Workbook

from backend.mse_chat import answer_division_mse
from backend.mse_divisions import read_division_mse_workbook


book = Workbook()
sheet = book.active
sheet.title = "SY- A"
sheet.append(["College"])
sheet.append(["Department"])
sheet.append(["Second Year Division - A Marksheet"])
sheet.append([None, "Roll No", "Name", "Marks", None, "Total"])
sheet.append([None, None, None, "FAI", "IPRM"])
sheet.append([None, "A01", "Student One", 11, "AB"])
sheet.append([None, "A02", "Student Two", 12, None])
payload = io.BytesIO()
book.save(payload)
analysis = read_division_mse_workbook(payload.getvalue(), "MSE MARKSHEET 2026-27.xlsx")

failed = answer_division_mse(analysis, "Who failed in FAI?")
assert len(failed.table) == 1
assert failed.table.iloc[0]["Roll No"] == "A01"
assert failed.table.iloc[0]["MSE Marks"] == 11

absent = answer_division_mse(analysis, "Who was absent in division A?")
assert len(absent.table) == 2
assert set(absent.table["Outcome"]) == {"Absent"}

summary = answer_division_mse(analysis, "Overall MSE summary")
assert set(summary.table["Course"]) == {"FAI", "IPRM"}
assert int(summary.table["Fail"].sum()) == 1
assert int(summary.table["Absent"].sum()) == 2

student = answer_division_mse(analysis, "Show results for student A01")
assert set(student.table["Course"]) == {"FAI", "IPRM"}
assert len(answer_division_mse(analysis, "failed in unknown subject").table) == 1
print("MSE chat tests: summary, subject, student, and outcome queries passed")
