"""Conservative routing tests for uploaded requests in MSE workspaces."""
import unittest
from types import SimpleNamespace

import pandas as pd

from backend.engine import Answer
from backend.mse_chat import answer_mse_request


class MseRequestTests(unittest.TestCase):
    def setUp(self):
        self.analysis = SimpleNamespace(
            records=pd.DataFrame({
                "Course": ["FAI"],
                "Student Name": ["Student One"],
                "Roll No": ["A01"],
            }),
            totals={"students": 1},
        )
        self.called = []

    def answer(self, analysis, question):
        self.called.append(question)
        return Answer("routed", pd.DataFrame([{"MSE": 12}]))

    def test_supported_summary_routes_to_existing_analysis(self):
        answer = answer_mse_request(self.analysis, "overall MSE summary", self.answer)
        self.assertEqual(answer.text, "routed")
        self.assertEqual(self.called, ["overall MSE summary"])

    def test_supported_student_or_subject_lookup_routes(self):
        answer_mse_request(self.analysis, "show results for student A01", self.answer)
        answer_mse_request(self.analysis, "FAI results", self.answer)
        self.assertEqual(len(self.called), 2)

    def test_unclear_and_unsupported_requests_do_not_route(self):
        unclear = answer_mse_request(self.analysis, "make a special analysis", self.answer)
        unsupported = answer_mse_request(self.analysis, "predict next semester results", self.answer)
        self.assertIn("couldn't confidently match", unclear.text)
        self.assertIn("can't safely produce", unsupported.text)
        self.assertEqual(self.called, [])


if __name__ == "__main__":
    unittest.main()
