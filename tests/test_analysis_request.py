"""Tests for local analysis-request extraction and conservative engine routing."""
import io
import unittest
from types import SimpleNamespace
from unittest.mock import patch

from PIL import Image

from backend.analysis_request import AnalysisRequestError, extract_analysis_request


class AnalysisRequestTests(unittest.TestCase):
    def test_image_ocr_is_normalized_and_local(self):
        image = Image.new("RGB", (20, 20), "white")
        output = io.BytesIO()
        image.save(output, format="PNG")
        with patch("backend.analysis_request.pytesseract.image_to_string", return_value="  class   summary\n\n"):
            self.assertEqual(extract_analysis_request(output.getvalue(), "request.png"), "class summary")

    def test_pdf_uses_text_layer_and_ocr_when_needed(self):
        pdf = SimpleNamespace(
            pages=[
                SimpleNamespace(extract_text=lambda: "Who failed in OS?", to_image=None),
                SimpleNamespace(
                    extract_text=lambda: "",
                    to_image=lambda resolution: SimpleNamespace(original="rendered-page"),
                ),
            ]
        )

        class PdfContext:
            pages = pdf.pages

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        with patch("backend.analysis_request.pdfplumber.open", return_value=PdfContext()), patch(
            "backend.analysis_request.pytesseract.image_to_string", return_value="Top 5 students"
        ) as ocr:
            result = extract_analysis_request(b"%PDF-test", "request.pdf")
        self.assertEqual(result, "Who failed in OS?\nTop 5 students")
        ocr.assert_called_once_with("rendered-page", config="--psm 6")

    def test_rejects_unsupported_or_oversized_uploads(self):
        with self.assertRaises(AnalysisRequestError):
            extract_analysis_request(b"data", "request.docx")
        with self.assertRaises(AnalysisRequestError):
            extract_analysis_request(b"x" * (10 * 1024 * 1024 + 1), "request.png")

    def test_rejects_pdf_with_too_many_pages(self):
        class PdfContext:
            pages = [object()] * 6

            def __enter__(self):
                return self

            def __exit__(self, *_args):
                return False

        with patch("backend.analysis_request.pdfplumber.open", return_value=PdfContext()):
            with self.assertRaisesRegex(AnalysisRequestError, "at most 5 pages"):
                extract_analysis_request(b"%PDF-test", "request.pdf")

    def test_empty_ocr_is_not_treated_as_a_successful_request(self):
        image = Image.new("RGB", (20, 20), "white")
        output = io.BytesIO()
        image.save(output, format="PNG")
        with patch("backend.analysis_request.pytesseract.image_to_string", return_value="  "):
            with self.assertRaisesRegex(AnalysisRequestError, "No readable request text"):
                extract_analysis_request(output.getvalue(), "request.png")


if __name__ == "__main__":
    unittest.main()
