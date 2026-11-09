"""Local OCR for user-provided analysis requests."""
import io
import re

import pdfplumber
import pytesseract
from PIL import Image

from .syllabus import _find_tesseract

_find_tesseract()


MAX_FILE_BYTES = 10 * 1024 * 1024
MAX_PDF_PAGES = 5
MAX_IMAGE_PIXELS = 40_000_000


class AnalysisRequestError(ValueError):
    """A request file could not be safely or reliably read."""


def _ocr_pdf(data):
    text = []
    with pdfplumber.open(io.BytesIO(data)) as pdf:
        if len(pdf.pages) > MAX_PDF_PAGES:
            raise AnalysisRequestError(
                f"PDF has {len(pdf.pages)} pages; request files may contain at most {MAX_PDF_PAGES} pages."
            )
        for page in pdf.pages:
            page_text = page.extract_text() or ""
            if not page_text.strip():
                page_text = pytesseract.image_to_string(
                    page.to_image(resolution=180).original, config="--psm 6"
                )
            if page_text.strip():
                text.append(page_text)
    return "\n".join(text)


def extract_analysis_request(data: bytes, filename: str) -> str:
    """Extract reviewable text from a PNG, JPEG, or PDF without sending it elsewhere."""
    suffix = filename.rsplit(".", 1)[-1].lower() if "." in filename else ""
    if suffix not in {"png", "jpg", "jpeg", "pdf"}:
        raise AnalysisRequestError("Choose a PNG, JPG, JPEG, or PDF request file.")
    if not data:
        raise AnalysisRequestError("The selected request file is empty.")
    if len(data) > MAX_FILE_BYTES:
        raise AnalysisRequestError("Request files must be 10 MB or smaller.")

    try:
        if suffix == "pdf":
            if not data.startswith(b"%PDF"):
                raise AnalysisRequestError("The selected file is not a valid PDF.")
            text = _ocr_pdf(data)
        else:
            with Image.open(io.BytesIO(data)) as image:
                if image.format not in {"PNG", "JPEG"}:
                    raise AnalysisRequestError("The file contents do not match a PNG or JPEG image.")
                if image.width * image.height > MAX_IMAGE_PIXELS:
                    raise AnalysisRequestError("Images must be 40 megapixels or smaller.")
                text = pytesseract.image_to_string(image.convert("L"), config="--psm 6")
    except pytesseract.TesseractNotFoundError as ex:
        raise AnalysisRequestError(
            "Local Tesseract OCR is not installed or configured. Install Tesseract to read request images."
        ) from ex
    except (OSError, Image.DecompressionBombError) as ex:
        raise AnalysisRequestError("Could not read this file as a valid image or PDF.") from ex

    cleaned = re.sub(r"[ \t]+", " ", text or "")
    cleaned = re.sub(r"\n{3,}", "\n\n", cleaned).strip()
    if not cleaned:
        raise AnalysisRequestError("No readable request text was found. Try a clearer image or PDF.")
    if len(cleaned) > 20_000:
        raise AnalysisRequestError("The extracted request is too long. Upload a shorter, focused request.")
    return cleaned
