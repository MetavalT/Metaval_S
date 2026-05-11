import pdfplumber
import pytesseract
from pdf2image import convert_from_path
from logger import get_logger

log = get_logger()

# ── CONFIGURE PATHS ─────────────────────────────────────────────

# Tesseract executable path (adjust if different on your system)
pytesseract.pytesseract.tesseract_cmd = (
    r"C:\Users\Stakshi\AppData\Local\Programs\Tesseract-OCR\tesseract.exe"
)
# Poppler bin path
POPPLER_PATH = r"C:\poppler-26.02.0\Library\bin"

log.info("Processing started")
log.error("Something failed")

# ── MAIN FUNCTION ──────────────────────────────────────────────

def extract_text(pdf_path: str) -> str:
    """
    Extract text from PDF.
    1. Try pdfplumber (fast, for normal PDFs)
    2. If no text → fallback to OCR (for scanned PDFs)
    """

    full_text = ""

    # ── Step 1: Try normal extraction ──────────────────────────
    try:
        with pdfplumber.open(pdf_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"

        if full_text.strip():
            log.info(f"Text extracted using pdfplumber: {pdf_path}")
            return full_text

    except Exception as e:
        log.warning(f"pdfplumber failed: {e}")

    # ── Step 2: OCR fallback ───────────────────────────────────
    try:
        log.info("No selectable text found. Running OCR...")

        images = convert_from_path(
            pdf_path,
            poppler_path=POPPLER_PATH
        )

        for i, image in enumerate(images):
            text = pytesseract.image_to_string(image)
            full_text += text + "\n"
            log.info(f"OCR processed page {i + 1}")

        if full_text.strip():
            log.info(f"Text extracted using OCR: {pdf_path}")
            return full_text
        else:
            log.error(f"OCR failed to extract any text: {pdf_path}")
            return ""

    except Exception as e:
        log.error(f"OCR extraction failed: {e}")
        return ""