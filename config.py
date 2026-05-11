import os
from dotenv import load_dotenv

load_dotenv()

# ── Folders ─────────────────────────────────────────────
INPUT_FOLDER = os.getenv("INPUT_FOLDER", "uploads")
OUTPUT_FOLDER = os.getenv("OUTPUT_FOLDER", "output")
LOG_FOLDER = os.getenv("LOG_FOLDER", "logs")

os.makedirs(INPUT_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(LOG_FOLDER, exist_ok=True)

# ── Excel File (XLSX ONLY) ──────────────────────────────
OUTPUT_FILE = os.path.join(OUTPUT_FOLDER, "master.xlsx")

# ── Logging ─────────────────────────────────────────────
LOG_FILE = os.path.join(LOG_FOLDER, "process.log")
LOG_LEVEL = os.getenv("LOG_LEVEL", "INFO")

# ── Email Config ────────────────────────────────────────
EMAIL_USER = os.getenv("EMAIL_USER", "")
EMAIL_PASS = os.getenv("EMAIL_PASS", "")
EMAIL_HOST = os.getenv("EMAIL_HOST", "imap.gmail.com")
EMAIL_PORT = int(os.getenv("EMAIL_PORT", "993"))