import os
import shutil
import tempfile
import threading

from flask import (
    Flask, request, jsonify,
    render_template, send_file, redirect, url_for,
)
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge

from config import INPUT_FOLDER, OUTPUT_FILE, EMAIL_USER
from processor import process_record, split_records
from extractor import extract_text
from email_poller import fetch_pdf_attachments
from logger import get_logger
from database import init_db, upsert_row, export_to_excel, Session, Instrument

# ── App setup ─────────────────────────────────────────────
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024
log = get_logger()

# Thread lock for Excel safety
_excel_lock = threading.Lock()

ALLOWED_EXTENSIONS = {"pdf"}
MASTER_FILE = OUTPUT_FILE


# ── Helpers ───────────────────────────────────────────────
def allowed_file(filename: str) -> bool:
    return "." in filename and filename.lower().endswith(".pdf")

def run_pipeline(pdf_path: str) -> dict:
    log.info(f"Pipeline started: {os.path.basename(pdf_path)}")

    # Step 1: Extract text
    text = extract_text(pdf_path)
    if not text.strip():
        raise ValueError("No text extracted from PDF")

    # Step 2: Split records
    records = split_records(text)

    # Step 3: Process records
    rows = []
    skipped = 0

    for rec in records:
        row = process_record(rec)

        tag = row.get("tag_number")

        # 🚫 Skip invalid / weak tags
        if not tag:
            skipped += 1
            continue

        rows.append(row)

    # Step 4: Save to Excel
    if rows:
        with _excel_lock:
           for row in rows:
            upsert_row(row)
    else:
        log.warning("No valid records found")

    return {
        "records_found": len(rows),
        "records_skipped": skipped,
        "file": os.path.basename(pdf_path),
    }


def cleanup_upload(path):
    if os.path.exists(path):
        os.remove(path)


# ── Error Handlers ────────────────────────────────────────
@app.errorhandler(RequestEntityTooLarge)
def too_large(e):
    return jsonify({"error": "File too large"}), 413


@app.errorhandler(500)
def server_error(e):
    return jsonify({"error": "Internal server error"}), 500


# ── Routes ────────────────────────────────────────────────
@app.route("/")
def index():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]

    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Only PDF allowed"}), 400

    filename = secure_filename(file.filename)
    path = os.path.join(INPUT_FOLDER, filename)

    file.save(path)
    log.info(f"Uploaded: {filename}")

    try:
        result = run_pipeline(path)
    except Exception as e:
        log.error(f"Pipeline failed: {e}", exc_info=True)
        return jsonify({"error": str(e)}), 500
    finally:
        cleanup_upload(path)

    return jsonify({"status": "success", **result})


@app.route("/download")
def download():
    if not os.path.exists(MASTER_FILE):
        return jsonify({"error": "No master file yet"}), 404

    tmp_fd, tmp_path = tempfile.mkstemp(suffix=".xlsx")
    os.close(tmp_fd)

    with _excel_lock:
       export_to_excel(tmp_path)
    return send_file(
        tmp_path,
        as_attachment=True,
        download_name="master.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )


@app.route("/status")
def status():
    session = Session()

    try:
        count = session.query(Instrument).count()

        return jsonify({
            "rows": count,
            "exists": count > 0
        })

    except Exception as e:
        log.error(f"Status check failed: {e}")
        return jsonify({
            "rows": 0,
            "exists": False,
            "error": "Database error"
        })

    finally:
        session.close()


# ── Entry Point ───────────────────────────────────────────
if __name__ == "__main__":
    os.makedirs(INPUT_FOLDER, exist_ok=True)

    init_db()

    app.run(
        debug=False,
        host="0.0.0.0",
        port=5000
    )

    