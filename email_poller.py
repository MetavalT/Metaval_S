import imaplib
import email
import os
from config import EMAIL_HOST, EMAIL_USER, EMAIL_PASS, INPUT_FOLDER
from logger import get_logger

log = get_logger()

def fetch_pdf_attachments() -> list[str]:
    """Connect to email, download PDF attachments, return their paths."""
    saved_paths = []

    try:
        mail = imaplib.IMAP4_SSL(EMAIL_HOST)
        mail.login(EMAIL_USER, EMAIL_PASSWORD)
        mail.select(EMAIL_FOLDER)

        # Search for unread emails
        _, msg_ids = mail.search(None, "UNSEEN")
        for msg_id in msg_ids[0].split():
            _, msg_data = mail.fetch(msg_id, "(RFC822)")
            raw = msg_data[0][1]
            msg = email.message_from_bytes(raw)

            for part in msg.walk():
                if part.get_content_type() == "application/pdf":
                    filename = part.get_filename()
                    if filename:
                        save_path = os.path.join(INPUT_FOLDER, filename)
                        with open(save_path, "wb") as f:
                            f.write(part.get_payload(decode=True))
                        saved_paths.append(save_path)
                        log.info(f"Downloaded PDF from email: {filename}")

        mail.logout()

    except Exception as e:
        log.error(f"Email fetch error: {e}")

    return saved_paths