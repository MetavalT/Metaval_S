import logging
import os
from config import LOG_FILE, LOG_LEVEL

def get_logger():
    logger = logging.getLogger("pdf_pipeline")

    if logger.handlers:
        return logger  # Prevent duplicate logs

    logger.setLevel(LOG_LEVEL)

    # Ensure log folder exists
    os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

    # File handler
    file_handler = logging.FileHandler(LOG_FILE)
    file_handler.setLevel(LOG_LEVEL)

    # Console handler (optional but useful)
    console_handler = logging.StreamHandler()
    console_handler.setLevel(LOG_LEVEL)

    # Format
    formatter = logging.Formatter(
        "%(asctime)s | %(levelname)s | %(message)s"
    )

    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)

    logger.addHandler(file_handler)
    logger.addHandler(console_handler)

    return logger