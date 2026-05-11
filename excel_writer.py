import os
import shutil
import tempfile
import pandas as pd

from config import OUTPUT_FILE
from field_map import EXCEL_COLUMNS
from logger import get_logger

log = get_logger()


# ── Internal helpers ─────────────────────────────────────────────

def _read_master() -> pd.DataFrame:
    """
    Read existing master file or return empty DataFrame.
    """
    if not os.path.exists(OUTPUT_FILE):
        log.info("Master file not found — starting fresh.")
        return pd.DataFrame(columns=EXCEL_COLUMNS)

    try:
        df = pd.read_excel(OUTPUT_FILE, engine="openpyxl")
    except Exception as e:
        log.error(f"Failed to read master file: {e}")
        return pd.DataFrame(columns=EXCEL_COLUMNS)

    # Ensure all expected columns exist
    for col in EXCEL_COLUMNS:
        if col not in df.columns:
            df[col] = pd.NA

    return df[EXCEL_COLUMNS]


def _write_master(df: pd.DataFrame) -> None:
    """
    Atomic write to prevent file corruption.
    """
    master_dir = os.path.dirname(os.path.abspath(OUTPUT_FILE))
    os.makedirs(master_dir, exist_ok=True)

    fd, tmp_file = tempfile.mkstemp(suffix=".xlsx", dir=master_dir)
    os.close(fd)

    try:
        df.to_excel(tmp_file, index=False, engine="openpyxl")
        shutil.move(tmp_file, OUTPUT_FILE)  # atomic replace
    except Exception as e:
        log.error(f"Failed to write master file: {e}")
        raise


# ── Public API ───────────────────────────────────────────────────

def save_to_excel(rows: list[dict]) -> int:
    """
    Append rows to master Excel file.
    """
    if not rows:
        log.warning("No rows to save.")
        return 0

    # Read existing
    existing_df = _read_master()

    # Create new DataFrame
    new_df = pd.DataFrame(rows, columns=EXCEL_COLUMNS)

    # Combine
    combined_df = pd.concat([existing_df, new_df], ignore_index=True)

    # OPTIONAL: remove duplicates if tag_number exists
    if "tag_number" in combined_df.columns:
        combined_df.drop_duplicates(subset=["tag_number"], keep="last", inplace=True)

    # Write
    _write_master(combined_df)

    total = len(combined_df)
    log.info(f"Master updated: +{len(rows)} rows, total={total}")

    return total


def get_row_count() -> int:
    """
    Get number of rows in master file.
    """
    if not os.path.exists(OUTPUT_FILE):
        return 0

    try:
        df = pd.read_excel(OUTPUT_FILE, engine="openpyxl")
        return len(df)
    except Exception as e:
        log.error(f"Failed to read row count: {e}")
        return 0