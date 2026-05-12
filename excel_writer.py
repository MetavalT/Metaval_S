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


def save_to_excel(rows: list[dict]):

    df_new = pd.DataFrame(rows)

    # Ensure TagNo exists
    if "TagNo" not in df_new.columns:
        return

    df_new["TagNo"] = df_new["TagNo"].astype(str).str.strip().str.upper()

    # ── CASE 1: File exists → UPDATE ─────────────────────
    if os.path.exists(OUTPUT_FILE):

        df_master = pd.read_excel(OUTPUT_FILE, engine="openpyxl")

        df_master["TagNo"] = df_master["TagNo"].astype(str).str.strip().str.upper()

        df_master.set_index("TagNo", inplace=True)
        df_new.set_index("TagNo", inplace=True)
        # 🚀 REMOVE DUPLICATE TAGS (KEEP LAST)
        df_new = df_new[~df_new.index.duplicated(keep="last")]
        # Update existing rows
        df_master.update(df_new)

        # Add new rows
        new_tags = df_new.index.difference(df_master.index)
        df_master = pd.concat([df_master, df_new.loc[new_tags]])

        df_master.reset_index(inplace=True)

    # ── CASE 2: File does NOT exist → CREATE ─────────────
    else:
        df_master = df_new.copy()

    # ── SAVE BACK TO SAME FILE ──────────────────────────
    df_master.to_excel(OUTPUT_FILE, index=False)
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
    
