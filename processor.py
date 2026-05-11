"""
processor.py
============
Splits raw PDF text into individual instrument records, extracts
key-value pairs from each record, and maps them to canonical Excel
column names using field_map.py.

Design notes
------------
* PDFs may contain multiple records separated by tag numbers
  (e.g., "303-FE-0701", "TAG-001").
* Each record is a block of text with label: value pairs on separate lines.
* Labels may vary in spelling, case, and special characters across PDFs
  — handled by field_map.resolve_column().
* Values are cleaned and stripped of surrounding whitespace/colons.
"""

import re
from typing import Optional
from logger import get_logger
from field_map import map_pdf_record, build_excel_row, resolve_column

log = get_logger()

# ── Tag number pattern — adjust regex to match your instrument tag format ──
# Examples: 303-FE-0701, TAG-001, FT-101A, LT-200
# Matches:
# Tag Number : 073-FE-1301
# Tag No.    303-FT-0701
# TAG NO     LT-200A

TAG_PATTERN = re.compile(
    r"\b[A-Z0-9]{1,5}(?:-[A-Z0-9]{1,5}){1,3}\b",
    re.IGNORECASE,
)

# ── Key-value extraction ───────────────────────────────────────────────────
# Matches lines like:
#   "Tag No.         : 303-FE-0701"
#   "Operating Temp  = 120"
#   "Flow Rate Min   120 Kg/h"       ← no separator, value follows label
KV_PATTERN = re.compile(
    r"^(?P<label>[A-Za-z][\w /.()\-]{1,60}?)"   # label (2–60 chars)
    r"(?:\s*[:=]\s*|\s{2,})"                      # separator: colon, equals, or 2+ spaces
    r"(?P<value>.+)$",
    re.MULTILINE,
)


# ── Record splitting ───────────────────────────────────────────────────────

def split_records(text: str) -> list[str]:
    """
    Split a multi-record PDF text blob into individual record strings.
    """

    if not text or not text.strip():
        return []

    positions = []

    for m in TAG_PATTERN.finditer(text):

        candidate = m.group(0)

        if _score_tag(candidate, text) >= 80:
            positions.append(m.start())

    if not positions:
        log.warning("No tag numbers found in PDF — treating as single record.")
        return [text.strip()]

    records = []

    for i, start in enumerate(positions):

        end = positions[i + 1] if i + 1 < len(positions) else len(text)

        chunk = text[start:end].strip()

        if chunk:
            records.append(chunk)

    log.info(f"split_records → {len(records)} record(s) found.")

    return records
# ── Key-value extraction from a single record ─────────────────────────────

def _extract_kv(record_text: str) -> dict:
    """
    Extract all key-value pairs from a single record's text block.

    Returns a raw dict: {pdf_label: value_string}
    """
    raw: dict = {}
    for m in KV_PATTERN.finditer(record_text):
        label = m.group("label").strip().rstrip(":= ")
        value = m.group("value").strip()
        if label and value and value.lower() not in ("", "n/a", "na", "-", "--"):
            raw[label] = value
    return raw

from rapidfuzz import fuzz


def _score_tag(candidate: str, text: str) -> int:
    """
    Score how likely a candidate is a real instrument tag.
    Higher score = better.
    """

    score = 0
    c = candidate.upper()

    # Strong pattern match
    if re.fullmatch(r"\d{3}-[A-Z]{2,4}-\d{3,4}[A-Z]?", c):
        score += 100

    # Appears near "Tag Number"
    tag_context = re.search(
        rf"Tag\s*(?:No|Number)?.{{0,40}}{re.escape(candidate)}",
        text,
        re.IGNORECASE,
    )

    if tag_context:
        score += 200

    # Penalize known junk patterns
    bad_words = [
        "MMH20",
        "SS316",
        "FEB",
        "EIL",
        "DOC",
    ]

    for bad in bad_words:
        if bad in c:
            score -= 100

    # Prefer FE / FT / LT / PT / TT instruments
    instrument_codes = [
        "FE", "FT", "LT", "PT", "TT",
        "TE", "PE", "FI", "PI"
    ]

    for code in instrument_codes:
        if f"-{code}-" in c:
            score += 50

    # Length sanity
    if len(c) < 7:
        score -= 40

    return score

def _find_tag(record_text: str) -> Optional[str]:
    """
    Find best instrument tag using fuzzy scoring.
    """

    candidates = TAG_PATTERN.findall(record_text)

    if not candidates:
        return None

    scored = []

    for c in candidates:
        s = _score_tag(c, record_text)
        scored.append((c, s))

    scored.sort(key=lambda x: x[1], reverse=True)

    best_tag, best_score = scored[0]

    log.info(f"Best tag candidate: {best_tag} (score={best_score})")

    # Reject weak matches
    if best_score < 80:
        return None

    return best_tag.upper()


# ── Main processor ────────────────────────────────────────────────────────

def process_record(record_text: str) -> dict:
    """
    Full processing pipeline for a single text record:
      1. Extract raw key-value pairs.
      2. Map PDF labels → canonical Excel column names.
      3. Build a full 129-column row dict (NaN for missing columns).
      4. Ensure tag_number is set (required field).

    Returns a row dict with canonical Excel column keys.
    """
    if not record_text.strip():
        return {"tag_number": None}

    # 1. Extract raw KV pairs
    raw = _extract_kv(record_text)

    # 2. Try to ensure tag number is captured
    #    (tag may appear before any KV line, so extract from text directly)
    tag = _find_tag(record_text)
    
    if tag:
        raw["TagNo"] = tag

    if not raw:
        log.warning("process_record: no KV pairs extracted from record.")
        return {"tag_number": None}

    # Log unmatched labels for debugging
    mapped = map_pdf_record(raw)
    unmatched = mapped.pop("__unmatched__", [])
    if unmatched:
        log.debug(f"Unmatched PDF labels (add to ALIASES if needed): {unmatched}")

    # 3. Build full row
    row = build_excel_row(mapped)

    # 4. Expose tag_number at top level for pipeline filtering
    row["tag_number"] = row.get("TagNo") or tag

    return row