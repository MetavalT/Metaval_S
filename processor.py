import re
from typing import Optional
from logger import get_logger
from field_map import map_pdf_record, build_excel_row

log = get_logger()

# ── TAG PATTERN ───────────────────────────────────────────
TAG_PATTERN = re.compile(
    r"\b[A-Z0-9]{1,5}(?:-[A-Z0-9]{1,5}){1,3}\b",
    re.IGNORECASE,
)

# ── VALIDATION ────────────────────────────────────────────
def is_valid_value(field, value):
    if not isinstance(value, str):
        return False

    value = value.strip()

    if not value:
        return False

    # Critical numeric fields must contain numbers
    if any(k in field.lower() for k in ["pressure", "temp", "flow", "rate", "density"]):
        return bool(re.search(r"\d", value))

    return True


# ── RECORD SPLITTING ──────────────────────────────────────
def split_records(text: str) -> list[str]:

    if not text or not text.strip():
        return []

    positions = []

    for m in TAG_PATTERN.finditer(text):
        candidate = m.group(0)

        if _score_tag(candidate, text) >= 120:
            positions.append(m.start())

    if not positions:
        log.warning("No tag numbers found")
        return [text.strip()]

    records = []

    for i, start in enumerate(positions):
        end = positions[i + 1] if i + 1 < len(positions) else len(text)
        chunk = text[start:end].strip()

        if chunk:
            records.append(chunk)

    log.info(f"split_records → {len(records)} records")

    return records


# ── SMART KV EXTRACTION (CORE FIX) ─────────────────────────
def _extract_kv(record_text: str) -> dict:

    raw = {}
    lines = record_text.split("\n")

    current_label = None

    for line in lines:
        line = line.strip()

        if not line:
            continue

        # Try structured split (OCR tables)
        parts = re.split(r"\s{2,}", line)

        if len(parts) >= 2:
            label = parts[0].strip()
            value = " ".join(parts[1:]).strip()

            current_label = label
            raw[current_label] = value
            continue

        # Try colon format
        if ":" in line:
            parts = line.split(":", 1)
            label = parts[0].strip()
            value = parts[1].strip()

            current_label = label
            raw[current_label] = value
            continue

        # MULTI-LINE VALUE HANDLING
        if current_label:
            raw[current_label] += " " + line.strip()

    return raw


# ── TAG SCORING ───────────────────────────────────────────
def _score_tag(candidate: str, text: str) -> int:

    score = 0
    c = candidate.upper()

    if re.fullmatch(r"\d{3}-[A-Z]{2,4}-\d{3,4}[A-Z]?", c):
        score += 100

    if re.search(rf"Tag.*{re.escape(candidate)}", text, re.IGNORECASE):
        score += 200

    bad_words = ["MMH20", "SS316", "FEB", "EIL", "DOC"]

    for bad in bad_words:
        if bad in c:
            score -= 100

    if len(c) < 7:
        score -= 40

    return score


# ── FIND TAG ───────────────────────────────────────────────
def _find_tag(record_text: str) -> Optional[str]:

    candidates = TAG_PATTERN.findall(record_text)

    if not candidates:
        return None

    scored = [(c, _score_tag(c, record_text)) for c in candidates]
    scored.sort(key=lambda x: x[1], reverse=True)

    best_tag, best_score = scored[0]

    log.info(f"Best tag: {best_tag} (score={best_score})")

    if best_score < 80:
        return None

    return best_tag.upper()


# ── MAIN PROCESSOR ─────────────────────────────────────────
def process_record(record_text: str) -> dict:

    if not record_text.strip():
        return {"tag_number": None}

    # 1. Extract KV
    raw = _extract_kv(record_text)
    print("RAW KV:", raw)

    # 2. Extract tag
    tag = _find_tag(record_text)
    print("FINAL TAG:", tag)

    if tag:
        raw["TagNo"] = tag

    # FAIL SAFE
    if len(raw) < 3:
        log.warning("Too few fields extracted — skipping record")
        return {"tag_number": None}

    # 3. Map fields
    mapped = map_pdf_record(raw)
    unmatched = mapped.pop("__unmatched__", [])

    if unmatched:
        print("⚠ Unmapped fields:", unmatched)

    # 4. VALIDATION
    cleaned = {}

    for field, value in mapped.items():
        if is_valid_value(field, value):
            cleaned[field] = value
        else:
            log.warning(f"Invalid removed → {field}: {value}")

    # FAIL SAFE AGAIN
    if len(cleaned) < 3:
        log.warning("Too few valid fields after cleaning — skipping")
        return {"tag_number": None}

    # 5. Build row
    row = build_excel_row(cleaned)

    if tag:
        row["TagNo"] = tag

    row["tag_number"] = tag

    print("ROW TAG:", row.get("tag_number"))

    return row