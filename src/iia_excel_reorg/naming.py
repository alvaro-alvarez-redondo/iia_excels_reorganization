from __future__ import annotations

import re
from pathlib import Path

from .unit_rules import normalize_text

REVIEWED_PREFIX = "reviewed_"
REVIEWED_RE = re.compile(r"^reviewed_(?P<start>\d+)_(?P<end>\d+)(?P<body>.+)$", re.IGNORECASE)
EXTRACTED_PAGES_RE = re.compile(r"^extracted_pages_(?P<year>\d{4})_\d{2}$", re.IGNORECASE)
SUFFIXES = ("sup", "prod", "rend", "imp", "exp", "num")
DEFAULT_PRODUCT_TRANSLATIONS = {
    "azucar cana bruta": "raw cane sugar",
    "arroz": "rice",
    "te": "tea",
}



def strip_known_suffixes(raw_product: str) -> str:
    cleaned = raw_product.strip("_")
    changed = True
    while cleaned and changed:
        changed = False
        for suffix in SUFFIXES:
            if cleaned.endswith(f"_{suffix}"):
                cleaned = cleaned[: -len(suffix) - 1].rstrip("_")
                changed = True
                break
            if cleaned.endswith(suffix):
                cleaned = cleaned[: -len(suffix)].rstrip("_")
                changed = True
                break
    return cleaned.strip("_")



def extract_source_product(document_path: str | Path) -> str:
    stem = Path(document_path).stem
    match = REVIEWED_RE.match(stem)
    if match:
        body = strip_known_suffixes(match.group("body"))
        return normalize_text(body)

    tokens = [token for token in stem.split("_") if token]
    if not tokens:
        return ""
    year_idx = next((idx for idx, token in enumerate(tokens) if len(token) == 4 and token.isdigit()), None)
    if year_idx is None:
        return normalize_text(stem)
    product_start = year_idx + 1
    while product_start < len(tokens) and tokens[product_start].isdigit():
        product_start += 1
    product_tokens = tokens[product_start:] or tokens[-1:]
    return normalize_text(" ".join(product_tokens))



def canonical_document_name(document_path: str | Path, product_translations: dict[str, str] | None = None) -> str:
    path = Path(document_path)
    stem = path.stem
    if stem.startswith("r_") and not stem.startswith(REVIEWED_PREFIX):
        return stem.lower()

    match = REVIEWED_RE.match(stem)
    if not match:
        return stem.lower()

    metadata = infer_yearbook_metadata(path)
    source_product = extract_source_product(path)
    translations = {
        **DEFAULT_PRODUCT_TRANSLATIONS,
        **{normalize_text(key): normalize_text(value) for key, value in (product_translations or {}).items()},
    }
    english_product = translations.get(source_product, source_product)
    product_slug = english_product.replace(" ", "_")
    return f"r_iia_{metadata['yearbook']}_{metadata['year']}_{match.group('start')}_{match.group('end')}_{product_slug}"



def infer_yearbook_metadata(document_path: str | Path) -> dict[str, str]:
    path = Path(document_path)
    yearbook = "unknown"
    year = "unknown"
    parts = list(path.parts)
    for idx, part in enumerate(parts):
        match = EXTRACTED_PAGES_RE.match(part)
        if match:
            year = match.group("year")
            if idx > 0:
                yearbook = normalize_text(parts[idx - 1]).replace(" ", "_")
            break
    return {"agency": "iia", "yearbook": yearbook, "year": year}
