"""Helpers for deriving stable canonical document and product names."""

from __future__ import annotations

import importlib
import importlib.util
import json
import re
from functools import lru_cache
from pathlib import Path
from urllib.parse import quote
from urllib.request import urlopen

from .text import normalize_text

REVIEWED_PREFIX = "reviewed_"
REVIEWED_RE = re.compile(
    r"^reviewed_(?P<start>\d+)_(?P<end>\d+)(?P<body>.+)$",
    re.IGNORECASE,
)
EXTRACTED_PAGES_RE = re.compile(
    r"^extracted_pages_(?P<year>\d{4})_\d{2}$",
    re.IGNORECASE,
)
SUFFIXES = ("sup", "prod", "rend", "imp", "exp", "num")
_MULTI_UNDERSCORE_RE = re.compile(r"_+")
_SUFFIX_RE = re.compile(
    r"_?(?:" + "|".join(re.escape(suffix) for suffix in SUFFIXES) + r")$",
    re.IGNORECASE,
)


def sanitize_name(name: str) -> str:
    """Return a filesystem-safe identifier derived from *name*."""
    result = name.replace(" ", "_")
    result = _MULTI_UNDERSCORE_RE.sub("_", result)
    return result.strip("_")


DEFAULT_PRODUCT_TRANSLATIONS: dict[str, str] = {}


@lru_cache(maxsize=512)
def _auto_translate_product(raw_product: str) -> str:
    """Translate *raw_product* to English with cached network fallbacks."""
    normalized_product = normalize_text(raw_product)
    if not normalized_product:
        return ""
    translated = _translate_with_deep_translator(normalized_product)
    if not translated:
        translated = _translate_with_mymemory(normalized_product)
    return normalize_text(translated) or normalized_product


def _translate_with_deep_translator(normalized_product: str) -> str:
    """Attempt translation through the ``deep_translator`` package."""
    if importlib.util.find_spec("deep_translator") is None:
        return ""

    translator_module = importlib.import_module("deep_translator")
    translator = translator_module.GoogleTranslator(source="auto", target="en")
    try:
        return str(translator.translate(normalized_product) or "")
    except Exception:
        return ""


def _translate_with_mymemory(normalized_product: str) -> str:
    """Attempt translation through the MyMemory public API."""
    url = (
        "https://api.mymemory.translated.net/get"
        f"?q={quote(normalized_product)}&langpair=auto|en"
    )
    try:
        with urlopen(url, timeout=5) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except Exception:
        return ""

    translated = payload.get("responseData", {}).get("translatedText", "")
    return str(translated or "")


def strip_known_suffixes(raw_product: str) -> str:
    """Strip repeated trailing agricultural-trade suffixes from *raw_product*."""
    cleaned = raw_product.strip("_")
    while True:
        stripped = _SUFFIX_RE.sub("", cleaned).strip("_")
        if stripped == cleaned:
            return cleaned
        cleaned = stripped


def _normalized_mapping(mapping: dict[str, str] | None) -> dict[str, str]:
    """Return a normalized-key/value copy of *mapping*."""
    return {
        normalize_text(key): normalize_text(value)
        for key, value in (mapping or {}).items()
    }


def translate_product_name(
    raw_product: str,
    product_translations: dict[str, str] | None = None,
) -> str:
    """Return the English translation of *raw_product*."""
    normalized_product = normalize_text(raw_product)
    translations = {
        **DEFAULT_PRODUCT_TRANSLATIONS,
        **_normalized_mapping(product_translations),
    }
    return translations.get(normalized_product, _auto_translate_product(normalized_product))


def extract_source_product(document_path: str | Path) -> str:
    """Derive the product name embedded in a source Excel filename."""
    stem = Path(document_path).stem
    match = REVIEWED_RE.match(stem)
    if match:
        body = strip_known_suffixes(match.group("body"))
        return normalize_text(body)

    tokens = [token for token in stem.split("_") if token]
    if not tokens:
        return ""
    year_index = next(
        (
            index
            for index, token in enumerate(tokens)
            if len(token) == 4 and token.isdigit()
        ),
        None,
    )
    if year_index is None:
        return normalize_text(stem)
    product_start = year_index + 1
    while product_start < len(tokens) and tokens[product_start].isdigit():
        product_start += 1
    return normalize_text(" ".join(tokens[product_start:] or tokens[-1:]))


def canonical_document_name(
    document_path: str | Path,
    product_translations: dict[str, str] | None = None,
    product_aliases: dict[str, str] | None = None,
) -> str:
    """Compute a stable, human-readable canonical name for a source workbook."""
    path = Path(document_path)
    stem = path.stem
    if stem.startswith("r_") and not stem.startswith(REVIEWED_PREFIX):
        return stem.lower()

    match = REVIEWED_RE.match(stem)
    if match is None:
        return stem.lower()

    metadata = infer_yearbook_metadata(path)
    source_product = extract_source_product(path)
    canonical_product = _normalized_mapping(product_aliases).get(
        source_product,
        source_product,
    )
    english_product = translate_product_name(canonical_product, product_translations)
    raw_name = (
        f"r_iia_{metadata['yearbook']}_{metadata['year']}_{match.group('start')}_"
        f"{match.group('end')}_{english_product.replace(' ', '_')}"
    )
    return sanitize_name(raw_name)


def infer_yearbook_metadata(document_path: str | Path) -> dict[str, str]:
    """Infer ``agency``, ``yearbook``, and ``year`` from *document_path*."""
    path = Path(document_path)
    yearbook = "unknown"
    year = "unknown"
    for index, part in enumerate(path.parts):
        match = EXTRACTED_PAGES_RE.match(part)
        if match is None:
            continue
        year = match.group("year")
        if index > 0:
            yearbook = sanitize_name(normalize_text(path.parts[index - 1]))
        break
    return {"agency": "iia", "yearbook": yearbook, "year": year}
