"""Cross-cutting utilities: text normalization and document naming.

Sub-modules
-----------
text
    Pure text normalization helpers (:func:`normalize_text`,
    :func:`derive_product_from_document`).
naming
    Document-level naming helpers: canonical names, product extraction,
    translation, and sanitization.
"""

from .naming import (
    DEFAULT_PRODUCT_TRANSLATIONS,
    EXTRACTED_PAGES_RE,
    REVIEWED_PREFIX,
    REVIEWED_RE,
    SUFFIXES,
    canonical_document_name,
    extract_source_product,
    infer_yearbook_metadata,
    sanitize_name,
    strip_known_suffixes,
    translate_product_name,
)
from .text import derive_product_from_document, normalize_text

__all__ = [
    "DEFAULT_PRODUCT_TRANSLATIONS",
    "EXTRACTED_PAGES_RE",
    "REVIEWED_PREFIX",
    "REVIEWED_RE",
    "SUFFIXES",
    "canonical_document_name",
    "derive_product_from_document",
    "extract_source_product",
    "infer_yearbook_metadata",
    "normalize_text",
    "sanitize_name",
    "strip_known_suffixes",
    "translate_product_name",
]
