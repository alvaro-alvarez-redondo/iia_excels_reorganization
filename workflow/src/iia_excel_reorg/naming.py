"""Backward-compatible re-export of :mod:`iia_excel_reorg.utils.naming`.

.. deprecated::
    Import directly from :mod:`iia_excel_reorg.utils` or
    :mod:`iia_excel_reorg.utils.naming` in new code.
"""

from .utils.naming import (  # noqa: F401
    DEFAULT_PRODUCT_TRANSLATIONS,
    EXTRACTED_PAGES_RE,
    REVIEWED_PREFIX,
    REVIEWED_RE,
    SUFFIXES,
    _auto_translate_product,
    _translate_with_deep_translator,
    _translate_with_mymemory,
    canonical_document_name,
    extract_source_product,
    infer_yearbook_metadata,
    sanitize_name,
    strip_known_suffixes,
    translate_product_name,
)
