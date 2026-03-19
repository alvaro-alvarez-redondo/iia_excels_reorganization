"""Backward-compatible re-export of :mod:`iia_excel_reorg.core.transformer`.

.. deprecated::
    Import directly from :mod:`iia_excel_reorg.core` or
    :mod:`iia_excel_reorg.core.transformer` in new code.
"""

from .core.transformer import (  # noqa: F401
    HEADER_COLUMNS,
    HEADER_FILL,
    HEMISPHERE_RE,
    KNOWN_CONTINENTS,
    KNOWN_HEMISPHERES,
    PAREN_RE,
    RAW_CONTINENT_LABELS,
    RAW_HEMISPHERE_LABELS,
    GeographyIndex,
    ProductIndex,
    TransformationError,
    _clean_text,
    _extract_country_and_footnotes,
    _extract_year_headers,
    _is_continent_row,
    _is_hemisphere_row,
    _normalize_footnote,
    _normalize_known_geography_label,
    _stringify_header,
    _strip_terminal_punctuation,
    _transform_sheet,
    _write_headers,
    transform_workbook,
)
