"""Backward-compatible re-export of :mod:`iia_excel_reorg.io.xlsx`.

.. deprecated::
    Import directly from :mod:`iia_excel_reorg.io` or
    :mod:`iia_excel_reorg.io.xlsx` in new code.
"""

from .io.xlsx import (  # noqa: F401
    CONTENT_NS,
    MAIN_NS,
    PKG_REL_NS,
    REL_NS,
    XML_NS,
    CellData,
    SheetData,
    WorkbookData,
    _collect_fill_styles,
    _column_index_from_letters,
    _column_letters,
    _normalize_rgb,
    _read_cell_value,
    _read_fill_map,
    _read_shared_strings,
    _render_content_types,
    _render_root_relationships,
    _render_sheet,
    _render_styles,
    _render_workbook,
    _render_workbook_relationships,
    _resolve_sheet_targets,
    _split_ref,
    read_workbook,
    write_workbook,
)
