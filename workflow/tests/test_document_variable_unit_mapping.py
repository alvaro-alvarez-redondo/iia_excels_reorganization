from __future__ import annotations

from pathlib import Path

from iia_excel_reorg.config import load_config
from iia_excel_reorg.io.xlsx import SheetData, WorkbookData, write_workbook


def test_load_config_reads_document_variable_unit_mapping() -> None:
    project_root = Path(__file__).resolve().parents[2]
    mapping_path = project_root / "data" / "document_variable_unit_mapping.xlsx"
    mapping_path.parent.mkdir(parents=True, exist_ok=True)

    mapping_sheet = SheetData(name="mapping")
    mapping_sheet.set_row(1, ["document", "variable", "unit"])
    mapping_sheet.set_row(2, ["reviewed_123_124coffee", "IMPORTS", "tonnes"])
    mapping_sheet.set_row(3, ["reviewed_123_124coffee.xlsx", "area", "ha"])
    write_workbook(mapping_path, WorkbookData(sheets=[mapping_sheet]))

    try:
        config = load_config(None)
    finally:
        mapping_path.unlink(missing_ok=True)

    assert (
        config.mapped_unit_for("reviewed_123_124coffee.xlsx", "imports") == "tonnes"
    )
    assert config.mapped_unit_for("reviewed_123_124coffee", "AREA") == "ha"
