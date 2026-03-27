from __future__ import annotations

from pathlib import Path

from iia_excel_reorg.footnote_pipeline import (
    apply_mapping_in_place,
    generate_mapping_template,
)
from iia_excel_reorg.xlsx_io import SheetData, WorkbookData, read_workbook, write_workbook


def _build_workbook(path: Path, sheet_name: str, rows: list[tuple[str, str]]) -> None:
    sheet = SheetData(name=sheet_name)
    sheet.set_row(1, ["hemisphere", "continent", "country", "unit", "footnotes", "1922"])
    for idx, (country, footnotes) in enumerate(rows, start=2):
        sheet.set_row(idx, ["HEMISPHERE", "EUROPE", country, "", footnotes, 1])
    write_workbook(path, WorkbookData(sheets=[sheet]))


def test_generate_mapping_template_scans_nested_tree_and_deduplicates(tmp_path: Path) -> None:
    source_root = tmp_path / "10-raw_imports"
    (source_root / "a" / "b").mkdir(parents=True)
    _build_workbook(
        source_root / "a" / "b" / "doc1.xlsx",
        "area",
        [
            ("Austria", "reexports; special case"),
            ("Belgium", "special case"),
        ],
    )
    _build_workbook(
        source_root / "doc2.xlsx",
        "imports",
        [("Canada", "unit note q; reexports")],
    )

    template_path = tmp_path / "footnote_mapping_template.xlsx"
    generate_mapping_template(source_root, template_path)

    template = read_workbook(template_path)
    sheet = template.sheets[0]
    originals = [
        str(sheet.get_cell(row, 1).value)
        for row in range(2, sheet.max_row + 1)
        if sheet.get_cell(row, 1).value
    ]
    assert originals == ["reexports", "special case", "unit note q"]
    assert all((sheet.get_cell(row, 2).value or "") == "" for row in range(2, 5))


def test_apply_mapping_in_place_supports_many_to_one(tmp_path: Path) -> None:
    source_root = tmp_path / "10-raw_imports"
    source_root.mkdir(parents=True)
    workbook_path = source_root / "doc.xlsx"
    _build_workbook(
        workbook_path,
        "area",
        [
            ("Austria", "r; reexports; note a"),
            ("Belgium", "note a"),
        ],
    )

    template_sheet = SheetData(name="footnote_mapping")
    template_sheet.set_row(1, ["Original Footnote", "Cleaned Footnote"])
    template_sheet.set_row(2, ["r", "reexports"])
    template_sheet.set_row(3, ["reexports", "reexports"])
    template_sheet.set_row(4, ["note a", "special note"])
    template_path = tmp_path / "mapping.xlsx"
    write_workbook(template_path, WorkbookData(sheets=[template_sheet]))

    changed = apply_mapping_in_place(source_root, template_path)
    assert changed == [workbook_path]

    workbook = read_workbook(workbook_path)
    sheet = workbook.sheets[0]
    assert sheet.get_cell(2, 5).value == "reexports; special note"
    assert sheet.get_cell(3, 5).value == "special note"

