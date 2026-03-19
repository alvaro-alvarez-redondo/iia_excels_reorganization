from __future__ import annotations

import os
from pathlib import Path

from iia_excel_reorg.cli import _compute_output_subdir, _ensure_workspace, _iter_workbooks_structured
from iia_excel_reorg.config import WorkbookConfig, load_config
from iia_excel_reorg.naming import canonical_document_name, extract_source_product, infer_yearbook_metadata, sanitize_name
from iia_excel_reorg.transformer import GeographyIndex, _is_continent_row, _is_hemisphere_row, transform_workbook
from iia_excel_reorg.unit_rules import assign_unit
from iia_excel_reorg.xlsx_io import SheetData, WorkbookData, read_workbook, write_workbook

GREEN = "FF00FF00"
YELLOW = "FFFFFF00"
ORANGE = "FFFFA500"



def _build_source_workbook(path: Path, *, include_imports: bool = False) -> None:
    area = SheetData(name="AREA")
    area.set_cell(1, 2, "1909-1913")
    area.set_cell(1, 3, "1922")
    area.set_cell(2, 1, "HÉMISPHÈRE SEPTENTRIONAL")
    area.set_cell(3, 1, "EUROPE.")
    area.set_cell(4, 1, "Belgique-Luxembourg (reexports) (special case)", fill_rgb=GREEN)
    area.set_cell(4, 2, 17268, fill_rgb=YELLOW)
    area.set_cell(4, 3, 11887, fill_rgb=ORANGE)
    area.set_cell(5, 1, "Germany", fill_rgb=GREEN)
    area.set_cell(5, 2, 284000, fill_rgb=GREEN)

    sheets = [area]

    production = SheetData(name="PRODUCTION")
    production.set_cell(1, 2, "1909-10/1913")
    production.set_cell(1, 3, "1922-23")
    production.set_cell(2, 1, "Hemisphère méridional")
    production.set_cell(3, 1, "Amérique")
    production.set_cell(4, 1, "Canada", fill_rgb=GREEN)
    production.set_cell(4, 2, 194876, fill_rgb=GREEN)
    production.set_cell(4, 3, 315569, fill_rgb=YELLOW)
    sheets.append(production)

    if include_imports:
        imports = SheetData(name="IMPORTS")
        imports.set_cell(1, 2, "1934-1938")
        imports.set_cell(1, 3, "1946")
        imports.set_cell(2, 1, "HÉMISPHÈRE SEPTENTRIONAL")
        imports.set_cell(3, 1, "EUROPE")
        imports.set_cell(4, 1, "Austria (unit note q)", fill_rgb=GREEN)
        imports.set_cell(4, 2, 7.5, fill_rgb=GREEN)
        imports.set_cell(4, 3, 0.2, fill_rgb=YELLOW)
        sheets.append(imports)

    write_workbook(path, WorkbookData(sheets=sheets))


def _build_numeric_year_workbook(path: Path) -> None:
    area = SheetData(name="AREA")
    area.set_cell(1, 2, 1900.0)
    area.set_cell(2, 1, "HÉMISPHÈRE SEPTENTRIONAL", fill_rgb=YELLOW)
    area.set_cell(3, 1, "EUROPE.", fill_rgb=ORANGE)
    area.set_cell(4, 1, "Austria (r)", fill_rgb=GREEN)
    area.set_cell(4, 2, 12, fill_rgb=GREEN)
    write_workbook(path, WorkbookData(sheets=[area]))





def test_geography_detection_handles_known_ocr_and_accent_variants() -> None:
    assert _is_continent_row("AMÉR DU NORD ET AMÉR CENTRALE")
    assert _is_continent_row("Amérique méridionale")
    assert _is_continent_row("OCRANIE.")
    assert _is_continent_row("AUSTRALIE")
    assert _is_hemisphere_row("HÉMISHPÈRE SEPTENTRIONAL")
    assert _is_hemisphere_row("HÊMISPHÊRE SEPTENTRIONAL")
    assert _is_hemisphere_row("HÉMISPHÈRE SUDAFRIQUE.")
    assert not _is_continent_row("Canada")
    assert not _is_hemisphere_row("Canada")


def test_transform_workbook_assigns_units_from_rules_and_preserves_notes(tmp_path: Path) -> None:
    source_path = tmp_path / "r_iia_trade_1950_3_5_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_source_workbook(source_path)

    config_path = tmp_path / "config.yml"
    config_path.write_text(
        "\n".join(
            [
                'unit_mode: standard',
                'document_categories:',
                '  r_iia_trade_1950_3_5_wheat: 1',
            ]
        ),
        encoding="utf-8",
    )

    transform_workbook(source_path, output_path, config=load_config(config_path))

    result = read_workbook(output_path)
    assert [sheet.name for sheet in result.sheets] == ["area", "production"]

    area = result.sheets[0]
    assert [area.get_cell(1, idx).value for idx in range(1, 8)] == [
        "hemisphere",
        "continent",
        "country",
        "unit",
        "footnotes",
        "1909-1913",
        "1922",
    ]
    assert area.get_cell(2, 1).value == "HÉMISPHÈRE SEPTENTRIONAL"
    assert area.get_cell(2, 2).value == "EUROPE"
    assert area.get_cell(2, 3).value == "Belgique-Luxembourg"
    assert area.get_cell(2, 4).value == "1000 ha"
    assert area.get_cell(2, 5).value == "reexports; special case"
    assert area.get_cell(2, 6).value == 17268
    assert area.get_cell(2, 7).value == 11887
    assert area.get_cell(2, 1).fill_rgb is None
    assert area.get_cell(2, 2).fill_rgb is None
    assert area.get_cell(2, 3).fill_rgb == GREEN
    assert area.get_cell(2, 4).fill_rgb is None
    assert area.get_cell(2, 5).fill_rgb is None
    assert area.get_cell(2, 6).fill_rgb == YELLOW
    assert area.get_cell(2, 7).fill_rgb == ORANGE

    production = result.sheets[1]
    assert production.get_cell(2, 1).value == "Hemisphère méridional"
    assert production.get_cell(2, 2).value == "Amérique"
    assert production.get_cell(2, 3).value == "Canada"
    assert production.get_cell(2, 4).value == "1000 q"
    assert production.get_cell(2, 6).value == 194876
    assert production.get_cell(2, 7).value == 315569


def test_transform_workbook_preserves_group_colors_and_normalizes_numeric_year_headers(tmp_path: Path) -> None:
    source_path = tmp_path / "r_iia_trade_1950_1_1_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_numeric_year_workbook(source_path)

    config_path = tmp_path / "config.yml"
    config_path.write_text(
        "\n".join(
            [
                "unit_mode: standard",
                "document_categories:",
                "  r_iia_trade_1950_1_1_wheat: 1",
            ]
        ),
        encoding="utf-8",
    )

    transform_workbook(source_path, output_path, config=load_config(config_path))

    result = read_workbook(output_path)
    area = result.sheets[0]
    assert area.get_cell(1, 6).value == "1900"
    assert area.get_cell(2, 1).fill_rgb == YELLOW
    assert area.get_cell(2, 2).fill_rgb == ORANGE
    assert area.get_cell(2, 3).fill_rgb == GREEN
    assert area.get_cell(2, 4).fill_rgb is None
    assert area.get_cell(2, 5).fill_rgb is None
    assert area.get_cell(2, 5).value == "reexports"


def test_transform_workbook_collects_unique_geography_labels(tmp_path: Path) -> None:
    source_path = tmp_path / "r_iia_trade_1950_3_5_wheat.xlsx"
    output_path = tmp_path / "standardized.xlsx"
    _build_source_workbook(source_path, include_imports=True)
    geography_index = GeographyIndex()

    config_path = tmp_path / "config.yml"
    config_path.write_text(
        "\n".join(
            [
                "unit_mode: standard",
                "document_categories:",
                "  r_iia_trade_1950_3_5_wheat: 1",
            ]
        ),
        encoding="utf-8",
    )

    transform_workbook(source_path, output_path, config=load_config(config_path), geography_index=geography_index)

    assert geography_index.hemispheres == {"HÉMISPHÈRE SEPTENTRIONAL", "Hemisphère méridional"}
    assert geography_index.continents == {"EUROPE", "Amérique"}
    assert geography_index.countries == {"Austria", "Belgique-Luxembourg", "Canada", "Germany"}

    index_path = tmp_path / "unique_geography_values.txt"
    geography_index.write_txt(index_path)
    assert index_path.read_text(encoding="utf-8") == "\n".join(
        [
            "[hemispheres]",
            "Hemisphère méridional",
            "HÉMISPHÈRE SEPTENTRIONAL",
            "",
            "[continents]",
            "Amérique",
            "EUROPE",
            "",
            "[countries]",
            "Austria",
            "Belgique-Luxembourg",
            "Canada",
            "Germany",
            "",
        ]
    )



def test_transform_workbook_supports_inputs_mode_and_harmonized_output_names(tmp_path: Path) -> None:
    source_dir = tmp_path / "raw_inputs" / "trade" / "extracted_pages_1938_39"
    source_dir.mkdir(parents=True)
    source_path = source_dir / "reviewed_466_475arrozimp_exp.xlsx"
    output_dir = tmp_path / "out"
    output_dir.mkdir()
    _build_source_workbook(source_path, include_imports=True)

    config_path = tmp_path / "config.yml"
    config_path.write_text(
        "\n".join(
            [
                'unit_mode: inputs',
                'document_categories:',
                '  reviewed_466_475arrozimp_exp: 2',
                'product_translations:',
                '  arroz: rice',
            ]
        ),
        encoding="utf-8",
    )

    config = load_config(config_path)
    output_path = output_dir / f"{config.canonical_name_for_document(source_path)}.xlsx"
    transform_workbook(source_path, output_path, config=config)

    assert output_path.name == "r_iia_trade_1938_466_475_rice.xlsx"
    result = read_workbook(output_path)
    imports = result.sheets[2]
    assert imports.name == "imports"
    assert imports.get_cell(2, 3).value == "Austria"
    assert imports.get_cell(2, 4).value == "1000 kg"
    assert imports.get_cell(2, 5).value == "unit note q"
    assert imports.get_cell(2, 6).value == 7.5
    assert imports.get_cell(2, 7).value == 0.2



def test_canonical_document_name_auto_translates_unknown_products(monkeypatch) -> None:
    from iia_excel_reorg import naming

    monkeypatch.setattr(naming, "_auto_translate_product", lambda value: "cocoa beans")
    path = Path("raw_inputs/trade/extracted_pages_1938_39/reviewed_12_13cacaoimp.xlsx")

    assert canonical_document_name(path) == "r_iia_trade_1938_12_13_cocoa_beans"


def test_canonical_document_name_applies_alias_before_translation() -> None:
    path = Path("raw_inputs/trade/extracted_pages_1938_39/reviewed_12_13teaimp.xlsx")

    assert canonical_document_name(path, product_aliases={"tea": "te"}) == "r_iia_trade_1938_12_13_tea"


def test_naming_and_unit_rules_cover_reviewed_documents() -> None:
    path = Path("raw_inputs/trade/extracted_pages_1938_39/reviewed_239_239azucar_caña_brutaprod.xlsx")
    assert infer_yearbook_metadata(path) == {"agency": "iia", "yearbook": "trade", "year": "1938"}
    assert extract_source_product(path) == "azucar cana bruta"
    assert canonical_document_name(path) == "r_iia_trade_1938_239_239_raw_cane_sugar"
    assert assign_unit("imports", "te", 1) == "tonnes"
    assert assign_unit("imports", "te", 2) == "q"
    assert assign_unit("production", "vino", 1) == "1000 hl"
    assert assign_unit("production", "huevos", 2) == "1000 eggs"
    assert assign_unit("livestock", "whatever", 1) == "1000 heads"
    assert assign_unit("production", "whatever", None) == ""



def test_load_config_parses_rule_based_yaml(tmp_path: Path) -> None:
    config_path = tmp_path / "units.yml"
    config_path.write_text(
        "\n".join(
            [
                'unit_mode: standard',
                'document_categories:',
                '  reviewed_466_475arrozimp_exp: 1',
                'product_aliases:',
                '  tea: te',
                'product_translations:',
                '  arroz: rice',
                'unit_overrides:',
                '  imports: tonnes',
                'include_sheets:',
                '  - AREA',
                '  - PRODUCTION',
            ]
        ),
        encoding="utf-8",
    )

    config = load_config(config_path)
    assert config.unit_mode == "standard"
    assert config.document_categories["reviewed_466_475arrozimp_exp"] == 1
    assert config.product_aliases["tea"] == "te"
    assert config.product_translations["arroz"] == "rice"
    assert config.unit_overrides["imports"] == "tonnes"
    assert config.include_sheets == ["AREA", "PRODUCTION"]


def test_workbook_config_canonical_name_uses_product_aliases() -> None:
    config = WorkbookConfig(product_aliases={"tea": "te"})
    path = Path("raw_inputs/trade/extracted_pages_1938_39/reviewed_12_13teaimp.xlsx")

    assert config.canonical_name_for_document(path) == "r_iia_trade_1938_12_13_tea"



def test_compute_output_subdir_with_extracted_pages_and_subfolder() -> None:
    # Excel file nested under extracted_pages_YYYY_YY/subfolder/
    path = Path("inputs/reviewed_iia/extracted_pages_1929_30/crops/reviewed_1_2_wheat.xlsx")
    result = _compute_output_subdir(path)
    assert result == Path("iia_extracted_pages_1929/iia_crops_1929")



def test_compute_output_subdir_with_extracted_pages_no_subfolder() -> None:
    # Excel file directly inside extracted_pages_YYYY_YY/ with a topic folder above.
    # The parent folder (trade) is used as the output subfolder.
    path = Path("raw_inputs/trade/extracted_pages_1938_39/reviewed_466_475arroz.xlsx")
    result = _compute_output_subdir(path)
    assert result == Path("iia_extracted_pages_1938/iia_trade_1938")



def test_compute_output_subdir_with_deep_nesting() -> None:
    # Excel file two levels below the topic root: topic/sub_category/extracted_pages_*/file.xlsx
    # The sub_category folder (directly above extracted_pages_*) becomes the output subfolder.
    path = Path("raw_inputs/area and production/multiple product/extracted_pages_1933_34/wb.xlsx")
    result = _compute_output_subdir(path)
    assert result == Path("iia_extracted_pages_1933/iia_multiple_product_1933")



def test_compute_output_subdir_without_extracted_pages() -> None:
    # Excel file with no extracted_pages_* segment → placed in output root
    path = Path("some/other/dir/workbook.xlsx")
    result = _compute_output_subdir(path)
    assert result == Path(".")



def test_iter_workbooks_structured_builds_correct_hierarchy(tmp_path: Path) -> None:
    # Set up an input tree with two extracted_pages folders and a subfolder
    crops_dir = tmp_path / "reviewed_iia" / "extracted_pages_1929_30" / "crops"
    trade_dir = tmp_path / "reviewed_iia" / "extracted_pages_1938_39" / "trade"
    crops_dir.mkdir(parents=True)
    trade_dir.mkdir(parents=True)

    wb1 = crops_dir / "reviewed_1_2_wheat.xlsx"
    wb2 = trade_dir / "reviewed_3_4_rice.xlsx"

    # Create minimal xlsx files so rglob finds them
    _build_source_workbook(wb1)
    _build_source_workbook(wb2)

    entries = _iter_workbooks_structured(tmp_path)

    paths_and_subdirs = {e[0].name: e[1] for e in entries}
    assert paths_and_subdirs["reviewed_1_2_wheat.xlsx"] == Path("iia_extracted_pages_1929/iia_crops_1929")
    assert paths_and_subdirs["reviewed_3_4_rice.xlsx"] == Path("iia_extracted_pages_1938/iia_trade_1938")



def test_cli_main_creates_structured_output(tmp_path: Path) -> None:
    """end-to-end: main() populates the iia_extracted_pages_*/iia_*_* hierarchy."""
    from iia_excel_reorg.cli import main

    crops_dir = tmp_path / "inputs" / "reviewed_iia" / "extracted_pages_1929_30" / "crops"
    crops_dir.mkdir(parents=True)
    source = crops_dir / "reviewed_1_2_wheat.xlsx"
    _build_source_workbook(source)

    output_root = tmp_path / "outputs"

    config_path = tmp_path / "config.yml"
    config_path.write_text(
        "\n".join(
            [
                "unit_mode: standard",
                "document_categories:",
                "  reviewed_1_2_wheat: 1",
            ]
        ),
        encoding="utf-8",
    )

    import sys
    orig_argv = sys.argv
    try:
        sys.argv = [
            "iia-excel-reorg",
            str(tmp_path / "inputs"),
            str(output_root),
            "--config",
            str(config_path),
        ]
        main()
    finally:
        sys.argv = orig_argv

    # The transformed file must land in iia_extracted_pages_1929/iia_crops_1929/
    output_subdir = output_root / "iia_extracted_pages_1929" / "iia_crops_1929"
    assert output_subdir.is_dir(), f"Expected output subdir not found: {output_subdir}"
    xlsx_files = list(output_subdir.glob("*.xlsx"))
    assert len(xlsx_files) == 1


def test_ensure_workspace_creates_missing_input_and_output_dirs(tmp_path: Path) -> None:
    input_dir = tmp_path / "data" / "raw_inputs"
    output_dir = tmp_path / "data" / "10-raw_imports"

    _ensure_workspace(input_dir, output_dir)

    assert input_dir.is_dir()
    assert output_dir.is_dir()


def test_ensure_workspace_overwrites_existing_output_dir(tmp_path: Path) -> None:
    input_dir = tmp_path / "data" / "raw_inputs"
    output_dir = tmp_path / "data" / "10-raw_imports"
    output_dir.mkdir(parents=True)
    stale_file = output_dir / "old.txt"
    stale_file.write_text("stale", encoding="utf-8")

    _ensure_workspace(input_dir, output_dir)

    assert input_dir.is_dir()
    assert output_dir.is_dir()
    assert not stale_file.exists()


def test_cli_main_creates_default_workspace_and_exits_cleanly_when_input_is_empty(tmp_path: Path, capsys) -> None:
    from iia_excel_reorg.cli import main

    config_path = tmp_path / "config.yml"
    config_path.write_text("unit_mode: standard\n", encoding="utf-8")

    import sys
    orig_argv = sys.argv
    orig_cwd = Path.cwd()
    try:
        sys.argv = [
            "iia-excel-reorg",
            str(tmp_path / "data" / "raw_inputs"),
            str(tmp_path / "data" / "10-raw_imports"),
            "--config",
            str(config_path),
        ]
        os.chdir(tmp_path)
        main()
    finally:
        sys.argv = orig_argv
        os.chdir(orig_cwd)

    captured = capsys.readouterr()
    assert "No Excel workbooks found in:" in captured.out
    assert (tmp_path / "data" / "raw_inputs").is_dir()
    assert (tmp_path / "data" / "10-raw_imports").is_dir()


def test_cli_main_reports_progress_bars(tmp_path: Path, capsys) -> None:
    from iia_excel_reorg.cli import main

    crops_dir = tmp_path / "inputs" / "reviewed_iia" / "extracted_pages_1929_30" / "crops"
    crops_dir.mkdir(parents=True)
    source = crops_dir / "reviewed_1_2_wheat.xlsx"
    _build_source_workbook(source)

    config_path = tmp_path / "config.yml"
    config_path.write_text(
        "\n".join(
            [
                "unit_mode: standard",
                "document_categories:",
                "  reviewed_1_2_wheat: 1",
            ]
        ),
        encoding="utf-8",
    )

    import sys
    orig_argv = sys.argv
    try:
        sys.argv = ["iia-excel-reorg", str(tmp_path / "inputs"), str(tmp_path / "outputs"), "--config", str(config_path)]
        main()
    finally:
        sys.argv = orig_argv

    captured = capsys.readouterr()
    assert captured.out == (
        "Reorganizing folders: [------------------------] 0/1\r"
        "Reorganizing folders: [########################] 1/1\n"
        "Reorganizing excels: [------------------------] 0/1\r"
        "Reorganizing excels: [########################] 1/1\n"
    )
    assert (Path.cwd() / "unique_geography_values.txt").is_file()
    (Path.cwd() / "unique_geography_values.txt").unlink()



# ---------------------------------------------------------------------------
# sanitize_name unit tests
# ---------------------------------------------------------------------------

def test_sanitize_name_replaces_spaces_with_underscores() -> None:
    assert sanitize_name("raw cane sugar") == "raw_cane_sugar"


def test_sanitize_name_collapses_duplicate_underscores() -> None:
    assert sanitize_name("r_iia__trade__1938") == "r_iia_trade_1938"


def test_sanitize_name_strips_leading_and_trailing_underscores() -> None:
    assert sanitize_name("_wheat_") == "wheat"


def test_sanitize_name_combined_spaces_and_underscores() -> None:
    assert sanitize_name("iia  crops  1929") == "iia_crops_1929"


def test_sanitize_name_combined_duplicate_underscores() -> None:
    assert sanitize_name("iia__crops__1929") == "iia_crops_1929"


def test_sanitize_name_no_change_when_already_clean() -> None:
    assert sanitize_name("r_iia_trade_1938_466_475_rice") == "r_iia_trade_1938_466_475_rice"


def test_canonical_document_name_has_no_duplicate_underscores() -> None:
    # Folder name "my trade" has a space → yearbook becomes "my_trade" (no double underscores)
    path = Path("inputs/my trade/extracted_pages_1938_39/reviewed_1_2_wheat.xlsx")
    name = canonical_document_name(path)
    assert " " not in name
    assert "__" not in name
    assert name == "r_iia_my_trade_1938_1_2_wheat"


def test_compute_output_subdir_sanitizes_space_in_folder_name() -> None:
    # Intermediate folder "my crops" has a space → must become "iia_my_crops_1929"
    path = Path("inputs/reviewed_iia/extracted_pages_1929_30/my crops/reviewed_1_2.xlsx")
    result = _compute_output_subdir(path)
    assert " " not in str(result)
    assert "__" not in str(result)
    assert result == Path("iia_extracted_pages_1929/iia_my_crops_1929")
