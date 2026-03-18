from __future__ import annotations

import re
from pathlib import Path

from .config import WorkbookConfig
from .unit_rules import assign_unit
from .xlsx_io import SheetData, WorkbookData, read_workbook, write_workbook

HEADER_FILL = "FF3CCB5A"
HEADER_COLUMNS = ["hemisphere", "continent", "country", "unit", "footnotes"]
PAREN_RE = re.compile(r"\(([^()]*)\)")
HEMISPHERE_RE = re.compile(r"h[eé]misph[eè]re|hemisphere", re.IGNORECASE)
KNOWN_CONTINENTS = {
    "EUROPE",
    "AMERIQUE",
    "AMERICA",
    "ASIE",
    "ASIA",
    "AFRIQUE",
    "AFRICA",
    "OCEANIE",
    "OCEANIA",
}


class TransformationError(RuntimeError):
    """Raised when a source worksheet cannot be transformed."""



def transform_workbook(
    input_path: str | Path,
    output_path: str | Path,
    config: WorkbookConfig | None = None,
) -> Path:
    config = config or WorkbookConfig()
    input_path = Path(input_path)
    source_wb = read_workbook(input_path)
    target_sheets: list[SheetData] = []
    product = config.product_for_document(input_path)
    category = config.category_for_document(input_path)

    for source_sheet in source_wb.sheets:
        if not config.should_include_sheet(source_sheet.name):
            continue
        years = _extract_year_headers(source_sheet)
        if not years:
            continue
        override_unit = config.override_for(input_path, source_sheet.name)
        unit = override_unit or assign_unit(
            variable=source_sheet.name,
            product=product,
            category=category,
            mode=config.unit_mode,
        )
        target_sheets.append(
            _transform_sheet(
                source_sheet=source_sheet,
                years=years,
                unit=unit,
            )
        )

    if not target_sheets:
        raise TransformationError(f"No transformable sheets found in workbook: {input_path.name}")

    return write_workbook(output_path, WorkbookData(sheets=target_sheets))



def _extract_year_headers(sheet: SheetData) -> list[tuple[int, str]]:
    headers: list[tuple[int, str]] = []
    for column in range(2, sheet.max_column + 1):
        value = sheet.get_cell(1, column).value
        if value is None or str(value).strip() == "":
            continue
        headers.append((column, str(value).strip()))
    return headers



def _transform_sheet(source_sheet: SheetData, years: list[tuple[int, str]], unit: str) -> SheetData:
    target = SheetData(name=source_sheet.name.lower())
    _write_headers(target, years)

    current_hemisphere = ""
    current_continent = ""
    target_row = 2

    for row in range(2, source_sheet.max_row + 1):
        label_cell = source_sheet.get_cell(row, 1)
        label = _clean_text(label_cell.value)
        if not label:
            continue

        if _is_hemisphere_row(label):
            current_hemisphere = _strip_terminal_punctuation(label)
            continue

        if _is_continent_row(label):
            current_continent = _strip_terminal_punctuation(label)
            continue

        country, footnotes = _extract_country_and_footnotes(label)
        for column, value in enumerate(
            [current_hemisphere, current_continent, country, unit, footnotes],
            start=1,
        ):
            target.set_cell(target_row, column, value, fill_rgb=label_cell.fill_rgb)

        for offset, (source_column, _label) in enumerate(years, start=6):
            source_value = source_sheet.get_cell(row, source_column)
            target.set_cell(target_row, offset, source_value.value, fill_rgb=source_value.fill_rgb)

        target_row += 1

    return target



def _write_headers(target: SheetData, years: list[tuple[int, str]]) -> None:
    for column, header in enumerate(HEADER_COLUMNS + [label for _, label in years], start=1):
        target.set_cell(1, column, header, fill_rgb=HEADER_FILL)



def _clean_text(value: object) -> str:
    if value is None:
        return ""
    return str(value).strip()



def _strip_terminal_punctuation(value: str) -> str:
    return value.strip().rstrip(".:")



def _normalize_label(value: str) -> str:
    substitutions = str.maketrans({
        "É": "E",
        "È": "E",
        "Ê": "E",
        "Á": "A",
        "À": "A",
        "Í": "I",
        "Ó": "O",
        "Ú": "U",
    })
    return value.upper().translate(substitutions)



def _is_hemisphere_row(label: str) -> bool:
    return bool(HEMISPHERE_RE.search(label))



def _is_continent_row(label: str) -> bool:
    normalized = _normalize_label(_strip_terminal_punctuation(label))
    return normalized in KNOWN_CONTINENTS



def _extract_country_and_footnotes(label: str) -> tuple[str, str]:
    notes = [match.strip() for match in PAREN_RE.findall(label) if match.strip()]
    country = PAREN_RE.sub("", label)
    country = re.sub(r"\s+", " ", country).strip().rstrip("-;,")
    return country, "; ".join(notes)
