from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
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


@dataclass(slots=True)
class GeographyIndex:
    countries: set[str] = field(default_factory=set)
    continents: set[str] = field(default_factory=set)
    hemispheres: set[str] = field(default_factory=set)

    def add_country(self, value: str) -> None:
        if value:
            self.countries.add(value)

    def add_continent(self, value: str) -> None:
        if value:
            self.continents.add(value)

    def add_hemisphere(self, value: str) -> None:
        if value:
            self.hemispheres.add(value)

    def write_txt(self, path: str | Path) -> Path:
        path = Path(path)
        lines = [
            "[hemispheres]",
            *sorted(self.hemispheres),
            "",
            "[continents]",
            *sorted(self.continents),
            "",
            "[countries]",
            *sorted(self.countries),
            "",
        ]
        path.write_text("\n".join(lines), encoding="utf-8")
        return path



def transform_workbook(
    input_path: str | Path,
    output_path: str | Path,
    config: WorkbookConfig | None = None,
    geography_index: GeographyIndex | None = None,
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
                geography_index=geography_index,
            )
        )

    if not target_sheets:
        raise TransformationError(f"No transformable sheets found in workbook: {input_path.name}")

    return write_workbook(output_path, WorkbookData(sheets=target_sheets))



def _extract_year_headers(sheet: SheetData) -> list[tuple[int, str]]:
    headers: list[tuple[int, str]] = []
    for column in range(2, sheet.max_column + 1):
        value = sheet.get_cell(1, column).value
        header = _stringify_header(value)
        if header == "":
            continue
        headers.append((column, header))
    return headers



def _transform_sheet(
    source_sheet: SheetData,
    years: list[tuple[int, str]],
    unit: str,
    geography_index: GeographyIndex | None = None,
) -> SheetData:
    target = SheetData(name=source_sheet.name.lower())
    _write_headers(target, years)

    current_hemisphere = ""
    current_hemisphere_fill: str | None = None
    current_continent = ""
    current_continent_fill: str | None = None
    target_row = 2

    for row in range(2, source_sheet.max_row + 1):
        label_cell = source_sheet.get_cell(row, 1)
        label = _clean_text(label_cell.value)
        if not label:
            continue

        if _is_hemisphere_row(label):
            current_hemisphere = _strip_terminal_punctuation(label)
            current_hemisphere_fill = label_cell.fill_rgb
            if geography_index is not None:
                geography_index.add_hemisphere(current_hemisphere)
            continue

        if _is_continent_row(label):
            current_continent = _strip_terminal_punctuation(label)
            current_continent_fill = label_cell.fill_rgb
            if geography_index is not None:
                geography_index.add_continent(current_continent)
            continue

        country, footnotes = _extract_country_and_footnotes(label)
        if geography_index is not None:
            geography_index.add_country(country)
        target.set_cell(target_row, 1, current_hemisphere, fill_rgb=current_hemisphere_fill)
        target.set_cell(target_row, 2, current_continent, fill_rgb=current_continent_fill)
        target.set_cell(target_row, 3, country, fill_rgb=label_cell.fill_rgb)
        target.set_cell(target_row, 4, unit)
        target.set_cell(target_row, 5, footnotes)

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


def _stringify_header(value: object) -> str:
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()



def _strip_terminal_punctuation(value: str) -> str:
    return value.strip().rstrip(".:")



def _normalize_label(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_only = normalized.encode("ascii", "ignore").decode("ascii")
    return ascii_only.upper()



def _is_hemisphere_row(label: str) -> bool:
    return bool(HEMISPHERE_RE.search(label))



def _is_continent_row(label: str) -> bool:
    normalized = _normalize_label(_strip_terminal_punctuation(label))
    return normalized in KNOWN_CONTINENTS



def _extract_country_and_footnotes(label: str) -> tuple[str, str]:
    notes = [_normalize_footnote(match.strip()) for match in PAREN_RE.findall(label) if match.strip()]
    country = PAREN_RE.sub("", label)
    country = re.sub(r"\s+", " ", country).strip().rstrip("-;,")
    return country, "; ".join(notes)


def _normalize_footnote(note: str) -> str:
    if note.lower() == "r":
        return "reexports"
    return note
