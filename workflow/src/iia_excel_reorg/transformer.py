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
HEMISPHERE_RE = re.compile(r"h[eéê]misph[eèê]?re|hemisphere", re.IGNORECASE)


def _normalize_known_geography_label(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    ascii_only = normalized.encode("ascii", "ignore").decode("ascii")
    return ascii_only.casefold().strip().rstrip(".:")


RAW_CONTINENT_LABELS = (
    "AFRIQUE",
    "AMÉR DU NORD ET AMÉR CENTRALE",
    "AMER. DU NORD ET AMER. CENTR.",
    "AMERIQUE",
    "AMERIQUÉ",
    "AMÉRIQUE",
    "AMÉRIQUE CENTRALE ET MEXIQUE.",
    "AMERIQUE CENTRALE.",
    "AMÉRIQUE CENTRALE.",
    "AMERIQUE DU NORD",
    "AMÉRIQUE DU NORD",
    "AMÉRIQUE DU NORD ET AMÉR CENTR",
    "AMERIQUE DU NORD ET AMERIQUE CENTRALE",
    "AMERIQUE DU NORD ET AMÉRIQUE CENTRALE",
    "AMÉRIQUE DU NORD ET AMERIQUE CENTRALE",
    "AMÉRIQUE DU NORD ET AMÉRIQUE CENTRALE",
    "AMERIQUE DU NORD ET AMERIQUE CENTRALE.",
    "AMERIQUE DU NORD ET AMÉRIQUE CENTRALE.",
    "AMÉRIQUE DU NORD ET AMERIQUE CENTRALE.",
    "AMÉRIQUE DU NORD ET AMÉRIQUE CENTRALE.",
    "AMÉRIQUE DU NORD ET CENTRALE.",
    "AMERIQUE DU NORD.",
    "AMÉRIQUE DU NORD.",
    "AMERIQUE DU SUD",
    "AMÉRIQUE DU SUD",
    "AMERIQUE DU SUD.",
    "AMÉRIQUE DU SUD.",
    "AMERIQUE MERIDIONALE",
    "AMERIQUE MÉRIDIONALE",
    "AMÉRIQUE MERIDIONALE",
    "Amérique méridionale",
    "AMERIQUE MERIDIONALE.",
    "AMERIQUE MÉRIDIONALE.",
    "AMÉRIQUE MERIDIONALE.",
    "AMÉRIQUE MÉRIDIONALE.",
    "AMERIQUE MÉRILIONALE",
    "AMÉRIQUE SEPIENTRIONALE ET CENTRALE",
    "AMERIQUE SEPT ET CENTRALE",
    "AMÉRIQUE SEPT. ET CENTR.",
    "AMÉRIQUE SEPT. ET CENTRALE",
    "AMÉRIQUE SEPT. ET CENTRALE.",
    "AMERIQUE SEPTENT ET CENTRALE",
    "AMÉRIQUE SEPTENT. ET CENTRALE",
    "AMERIQUE SEPTENTR ET CENTR",
    "AMÉRIQUE SEPTENTR. ET CENTP.",
    "AMÉRIQUE SEPTENTR. ET CENTR.",
    "AMERIQUE SEPTENTR. ET CENTRALE",
    "AMÉRIQUE SEPTENTR. ET CENTRALE",
    "AMÉRIQUE SEPTENTRION ET CENTRALE",
    "AMERIQUE SEPTENTRIONA LE ET CENTRALE",
    "AMÉRIQUE SEPTENTRIONALD ET CENTRALE.",
    "AMERIQUE SEPTENTRIONALE",
    "AMÉRIQUE SEPTENTRIONALE",
    "AMERIQUE SEPTENTRIONALE ET CENTR",
    "AMÉRIQUE SEPTENTRIONALE ET CENTR",
    "AMERIQUE SEPTENTRIONALE ET CENTRALE",
    "Amérique septentrionale et centrale",
    "AMERIQUE SEPTENTRIONALE ET CENTRALE.",
    "AMÉRIQUE SEPTENTRIONALE ET CENTRALE.",
    "AMÉRIQUE SEPTENTRIONALE.",
    "AMÉRIQUEMÉRIDIONALE",
    "ASIE",
    "AUSTRALIE",
    "EUROPE",
    "OCEANIE",
    "OCÉANIE",
    "OCEANIR.",
    "OCÉANTE",
    "OCEANTE.",
    "OCRANIE",
    "OCRANIE.",
)
KNOWN_CONTINENTS = {_normalize_known_geography_label(label) for label in RAW_CONTINENT_LABELS}

RAW_HEMISPHERE_LABELS = (
    "HÉMISHPÈRE SEPTENTRIONAL",
    "HÉMISPHERE MÉRIDIONAL",
    "HÉMISPHÈRE MERIDIONAL",
    "HÉMISPHÈRE MÉRIDIONAL",
    "HÉMISPHÊRE MÉRIDIONAL",
    "HEMISPHERE NORD",
    "HEMISPHÈRE NORD",
    "HÉMISPHÈRE NORD",
    "HEMISPHERE SEPTENTRIONAL",
    "HÉMISPHERE SEPTENTRIONAL",
    "HÉMISPHÈRE SEPTENTRIONAL",
    "HÊMISPHÊRE SEPTENTRIONAL",
    "HEMISPHERE SUD",
    "HEMISPHÈRE SUD",
    "HÉMISPHERE SUD",
    "HÉMISPHÈRE SUD",
    "HÉMISPHÈRE SUDAFRIQUE.",
)
KNOWN_HEMISPHERES = {_normalize_known_geography_label(label) for label in RAW_HEMISPHERE_LABELS}



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


@dataclass(slots=True)
class ProductIndex:
    products: set[str] = field(default_factory=set)

    def add_product(self, value: str) -> None:
        if value:
            self.products.add(value)

    def write_txt(self, path: str | Path) -> Path:
        path = Path(path)
        lines = [
            "[products]",
            *sorted(self.products),
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
            if current_continent and target_row > 2:
                target_row += 1
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
    return _normalize_known_geography_label(value)



def _is_hemisphere_row(label: str) -> bool:
    normalized = _normalize_label(_strip_terminal_punctuation(label))
    return normalized in KNOWN_HEMISPHERES or bool(HEMISPHERE_RE.search(label))



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
