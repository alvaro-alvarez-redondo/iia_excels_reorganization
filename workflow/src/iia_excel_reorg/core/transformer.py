"""Workbook transformation logic for the normalization pipeline."""

from __future__ import annotations

import itertools
import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path
from typing import TypeAlias

from ..config import WorkbookConfig
from ..io.xlsx import SheetData, WorkbookData, read_workbook, write_workbook
from ..services.units import assign_unit

HeaderYear: TypeAlias = tuple[int, str]
RowValue: TypeAlias = str | int | float | None


@dataclass(slots=True)
class OutputRow:
    """Row-oriented normalized output used before worksheet materialization."""

    values: list[RowValue]
    fills: list[str | None]


HEADER_FILL = "FF3CCB5A"
HEADER_COLUMNS = ["hemisphere", "continent", "country", "unit", "footnotes"]
PAREN_RE = re.compile(r"\(([^()]*)\)")
HEMISPHERE_RE = re.compile(r"h[eéê]misph[eèê]?re|hemisphere", re.IGNORECASE)


def _normalize_known_geography_label(value: str) -> str:
    """Strip accents, fold to ASCII lowercase, and strip trailing punctuation."""
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
KNOWN_CONTINENTS = {
    _normalize_known_geography_label(label) for label in RAW_CONTINENT_LABELS
}

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
KNOWN_HEMISPHERES = {
    _normalize_known_geography_label(label) for label in RAW_HEMISPHERE_LABELS
}


class TransformationError(RuntimeError):
    """Raised when a source worksheet cannot be transformed."""


@dataclass(slots=True)
class GeographyIndex:
    """Accumulate unique hemisphere, continent, and country labels."""

    countries: set[str] = field(default_factory=set)
    continents: set[str] = field(default_factory=set)
    hemispheres: set[str] = field(default_factory=set)

    def add_country(self, value: str) -> None:
        """Add *value* to the countries set when non-empty."""
        if value:
            self.countries.add(value)

    def add_continent(self, value: str) -> None:
        """Add *value* to the continents set when non-empty."""
        if value:
            self.continents.add(value)

    def add_hemisphere(self, value: str) -> None:
        """Add *value* to the hemispheres set when non-empty."""
        if value:
            self.hemispheres.add(value)

    def write_txt(self, path: str | Path) -> Path:
        """Write sorted geography labels to *path* in an INI-like text format."""
        output_path = Path(path)
        sections = {
            "hemispheres": sorted(self.hemispheres),
            "continents": sorted(self.continents),
            "countries": sorted(self.countries),
        }
        output_lines: list[str] = []
        for label, values in sections.items():
            output_lines.extend([f"[{label}]", *values, ""])
        output_path.write_text("\n".join(output_lines), encoding="utf-8")
        return output_path


@dataclass(slots=True)
class ProductIndex:
    """Accumulate unique product labels seen across transformed workbooks."""

    products: set[str] = field(default_factory=set)

    def add_product(self, value: str) -> None:
        """Add *value* to the products set when non-empty."""
        if value:
            self.products.add(value)

    def write_txt(self, path: str | Path) -> Path:
        """Write sorted product labels to *path* in an INI-like text format."""
        output_path = Path(path)
        output_path.write_text(
            "\n".join(["[products]", *sorted(self.products), ""]),
            encoding="utf-8",
        )
        return output_path


@dataclass(slots=True)
class UnitFootnoteDocumentIndex:
    """Track transformed document names whose footnotes reference units."""

    documents: set[str] = field(default_factory=set)

    def add_document(self, value: str) -> None:
        """Add *value* when a transformed document contains unit-related footnotes."""
        if value:
            self.documents.add(value)

    def write_txt(self, path: str | Path) -> Path:
        """Write transformed document names with unit-related footnotes to *path*."""
        output_path = Path(path)
        output_path.write_text(
            "\n".join(["[documents]", *sorted(self.documents), ""]),
            encoding="utf-8",
        )
        return output_path


def transform_workbook(
    input_path: str | Path,
    output_path: str | Path,
    config: WorkbookConfig | None = None,
    geography_index: GeographyIndex | None = None,
    unit_footnote_document_index: UnitFootnoteDocumentIndex | None = None,
) -> Path:
    """Read *input_path*, transform each eligible sheet, and write to *output_path*."""
    workbook_config = config or WorkbookConfig()
    source_path = Path(input_path)
    source_workbook = read_workbook(source_path)
    target_sheets: list[SheetData] = []
    product = workbook_config.product_for_document(source_path)
    category = workbook_config.category_for_document(source_path)

    has_unit_related_footnotes = False

    for source_sheet in source_workbook.sheets:
        if not workbook_config.should_include_sheet(source_sheet.name):
            continue

        years = _extract_year_headers(source_sheet)
        if not years:
            continue

        unit = workbook_config.override_for(
            source_path, source_sheet.name
        ) or assign_unit(
            variable=source_sheet.name,
            product=product,
            category=category,
            mode=workbook_config.unit_mode,
        )
        transformed_sheet, sheet_has_unit_related_footnotes = _transform_sheet(
            source_sheet=source_sheet,
            years=years,
            unit=unit,
            geography_index=geography_index,
        )
        target_sheets.append(transformed_sheet)
        has_unit_related_footnotes = (
            has_unit_related_footnotes or sheet_has_unit_related_footnotes
        )

    if not target_sheets:
        raise TransformationError(
            f"No transformable sheets found in workbook: {source_path.name}"
        )

    written_output_path = write_workbook(
        output_path, WorkbookData(sheets=target_sheets)
    )
    if has_unit_related_footnotes and unit_footnote_document_index is not None:
        unit_footnote_document_index.add_document(written_output_path.name)
    return written_output_path


def _extract_year_headers(sheet: SheetData) -> list[HeaderYear]:
    """Return ``(column_index, label)`` pairs for populated header cells."""
    return [
        (column, header)
        for column in range(2, sheet.max_column + 1)
        if (header := _stringify_header(sheet.get_cell(1, column).value))
    ]


def _transform_sheet(
    source_sheet: SheetData,
    years: list[HeaderYear],
    unit: str,
    geography_index: GeographyIndex | None = None,
) -> tuple[SheetData, bool]:
    """Convert one source sheet into the standardized long-format layout."""
    target_sheet = SheetData(name=source_sheet.name.lower())
    _write_headers(target_sheet, years)

    output_rows, has_unit_related_footnotes = _build_output_rows(
        source_sheet,
        years,
        unit,
        geography_index,
    )
    target_row = 2
    for output_row in output_rows:
        if output_row is None:
            target_row += 1
            continue
        target_sheet.set_row(target_row, output_row.values, output_row.fills)
        target_row += 1

    return target_sheet, has_unit_related_footnotes


def _write_headers(target: SheetData, years: list[HeaderYear]) -> None:
    """Write the fixed header row plus one column per year label."""
    header_values = list(itertools.chain(HEADER_COLUMNS, (label for _, label in years)))
    target.set_row(1, header_values, [HEADER_FILL] * len(header_values))


def _build_output_rows(
    source_sheet: SheetData,
    years: list[HeaderYear],
    unit: str,
    geography_index: GeographyIndex | None,
) -> tuple[list[OutputRow | None], bool]:
    """Build normalized output rows before materializing the target worksheet."""
    output_rows: list[OutputRow | None] = []
    has_unit_related_footnotes = False
    current_hemisphere = ""
    current_hemisphere_fill: str | None = None
    current_continent = ""
    current_continent_fill: str | None = None

    for source_row in range(2, source_sheet.max_row + 1):
        label_cell = source_sheet.get_cell(source_row, 1)
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
            if current_continent and output_rows:
                output_rows.append(None)
            current_continent = _strip_terminal_punctuation(label)
            current_continent_fill = label_cell.fill_rgb
            if geography_index is not None:
                geography_index.add_continent(current_continent)
            continue

        country, footnotes = _extract_country_and_footnotes(label)
        has_unit_related_footnotes = (
            has_unit_related_footnotes or _has_unit_related_footnote(footnotes)
        )
        if geography_index is not None:
            geography_index.add_country(country)
        output_rows.append(
            _build_output_row(
                source_sheet=source_sheet,
                source_row=source_row,
                years=years,
                hemisphere=current_hemisphere,
                hemisphere_fill=current_hemisphere_fill,
                continent=current_continent,
                continent_fill=current_continent_fill,
                country=country,
                country_fill=label_cell.fill_rgb,
                unit=unit,
                footnotes=footnotes,
            )
        )

    return output_rows, has_unit_related_footnotes


def _build_output_row(
    *,
    source_sheet: SheetData,
    source_row: int,
    years: list[HeaderYear],
    hemisphere: str,
    hemisphere_fill: str | None,
    continent: str,
    continent_fill: str | None,
    country: str,
    country_fill: str | None,
    unit: str,
    footnotes: str,
) -> OutputRow:
    """Return one normalized output row for a source data row."""
    values: list[RowValue] = [hemisphere, continent, country, unit, footnotes]
    fills: list[str | None] = [
        hemisphere_fill,
        continent_fill,
        country_fill,
        None,
        None,
    ]
    for source_column, _ in years:
        source_value = source_sheet.get_cell(source_row, source_column)
        values.append(source_value.value)
        fills.append(source_value.fill_rgb)
    return OutputRow(values=values, fills=fills)


def _clean_text(value: str | int | float | None) -> str:
    """Return *value* as a stripped string, or ``""`` for null-like values."""
    return str(value).strip() if value is not None else ""


def _strip_terminal_punctuation(value: str) -> str:
    """Strip terminal periods and colons from *value*."""
    return value.rstrip().rstrip(".:")


_UNIT_FOOTNOTE_RE = re.compile(
    r"\b(?:unit|units|tonne|tonnes|kg|kilogram|kilograms|q|quintal|quintals|"
    r"ha|hectare|hectares|hl|hectoliter|hectoliters|head|heads|egg|eggs|"
    r"hg)\b",
    re.IGNORECASE,
)


def _has_unit_related_footnote(value: str) -> bool:
    """Return whether *value* mentions a measurement unit or unit hint."""
    return bool(_UNIT_FOOTNOTE_RE.search(value))


def _normalize_footnote(value: str) -> str:
    """Normalize extracted footnote text for output."""
    return re.sub(r"\s+", " ", value.strip(" .;,")).strip()


def _extract_country_and_footnotes(label: str) -> tuple[str, str]:
    """Split a row label into country name and semicolon-separated footnotes."""
    footnotes = [_normalize_footnote(match) for match in PAREN_RE.findall(label)]
    country = PAREN_RE.sub("", label).strip()
    country = _strip_terminal_punctuation(country)
    normalized_notes = [note for note in footnotes if note]
    if not normalized_notes and label.endswith("(r)"):
        normalized_notes = ["reexports"]
    elif any(note == "r" for note in normalized_notes):
        normalized_notes = [
            "reexports" if note == "r" else note for note in normalized_notes
        ]
    return country, "; ".join(normalized_notes)


def _stringify_header(value: str | int | float | None) -> str:
    """Return a normalized year/header label string."""
    if value is None:
        return ""
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _is_continent_row(value: str) -> bool:
    """Return whether *value* matches a known continent label."""
    return _normalize_known_geography_label(value) in KNOWN_CONTINENTS


def _is_hemisphere_row(value: str) -> bool:
    """Return whether *value* matches a known hemisphere label."""
    normalized_value = _normalize_known_geography_label(value)
    return normalized_value in KNOWN_HEMISPHERES or bool(HEMISPHERE_RE.search(value))
