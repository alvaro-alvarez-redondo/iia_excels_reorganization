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
from ..services.units import UNIT_PLACEHOLDER

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

RAW_WORLD_TOTAL_COUNTRY_LABELS = (
    "totaux generattx",
    "totaux generaux",
    "totaux generaux des imp et des exp",
    "totaux generaux des impet des exp",
    "totaux generaux non compris i'urss",
    "totaux generaux non compris l'urss",
    "totaux non compris l'u r s s generaux",
    "generaux y compris l'u r s s",
    "generaux y compris l'u r ss",
    "totaux generaux compris i'urss",
    "totaux generaux compris l'urss",
    "y compris l'u r ss",
    "totaux generaux des imp et des exp nettes",
    "totaux generaux des impet des expnettes",
    "total general net excluding the ussr",
)


def _normalize_country_match_label(value: str) -> str:
    """Return a normalized country label for special-case matching."""
    normalized = _normalize_known_geography_label(value)
    return re.sub(r"[^a-z0-9]+", "", normalized)


WORLD_TOTAL_COUNTRY_LABELS = {
    _normalize_country_match_label(label) for label in RAW_WORLD_TOTAL_COUNTRY_LABELS
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
        """Write all geography labels to *path* in the legacy combined format."""
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

    def write_dimension_txt(self, path: str | Path, *, label: str) -> Path:
        """Write one geography dimension to *path* in a deduplicated TXT format."""
        output_path = Path(path)
        values_by_label = {
            "hemispheres": self.hemispheres,
            "continents": self.continents,
            "countries": self.countries,
        }
        values = values_by_label[label]
        output_path.write_text(
            "\n".join([f"[{label}]", *sorted(values), ""]),
            encoding="utf-8",
        )
        return output_path

    def write_split_txts(self, directory: str | Path) -> list[Path]:
        """Write separate deduplicated TXT files for each geography dimension."""
        output_dir = Path(directory)
        output_dir.mkdir(parents=True, exist_ok=True)
        return [
            self.write_dimension_txt(
                output_dir / "unique_hemisphere_values.txt",
                label="hemispheres",
            ),
            self.write_dimension_txt(
                output_dir / "unique_continent_values.txt",
                label="continents",
            ),
            self.write_dimension_txt(
                output_dir / "unique_country_values.txt",
                label="countries",
            ),
        ]


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
class DocumentIndex:
    """Track transformed document names."""

    documents: set[str] = field(default_factory=set)

    def add_document(self, value: str) -> None:
        """Add *value* when non-empty."""
        if value:
            self.documents.add(value)

    def write_txt(self, path: str | Path) -> Path:
        """Write sorted transformed document names to *path*."""
        output_path = Path(path)
        output_path.write_text(
            "\n".join(["[documents]", *sorted(self.documents), ""]),
            encoding="utf-8",
        )
        return output_path


@dataclass(slots=True)
class UnitFootnoteDocumentIndex(DocumentIndex):
    """Track transformed document names whose footnotes reference units."""


@dataclass(slots=True)
class MissingUnitCountryDocumentIndex(DocumentIndex):
    """Track transformed documents that contain at least one country with no unit."""


@dataclass(slots=True)
class FootnoteIndex:
    """Accumulate unique normalized footnotes."""

    footnotes: set[str] = field(default_factory=set)

    def add_footnotes(self, values: list[str]) -> None:
        """Add all non-empty footnotes from *values*."""
        self.footnotes.update(value for value in values if value)

    def write_txt(self, path: str | Path) -> Path:
        """Write sorted footnotes to *path*."""
        output_path = Path(path)
        output_path.write_text(
            "\n".join(["[footnotes]", *sorted(self.footnotes), ""]),
            encoding="utf-8",
        )
        return output_path


def transform_workbook(
    input_path: str | Path,
    output_path: str | Path,
    config: WorkbookConfig | None = None,
    geography_index: GeographyIndex | None = None,
    unit_footnote_document_index: UnitFootnoteDocumentIndex | None = None,
    missing_unit_country_document_index: MissingUnitCountryDocumentIndex | None = None,
) -> Path:
    """Read *input_path*, transform each eligible sheet, and write to *output_path*."""
    workbook_config = config or WorkbookConfig()
    source_path = Path(input_path)
    source_workbook = read_workbook(source_path)
    target_sheets: list[SheetData] = []
    has_unit_related_footnotes = False
    has_countries_with_missing_units = False

    for source_sheet in source_workbook.sheets:
        if not workbook_config.should_include_sheet(source_sheet.name):
            continue

        years = _extract_year_headers(source_sheet)
        if not years:
            continue

        mapped_unit = workbook_config.mapped_unit_for(source_path, source_sheet.name)
        document_unit = (
            mapped_unit
            or workbook_config.override_for(source_path, source_sheet.name)
            or UNIT_PLACEHOLDER
        )

        (
            transformed_sheet,
            sheet_has_unit_related_footnotes,
            sheet_has_countries_with_missing_units,
        ) = _transform_sheet(
            source_sheet=source_sheet,
            years=years,
            unit=document_unit,
            geography_index=geography_index,
        )
        target_sheets.append(transformed_sheet)
        has_unit_related_footnotes = (
            has_unit_related_footnotes or sheet_has_unit_related_footnotes
        )
        has_countries_with_missing_units = (
            has_countries_with_missing_units or sheet_has_countries_with_missing_units
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
    if (
        has_countries_with_missing_units
        and missing_unit_country_document_index is not None
    ):
        missing_unit_country_document_index.add_document(written_output_path.name)
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
) -> tuple[SheetData, bool, bool]:
    """Convert one source sheet into the standardized long-format layout."""
    target_sheet = SheetData(name=source_sheet.name.lower())
    _write_headers(target_sheet, years)

    (
        output_rows,
        has_unit_related_footnotes,
        has_countries_with_missing_units,
    ) = _build_output_rows(
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

    return target_sheet, has_unit_related_footnotes, has_countries_with_missing_units


def _write_headers(target: SheetData, years: list[HeaderYear]) -> None:
    """Write the fixed header row plus one column per year label."""
    header_values = list(itertools.chain(HEADER_COLUMNS, (label for _, label in years)))
    target.set_row(1, header_values, [HEADER_FILL] * len(header_values))


def _build_output_rows(
    source_sheet: SheetData,
    years: list[HeaderYear],
    unit: str,
    geography_index: GeographyIndex | None,
    footnote_index: FootnoteIndex | None = None,
) -> tuple[list[OutputRow | None], bool, bool]:
    """Build normalized output rows before materializing the target worksheet."""
    output_rows: list[OutputRow | None] = []
    has_unit_related_footnotes = False
    has_countries_with_missing_units = False
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
        output_continent = current_continent
        output_continent_fill = current_continent_fill
        if _normalize_country_match_label(country) in WORLD_TOTAL_COUNTRY_LABELS:
            output_continent = "WORLD"
            output_continent_fill = None
            if geography_index is not None:
                geography_index.add_continent(output_continent)
        footnote_values = _extract_footnotes(label)
        has_unit_related_footnotes = (
            has_unit_related_footnotes or _has_unit_related_footnote(footnotes)
        )
        if geography_index is not None:
            geography_index.add_country(country)
        if _is_missing_unit(unit):
            has_countries_with_missing_units = True
        if footnote_index is not None:
            footnote_index.add_footnotes(footnote_values)
        output_rows.append(
            _build_output_row(
                source_sheet=source_sheet,
                source_row=source_row,
                years=years,
                hemisphere=current_hemisphere,
                hemisphere_fill=current_hemisphere_fill,
                continent=output_continent,
                continent_fill=output_continent_fill,
                country=country,
                country_fill=label_cell.fill_rgb,
                unit=unit,
                footnotes=footnotes,
            )
        )

    return output_rows, has_unit_related_footnotes, has_countries_with_missing_units


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
    normalized_unit = "" if _is_missing_unit(unit) else unit
    values: list[RowValue] = [hemisphere, continent, country, normalized_unit, footnotes]
    fills: list[str | None] = [
        hemisphere_fill,
        continent_fill,
        country_fill,
        None,
        None,
    ]
    for source_column, _ in years:
        source_value = source_sheet.get_cell(source_row, source_column)
        values.append(_normalize_year_value(source_value.value))
        fills.append(source_value.fill_rgb)
    return OutputRow(values=values, fills=fills)


def _normalize_year_value(value: RowValue) -> RowValue:
    """Normalize OCR-confused characters in year-column values."""
    if not isinstance(value, str):
        return value
    normalized = value.translate(str.maketrans({"i": "1", "I": "1", "o": "0", "O": "0"}))
    cleaned = re.sub(r"[^\d.]", "", normalized)
    if cleaned.count(".") <= 1:
        return cleaned

    integer_part, decimal_part = cleaned.rsplit(".", 1)
    integer_part = integer_part.replace(".", "")
    return f"{integer_part}.{decimal_part}" if decimal_part else integer_part


def _clean_text(value: str | int | float | None) -> str:
    """Return *value* as a stripped string, or ``""`` for null-like values."""
    return str(value).strip() if value is not None else ""


def _strip_terminal_punctuation(value: str) -> str:
    """Strip terminal periods and colons from *value*."""
    return value.rstrip().rstrip(".:")


_MISSING_UNIT_SENTINELS = {
    "",
    "__na_unit__",
    "na",
    "n/a",
    "n.a.",
    "none",
    "null",
}


def _is_missing_unit(value: str) -> bool:
    """Return whether *value* should be treated as a missing/unknown unit."""
    normalized = value.strip().casefold().replace(" ", "")
    return normalized in _MISSING_UNIT_SENTINELS


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


def _extract_footnotes(label: str) -> list[str]:
    """Return normalized footnotes extracted from *label*."""
    footnotes = [_normalize_footnote(match) for match in PAREN_RE.findall(label)]
    normalized_notes = [note for note in footnotes if note]
    if not normalized_notes and label.endswith("(r)"):
        normalized_notes = ["reexports"]
    elif any(note == "r" for note in normalized_notes):
        normalized_notes = [
            "reexports" if note == "r" else note for note in normalized_notes
        ]
    return normalized_notes


def _extract_country(label: str) -> str:
    """Return the country/component label with parenthesized footnotes removed."""
    return _clean_text(PAREN_RE.sub("", label)).rstrip()


def _extract_country_and_footnotes(label: str) -> tuple[str, str]:
    """Return the normalized country label and joined footnotes for *label*."""
    country = _extract_country(label)
    return country, "; ".join(_extract_footnotes(label))


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
