"""Configuration models and lightweight YAML parsing helpers."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import TypeAlias

from .io.xlsx import read_workbook
from .utils.naming import canonical_document_name, extract_source_product
from .utils.text import normalize_text

ScalarValue: TypeAlias = str | int
ConfigSection: TypeAlias = dict[str, object]
DocumentVariableUnitMap: TypeAlias = dict[tuple[str, str], str]


@dataclass(slots=True)
class WorkbookConfig:
    """Runtime configuration for workbook normalization."""

    unit_mode: str = "standard"
    include_sheets: list[str] | None = None
    document_categories: dict[str, int] = field(default_factory=dict)
    product_aliases: dict[str, str] = field(default_factory=dict)
    product_translations: dict[str, str] = field(default_factory=dict)
    unit_overrides: dict[str, str] = field(default_factory=dict)
    document_variable_units: DocumentVariableUnitMap = field(default_factory=dict)

    def should_include_sheet(self, sheet_name: str) -> bool:
        """Return ``True`` when *sheet_name* passes the inclusion filter."""
        if self.include_sheets is None:
            return True
        allowed_sheets = {name.upper() for name in self.include_sheets}
        return sheet_name.upper() in allowed_sheets

    def canonical_name_for_document(self, document_name: str | Path) -> str:
        """Return the canonical output filename stem for *document_name*."""
        return canonical_document_name(
            document_name,
            self.product_translations,
            self.product_aliases,
        )

    def category_for_document(self, document_name: str | Path) -> int | None:
        """Return the configured numeric size category for *document_name*."""
        stem = Path(document_name).stem
        canonical_name = self.canonical_name_for_document(document_name)
        raw_value = self.document_categories.get(stem)
        if raw_value is None:
            raw_value = self.document_categories.get(canonical_name)
        return int(raw_value) if raw_value is not None else None

    def product_for_document(self, document_name: str | Path) -> str:
        """Return the normalized product name for *document_name*."""
        derived_product = extract_source_product(document_name)
        return self.product_aliases.get(derived_product, derived_product)

    def override_for(self, document_name: str | Path, sheet_name: str) -> str:
        """Return any explicit unit override for *document_name* / *sheet_name*."""
        stem = Path(document_name).stem
        canonical_name = self.canonical_name_for_document(document_name)
        candidate_keys = (
            f"{stem}:{sheet_name.lower()}",
            f"{canonical_name}:{sheet_name.lower()}",
        )
        for key in candidate_keys:
            if key in self.unit_overrides:
                return self.unit_overrides[key]
        return self.unit_overrides.get(sheet_name.lower(), "")

    def mapped_unit_for(self, document_name: str | Path, sheet_name: str) -> str:
        """Return mapped unit from the document-variable Excel mapping, if present."""
        normalized_sheet = normalize_text(sheet_name)
        document_path = Path(document_name)
        candidates = [
            normalize_text(document_path.stem),
            normalize_text(document_path.name),
        ]
        for document in candidates:
            mapped_value = self.document_variable_units.get((document, normalized_sheet))
            if mapped_value:
                return mapped_value
        return ""


def _coerce_scalar(value: str) -> ScalarValue:
    """Coerce a raw YAML scalar string to an ``int`` when appropriate."""
    stripped_value = value.strip()
    if stripped_value.isdigit() or (
        stripped_value.startswith("-") and stripped_value[1:].isdigit()
    ):
        return int(stripped_value)
    if (
        len(stripped_value) >= 2
        and stripped_value[0] == stripped_value[-1]
        and stripped_value[0] in {'"', "'"}
    ):
        return stripped_value[1:-1]
    return stripped_value


def _parse_simple_yaml(text: str) -> ConfigSection:
    """Parse the subset of YAML used by the project's configuration files."""
    result: ConfigSection = {}
    current_section: str | None = None

    for raw_line in text.splitlines():
        line = raw_line.rstrip()
        if not line or line.lstrip().startswith("#"):
            continue

        indent = len(line) - len(line.lstrip(" "))
        stripped = line.strip()

        if indent == 0:
            current_section = None
            if ":" not in stripped:
                raise ValueError(f"Invalid configuration line: {raw_line}")
            key, value = stripped.split(":", 1)
            normalized_key = key.strip()
            normalized_value = value.strip()
            if normalized_value:
                result[normalized_key] = _coerce_scalar(normalized_value)
            else:
                current_section = normalized_key
                result[current_section] = {}
            continue

        if current_section is None:
            raise ValueError(
                f"Unexpected indentation in configuration line: {raw_line}"
            )

        container = result[current_section]
        if stripped.startswith("- "):
            if not isinstance(container, list):
                container = []
                result[current_section] = container
            container.append(_coerce_scalar(stripped[2:]))
            continue

        if ":" not in stripped:
            raise ValueError(f"Invalid nested configuration line: {raw_line}")
        key, value = stripped.split(":", 1)
        if not isinstance(container, dict):
            raise ValueError(
                f"Section {current_section} must be a mapping to contain key/value pairs"
            )
        container[key.strip()] = _coerce_scalar(value.strip())

    return result


def _normalize_alias_map(raw_mapping: dict[str, object] | None) -> dict[str, str]:
    """Normalize both keys and values of *raw_mapping*."""
    mapping = raw_mapping or {}
    return {
        normalize_text(str(key)): normalize_text(str(value))
        for key, value in mapping.items()
    }


def _validate_mapping(raw_config: ConfigSection, key: str) -> dict[str, object]:
    """Return a validated mapping section from *raw_config*."""
    value = raw_config.get(key)
    if value is None:
        return {}
    if not isinstance(value, dict):
        raise ValueError(f"{key} must be expressed as a YAML mapping.")
    return value


def _load_document_variable_units(mapping_path: str | Path) -> DocumentVariableUnitMap:
    """Load ``document``/``variable``/``unit`` rows from an Excel mapping file."""
    path = Path(mapping_path)
    if not path.exists():
        return {}

    workbook = read_workbook(path)
    if not workbook.sheets:
        return {}

    sheet = workbook.sheets[0]
    header_columns: dict[str, int] = {}
    for column in range(1, sheet.max_column + 1):
        header = normalize_text(sheet.get_cell(1, column).value)
        if header in {"document", "variable", "unit"}:
            header_columns[header] = column

    required_headers = {"document", "variable", "unit"}
    if not required_headers.issubset(header_columns):
        return {}

    mapping: DocumentVariableUnitMap = {}
    for row in range(2, sheet.max_row + 1):
        raw_document = sheet.get_cell(row, header_columns["document"]).value
        raw_variable = sheet.get_cell(row, header_columns["variable"]).value
        raw_unit = sheet.get_cell(row, header_columns["unit"]).value
        if raw_document is None or raw_variable is None or raw_unit is None:
            continue
        normalized_document = normalize_text(raw_document)
        normalized_variable = normalize_text(raw_variable)
        unit = str(raw_unit).strip()
        if not normalized_document or not normalized_variable or not unit:
            continue
        mapping[(normalized_document, normalized_variable)] = unit
        document_stem = normalize_text(Path(str(raw_document)).stem)
        if document_stem:
            mapping[(document_stem, normalized_variable)] = unit
    return mapping


def load_config(config_path: str | Path | None) -> WorkbookConfig:
    """Read and validate a YAML configuration file."""
    project_root = Path(__file__).resolve().parents[3]
    mapping_path = project_root / "data" / "document_variable_unit_mapping.xlsx"

    if config_path is None:
        return WorkbookConfig(
            document_variable_units=_load_document_variable_units(mapping_path)
        )

    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"Configuration file not found: {path}")

    raw_config = _parse_simple_yaml(path.read_text(encoding="utf-8"))
    include_sheets = raw_config.get("include_sheets")
    if include_sheets is not None and not isinstance(include_sheets, list):
        raise ValueError("include_sheets must be expressed as a YAML list.")

    document_categories = _validate_mapping(raw_config, "document_categories")
    product_aliases = _validate_mapping(raw_config, "product_aliases")
    product_translations = _validate_mapping(raw_config, "product_translations")
    unit_overrides = _validate_mapping(raw_config, "unit_overrides")

    return WorkbookConfig(
        unit_mode=str(raw_config.get("unit_mode", "standard") or "standard"),
        include_sheets=[str(value) for value in include_sheets] if include_sheets else None,
        document_categories={
            str(key): int(value) for key, value in document_categories.items()
        },
        product_aliases=_normalize_alias_map(product_aliases),
        product_translations={
            normalize_text(str(key)): normalize_text(str(value))
            for key, value in product_translations.items()
        },
        unit_overrides={
            normalize_text(str(key)): str(value)
            for key, value in unit_overrides.items()
        },
        document_variable_units=_load_document_variable_units(mapping_path),
    )
