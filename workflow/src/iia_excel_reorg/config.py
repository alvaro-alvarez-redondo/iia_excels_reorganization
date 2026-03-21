"""Configuration models and lightweight YAML parsing helpers."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import TypeAlias

from .io.xlsx import SheetData, WorkbookData, read_workbook, write_workbook
from .utils.naming import canonical_document_name, extract_source_product
from .utils.text import normalize_text

ScalarValue: TypeAlias = str | int
ConfigSection: TypeAlias = dict[str, object]


@dataclass(slots=True)
class WorkbookConfig:
    """Runtime configuration for workbook normalization."""

    unit_mode: str = "standard"
    include_sheets: list[str] | None = None
    document_categories: dict[str, int] = field(default_factory=dict)
    product_aliases: dict[str, str] = field(default_factory=dict)
    product_translations: dict[str, str] = field(default_factory=dict)
    unit_overrides: dict[str, str] = field(default_factory=dict)
    document_variable_units: dict[tuple[str, str], str] = field(default_factory=dict)

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

    def mapped_unit_for(self, document_name: str | Path, variable: str) -> str:
        """Return mapped unit for source *document_name* and *variable* when set."""
        normalized_variable = normalize_text(variable)
        document_stem = Path(document_name).stem
        normalized_document = normalize_text(document_stem)
        return self.document_variable_units.get(
            (normalized_document, normalized_variable),
            "",
        )


def _load_document_variable_units(mapping_path: Path) -> dict[tuple[str, str], str]:
    """Load ``(document, variable) -> unit`` mappings from an Excel workbook.

    Vectorized implementation: replace the row-by-row ``for row in range(2, …)``
    loop with a single dict-comprehension that iterates over a generator of
    pre-extracted ``(document, variable, unit)`` tuples, reducing Python
    interpreter overhead from O(n) loop iterations to one comprehension pass.
    """
    if not mapping_path.exists():
        mapping_path.parent.mkdir(parents=True, exist_ok=True)
        template = SheetData(name="document_variable_unit_mapping")
        template.set_cell(1, 1, "document")
        template.set_cell(1, 2, "variable")
        template.set_cell(1, 3, "unit")
        write_workbook(mapping_path, WorkbookData(sheets=[template]))
        return {}

    workbook = read_workbook(mapping_path)
    if not workbook.sheets:
        return {}
    mapping_sheet = workbook.sheets[0]

    header_positions: dict[str, int] = {}
    for column in range(1, mapping_sheet.max_column + 1):
        raw_header = mapping_sheet.get_cell(1, column).value
        header = normalize_text(str(raw_header)) if raw_header is not None else ""
        if header in {"document", "variable", "unit"}:
            header_positions[header] = column

    required_headers = {"document", "variable", "unit"}
    if not required_headers.issubset(header_positions):
        return {}

    doc_col = header_positions["document"]
    var_col = header_positions["variable"]
    unit_col = header_positions["unit"]

    def _extract_row(row: int) -> tuple[tuple[str, str], str] | None:
        """Return a ``((document, variable), unit)`` pair or ``None`` to skip."""
        raw_doc = mapping_sheet.get_cell(row, doc_col).value
        raw_var = mapping_sheet.get_cell(row, var_col).value
        raw_unit = mapping_sheet.get_cell(row, unit_col).value
        if raw_doc is None or raw_var is None or raw_unit is None:
            return None
        document = normalize_text(Path(str(raw_doc)).stem)
        variable = normalize_text(str(raw_var))
        unit = str(raw_unit).strip()
        if not document or not variable or not unit:
            return None
        return (document, variable), unit

    return dict(
        filter(
            None,
            (_extract_row(row) for row in range(2, mapping_sheet.max_row + 1)),
        )
    )


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


def load_config(config_path: str | Path | None) -> WorkbookConfig:
    """Read and validate a YAML configuration file."""
    project_root = Path(__file__).resolve().parents[3]
    default_mapping_path = project_root / "data" / "document_variable_unit_mapping.xlsx"

    if config_path is None:
        return WorkbookConfig(
            document_variable_units=_load_document_variable_units(default_mapping_path),
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
    mapping_file = raw_config.get("document_variable_unit_mapping_file")
    mapping_path = (
        Path(str(mapping_file))
        if mapping_file is not None
        else default_mapping_path
    )

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
