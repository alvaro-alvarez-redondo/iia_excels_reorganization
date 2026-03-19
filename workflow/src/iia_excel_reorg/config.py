"""Configuration models and lightweight YAML parsing helpers."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path
from typing import TypeAlias

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
    if config_path is None:
        return WorkbookConfig()

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
    )
