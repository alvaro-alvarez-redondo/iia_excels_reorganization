from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from .naming import canonical_document_name, extract_source_product
from .unit_rules import normalize_text


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
        if self.include_sheets is None:
            return True
        allowed = {name.upper() for name in self.include_sheets}
        return sheet_name.upper() in allowed

    def canonical_name_for_document(self, document_name: str | Path) -> str:
        return canonical_document_name(document_name, self.product_translations, self.product_aliases)

    def category_for_document(self, document_name: str | Path) -> int | None:
        stem = Path(document_name).stem
        canonical = self.canonical_name_for_document(document_name)
        raw_value = self.document_categories.get(stem)
        if raw_value is None:
            raw_value = self.document_categories.get(canonical)
        return int(raw_value) if raw_value is not None else None

    def product_for_document(self, document_name: str | Path) -> str:
        derived = extract_source_product(document_name)
        return self.product_aliases.get(derived, derived)

    def override_for(self, document_name: str | Path, sheet_name: str) -> str:
        stem = Path(document_name).stem
        canonical = self.canonical_name_for_document(document_name)
        specific_keys = (f"{stem}:{sheet_name.lower()}", f"{canonical}:{sheet_name.lower()}")
        for key in specific_keys:
            if key in self.unit_overrides:
                return self.unit_overrides[key]
        return self.unit_overrides.get(sheet_name.lower(), "")



def _coerce_scalar(value: str) -> str | int:
    value = value.strip()
    if value.isdigit() or (value.startswith("-") and value[1:].isdigit()):
        return int(value)
    if len(value) >= 2 and value[0] == value[-1] and value[0] in {'"', "'"}:
        return value[1:-1]
    return value



def _parse_simple_yaml(text: str) -> dict[str, object]:
    result: dict[str, object] = {}
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
            key = key.strip()
            value = value.strip()
            if value == "":
                current_section = key
                result[key] = {}
            else:
                result[key] = _coerce_scalar(value)
            continue

        if current_section is None:
            raise ValueError(f"Unexpected indentation in configuration line: {raw_line}")

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
        key = key.strip()
        value = _coerce_scalar(value.strip())
        if not isinstance(container, dict):
            raise ValueError(f"Section {current_section} must be a mapping to contain key/value pairs")
        container[key] = value

    return result



def _normalize_alias_map(raw_mapping: dict[str, object] | None) -> dict[str, str]:
    mapping = raw_mapping or {}
    return {
        normalize_text(str(key)): normalize_text(str(value))
        for key, value in mapping.items()
    }



def load_config(config_path: str | Path | None) -> WorkbookConfig:
    if config_path is None:
        return WorkbookConfig()

    path = Path(config_path)
    if not path.exists():
        raise FileNotFoundError(f"Configuration file not found: {path}")

    raw = _parse_simple_yaml(path.read_text(encoding="utf-8"))
    include_sheets = raw.get("include_sheets")
    if include_sheets is not None and not isinstance(include_sheets, list):
        raise ValueError("include_sheets must be expressed as a YAML list.")

    document_categories = raw.get("document_categories")
    if document_categories is not None and not isinstance(document_categories, dict):
        raise ValueError("document_categories must be expressed as a YAML mapping.")

    product_aliases = raw.get("product_aliases")
    if product_aliases is not None and not isinstance(product_aliases, dict):
        raise ValueError("product_aliases must be expressed as a YAML mapping.")

    product_translations = raw.get("product_translations")
    if product_translations is not None and not isinstance(product_translations, dict):
        raise ValueError("product_translations must be expressed as a YAML mapping.")

    unit_overrides = raw.get("unit_overrides")
    if unit_overrides is not None and not isinstance(unit_overrides, dict):
        raise ValueError("unit_overrides must be expressed as a YAML mapping.")

    return WorkbookConfig(
        unit_mode=str(raw.get("unit_mode", "standard") or "standard"),
        include_sheets=[str(value) for value in include_sheets] if include_sheets else None,
        document_categories={
            str(key): int(value)
            for key, value in (document_categories or {}).items()
        },
        product_aliases=_normalize_alias_map(product_aliases),
        product_translations={
            normalize_text(str(key)): normalize_text(str(value))
            for key, value in (product_translations or {}).items()
        },
        unit_overrides={
            normalize_text(str(key)): str(value)
            for key, value in (unit_overrides or {}).items()
        },
    )
