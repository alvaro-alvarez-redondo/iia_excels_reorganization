"""Domain service: measurement-unit assignment rules."""

from __future__ import annotations

from ..utils.text import normalize_text

SPECIAL_TONNES_OR_Q_PRODUCTS = frozenset(
    {
        "te",
        "pimienta",
        "tabaco",
        "lupulo",
        "aceite soja",
        "aceite mani",
        "aceite sesamo",
        "aceite semilla ricino",
        "aceite lino",
        "aceite algodon",
        "aceite coco",
        "aceite palma",
        "aceite de palmiste",
        "leche",
        "mantequilla",
        "queso",
        "huevos",
        "derivados huevos",
        "seda",
    }
)
INPUT_VARIABLES = frozenset(
    {
        "production",
        "imports",
        "exports",
        "production (k20)",
        "consumption",
        "stocks",
    }
)
AREA_VARIABLES = frozenset({"area", "bearing area", "planted area", "tappable area"})
TRADE_VARIABLES = frozenset({"production", "imports", "exports"})
HEADCOUNT_VARIABLES = frozenset({"laying hens", "livestock"})
HEADCOUNT_PRODUCTS = frozenset({"bovino", "porcino"})

_PRODUCTION_UNIT_MAP: dict[str, tuple[str, str]] = {
    "huevos": ("1000000 eggs", "1000 eggs"),
    "sericultura huevos": ("1000 hg", "hg"),
    "sericultura capullos": ("1000 kg", "kg"),
    "te": ("1000 kg", "kg"),
}


def _unit_for_category(category: int, large_unit: str, small_unit: str) -> str:
    """Return the category-dependent unit string."""
    return large_unit if category == 1 else small_unit


def assign_unit(
    variable: str,
    product: str,
    category: int | None,
    *,
    mode: str = "standard",
) -> str:
    """Return the measurement unit string for *variable*/*product*/*category*."""
    if category is None:
        return ""

    normalized_variable = normalize_text(variable)
    normalized_product = normalize_text(product)
    normalized_mode = normalize_text(mode)

    if normalized_mode == "inputs":
        if normalized_variable in INPUT_VARIABLES:
            return _unit_for_category(category, "1000 tonnes", "1000 kg")
        return ""

    if normalized_variable in AREA_VARIABLES:
        return _unit_for_category(category, "1000 ha", "ha")

    if normalized_variable == "production":
        production_unit = _PRODUCTION_UNIT_MAP.get(normalized_product)
        if production_unit is not None:
            return _unit_for_category(category, *production_unit)

    if (
        normalized_variable in {"imports", "exports"}
        and normalized_product in SPECIAL_TONNES_OR_Q_PRODUCTS
    ):
        return _unit_for_category(category, "tonnes", "q")

    if normalized_variable in TRADE_VARIABLES and normalized_product == "vino":
        return _unit_for_category(category, "1000 hl", "hl")

    if (
        normalized_variable in {"imports", "exports"}
        and normalized_product in HEADCOUNT_PRODUCTS
    ):
        return _unit_for_category(category, "1000 heads", "heads")

    if normalized_variable in TRADE_VARIABLES:
        return _unit_for_category(category, "1000 q", "q")

    if normalized_variable in HEADCOUNT_VARIABLES:
        return _unit_for_category(category, "1000 heads", "heads")

    return ""
