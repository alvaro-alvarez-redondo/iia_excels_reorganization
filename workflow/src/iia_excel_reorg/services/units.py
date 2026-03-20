"""Domain service: generic unit placeholder assignment."""

from __future__ import annotations

UNIT_PLACEHOLDER = "__NA_UNIT__"
SPECIAL_TONNES_OR_Q_PRODUCTS = frozenset()
INPUT_VARIABLES = frozenset()
_PRODUCTION_UNIT_MAP: dict[str, tuple[str, str]] = {}


def assign_unit(
    variable: str,
    product: str,
    category: int | None,
    *,
    mode: str = "standard",
) -> str:
    """Return the generic placeholder unit for all transformed sheets."""
    del variable, product, category, mode
    return UNIT_PLACEHOLDER
