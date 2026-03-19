from __future__ import annotations

import re
import unicodedata
from pathlib import Path

# Matches all Unicode combining (diacritic) characters after NFKD normalization.
# Covers the standard combining diacritics blocks; faster than a per-character
# unicodedata.combining() call inside a generator expression.
_COMBINING_RE = re.compile(
    r"[\u0300-\u036f"   # Combining Diacritical Marks
    r"\u1dc0-\u1dff"    # Combining Diacritical Marks Supplement
    r"\u20d0-\u20ff"    # Combining Diacritical Marks for Symbols
    r"\ufe20-\ufe2f]+"  # Combining Half Marks
)

SPECIAL_TONNES_OR_Q_PRODUCTS = frozenset((
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
))
INPUT_VARIABLES = frozenset((
    "production",
    "imports",
    "exports",
    "production (k20)",
    "consumption",
    "stocks",
))

# Unit lookup tables keyed by (variable, product) for production-specific rules.
# Each entry maps category 1 → unit and category ≠ 1 → unit.
_PRODUCTION_UNIT_MAP: dict[str, tuple[str, str]] = {
    "huevos":             ("1000000 eggs", "1000 eggs"),
    "sericultura huevos": ("1000 hg",      "hg"),
    "sericultura capullos": ("1000 kg",    "kg"),
    "te":                 ("1000 kg",      "kg"),
}


def normalize_text(value: str) -> str:
    """Return *value* lowercased, accent-stripped, and whitespace-normalized.

    Uses a compiled regex to strip Unicode combining characters after NFKD
    decomposition — significantly faster than the equivalent per-character
    ``unicodedata.combining()`` generator expression for long strings.
    """
    decomposed = unicodedata.normalize("NFKD", value)
    without_accents = _COMBINING_RE.sub("", decomposed)
    return " ".join(without_accents.replace("_", " ").strip().lower().split())


def derive_product_from_document(document_name: str) -> str:
    """Infer the product name from an Excel document filename stem.

    Tokens after the first 4-digit year token (skipping any following numeric
    tokens) are joined and normalized.  Falls back to the full stem when no
    year token is found.
    """
    stem = Path(document_name).stem
    tokens = [token for token in stem.split("_") if token]
    if not tokens:
        return ""

    year_idx = next(
        (idx for idx, token in enumerate(tokens) if len(token) == 4 and token.isdigit()),
        None,
    )
    if year_idx is None:
        return normalize_text(stem)

    product_start = year_idx + 1
    while product_start < len(tokens) and tokens[product_start].isdigit():
        product_start += 1

    product_tokens = tokens[product_start:] or tokens[-1:]
    return normalize_text(" ".join(product_tokens))


def assign_unit(variable: str, product: str, category: int | None, *, mode: str = "standard") -> str:
    """Return the measurement unit string for *variable*/*product*/*category*.

    Parameters
    ----------
    variable:
        Normalized sheet/variable name (e.g. ``"area"``, ``"production"``).
    product:
        Normalized product name derived from the source document.
    category:
        Numeric size category from the configuration (``1`` = large/thousands,
        other values = smaller unit).  ``None`` means unknown → returns ``""``.
    mode:
        ``"inputs"`` applies a simplified rule set; any other value uses the
        full standard rule set.
    """
    if category is None:
        return ""

    variable = normalize_text(variable)
    product = normalize_text(product)
    mode = normalize_text(mode)

    if mode == "inputs":
        if variable in INPUT_VARIABLES:
            return "1000 tonnes" if category == 1 else "1000 kg"
        return ""

    if variable in {"area", "bearing area", "planted area", "tappable area"}:
        return "1000 ha" if category == 1 else "ha"

    if variable == "production":
        unit_pair = _PRODUCTION_UNIT_MAP.get(product)
        if unit_pair:
            return unit_pair[0] if category == 1 else unit_pair[1]

    if variable in {"imports", "exports"} and product in SPECIAL_TONNES_OR_Q_PRODUCTS:
        return "tonnes" if category == 1 else "q"

    if variable in {"production", "imports", "exports"} and product == "vino":
        return "1000 hl" if category == 1 else "hl"

    if variable in {"imports", "exports"} and product in {"bovino", "porcino"}:
        return "1000 heads" if category == 1 else "heads"

    if variable in {"production", "imports", "exports"}:
        return "1000 q" if category == 1 else "q"

    if variable in {"laying hens", "livestock"}:
        return "1000 heads" if category == 1 else "heads"

    return ""
