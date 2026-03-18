from __future__ import annotations

from pathlib import Path
import unicodedata

SPECIAL_TONNES_OR_Q_PRODUCTS = {
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
INPUT_VARIABLES = {
    "production",
    "imports",
    "exports",
    "production (k20)",
    "consumption",
    "stocks",
}


def normalize_text(value: str) -> str:
    normalized = unicodedata.normalize("NFKD", value)
    without_accents = "".join(char for char in normalized if not unicodedata.combining(char))
    return " ".join(without_accents.replace("_", " ").strip().lower().split())



def derive_product_from_document(document_name: str) -> str:
    stem = Path(document_name).stem
    tokens = [token for token in stem.split("_") if token]
    if not tokens:
        return ""

    year_idx = next((idx for idx, token in enumerate(tokens) if len(token) == 4 and token.isdigit()), None)
    if year_idx is None:
        return normalize_text(stem)

    product_start = year_idx + 1
    while product_start < len(tokens) and tokens[product_start].isdigit():
        product_start += 1

    product_tokens = tokens[product_start:] or tokens[-1:]
    return normalize_text(" ".join(product_tokens))



def assign_unit(variable: str, product: str, category: int | None, *, mode: str = "standard") -> str:
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
    if variable == "production" and product == "huevos":
        return "1000000 eggs" if category == 1 else "1000 eggs"
    if variable == "production" and product == "sericultura huevos":
        return "1000 hg" if category == 1 else "hg"
    if variable == "production" and product == "sericultura capullos":
        return "1000 kg" if category == 1 else "kg"
    if variable == "production" and product == "te":
        return "1000 kg" if category == 1 else "kg"
    if variable in {"imports", "exports"} and product in SPECIAL_TONNES_OR_Q_PRODUCTS:
        return "tonnes" if category == 1 else "q"
    if variable in {"production", "imports", "exports"} and product == "vino":
        return "1000 hl" if category == 1 else "hl"
    if variable in {"imports", "exports"} and product in {"bovino", "porcino"}:
        return "1000 heads" if category == 1 else "heads"
    if variable in {"production", "imports", "exports"}:
        return "1000 q" if category == 1 else "q"
    if variable == "production" and product == "goma":
        return "1000 kg"
    if variable in {"laying hens", "livestock"}:
        return "1000 heads" if category == 1 else "heads"
    return ""
