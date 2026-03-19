"""Backward-compatible re-export of :mod:`iia_excel_reorg.utils.text` and
:mod:`iia_excel_reorg.services.units`.

.. deprecated::
    Import directly from :mod:`iia_excel_reorg.utils` or
    :mod:`iia_excel_reorg.services` in new code.
"""

from .services.units import (  # noqa: F401
    INPUT_VARIABLES,
    SPECIAL_TONNES_OR_Q_PRODUCTS,
    _PRODUCTION_UNIT_MAP,
    assign_unit,
)
from .utils.text import (  # noqa: F401
    _COMBINING_RE,
    derive_product_from_document,
    normalize_text,
)
