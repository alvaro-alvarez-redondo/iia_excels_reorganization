from __future__ import annotations

import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parent
SRC_ROOT = REPO_ROOT / "workflow" / "src"
DEFAULT_INPUT_DIR = REPO_ROOT / "data" / "10-raw_imports"
DEFAULT_TEMPLATE_PATH = REPO_ROOT / "data" / "lists" / "footnote_mapping_template.xlsx"


def main() -> None:
    """Run the independent footnote harmonization pipeline from repo root.

    This wrapper enables one-file execution (for example via VS Code Run button)
    without requiring package installation first.
    """
    sys.path.insert(0, str(SRC_ROOT))

    from iia_excel_reorg.footnote_pipeline import main as footnote_main

    if len(sys.argv) == 1:
        sys.argv.extend(
            [
                "generate-template",
                str(DEFAULT_INPUT_DIR),
                str(DEFAULT_TEMPLATE_PATH),
            ]
        )

    footnote_main()


if __name__ == "__main__":
    main()
