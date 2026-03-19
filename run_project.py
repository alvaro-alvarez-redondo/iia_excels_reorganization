from __future__ import annotations

import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parent
SRC_ROOT = REPO_ROOT / "workflow" / "src"
DEFAULT_CONFIG = REPO_ROOT / "workflow" / "config" / "example.units.yml"


def main() -> None:
    """Run the workbook reorganization workflow from VS Code's Run button.

    When no command-line arguments are provided, this wrapper automatically uses
    the repository's example configuration file so the project can be launched
    directly as a single script from the repository root.
    """
    sys.path.insert(0, str(SRC_ROOT))

    from iia_excel_reorg.cli import main as cli_main

    if len(sys.argv) == 1 and DEFAULT_CONFIG.exists():
        sys.argv.extend(["--config", str(DEFAULT_CONFIG)])

    cli_main()


if __name__ == "__main__":
    main()
