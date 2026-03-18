from __future__ import annotations

import argparse
import re
from pathlib import Path

from .config import load_config
from .naming import sanitize_name
from .transformer import transform_workbook

_EXTRACTED_PAGES_RE = re.compile(r"^extracted_pages_(?P<year>\d{4})_\d{2}$", re.IGNORECASE)
_EXCEL_PATTERNS = ("*.xlsx", "*.xlsm")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Reorganize historical Excel workbooks into a standardized structure.",
    )
    parser.add_argument(
        "input",
        nargs="?",
        default="data/raw inputs",
        help="Excel workbook file or directory containing workbook files. "
             "Defaults to the 'data/raw inputs/' folder in the current directory. "
             "Quote the path when it contains spaces: \"data/raw inputs\".",
    )
    parser.add_argument(
        "output_dir",
        nargs="?",
        default="data/10-raw_imports",
        help="Directory where transformed workbooks will be written. "
             "Defaults to '10-raw_imports/' in the current directory.",
    )
    parser.add_argument(
        "--config",
        help="Path to YAML configuration for categories, aliases, filters, and unit overrides.",
    )
    return parser


def _compute_output_subdir(workbook_path: Path) -> Path:
    """Return the relative output subdirectory for *workbook_path*.

    When the path contains a directory matching ``extracted_pages_YYYY_YY`` the
    output hierarchy is built as::

        fao_extracted_pages_YYYY/
        └── fao_{subfolder}_YYYY/

    The subfolder is determined in this order:

    1. If a directory sits **between** ``extracted_pages_*`` and the workbook
       file, its name is used (e.g. ``extracted_pages_*/crops/file.xlsx``
       → ``fao_crops_YYYY``).
    2. Otherwise the directory that sits **directly above**
       ``extracted_pages_*`` is used (e.g. ``trade/extracted_pages_*/file.xlsx``
       → ``fao_trade_YYYY``).  This handles the common structure where the
       yearbook topic folder wraps the year directory.

    If no ``extracted_pages_*`` segment is found the file is placed directly in
    the output root (relative path ``Path(".")``).
    """
    parts = workbook_path.parts
    for idx, part in enumerate(parts):
        match = _EXTRACTED_PAGES_RE.match(part)
        if match:
            year = match.group("year")
            parent_dir = f"fao_extracted_pages_{year}"
            # Priority 1: a subfolder between extracted_pages_* and the file
            intermediate = parts[idx + 1 : -1]
            if intermediate:
                child_dir = sanitize_name(f"fao_{intermediate[0]}_{year}")
                return Path(parent_dir) / child_dir
            # Priority 2: the folder directly above extracted_pages_*
            if idx > 0:
                topic = parts[idx - 1]
                child_dir = sanitize_name(f"fao_{topic}_{year}")
                return Path(parent_dir) / child_dir
            return Path(parent_dir)
    return Path(".")


def _iter_workbooks(path: Path) -> list[Path]:
    """Return Excel workbooks under *path* (non-recursive flat scan)."""
    if path.is_file():
        return [path]
    workbooks: list[Path] = []
    for pattern in _EXCEL_PATTERNS:
        workbooks.extend(sorted(path.glob(pattern)))
    return workbooks


def _iter_workbooks_structured(root: Path) -> list[tuple[Path, Path]]:
    """Walk *root* recursively and return ``(workbook_path, output_subdir)`` pairs.

    The output subdirectory for each workbook is derived from the
    ``extracted_pages_YYYY_YY`` directory structure when present; otherwise the
    workbook is placed directly under the output root.
    """
    entries: list[tuple[Path, Path]] = []
    for pattern in _EXCEL_PATTERNS:
        for wb_path in sorted(root.rglob(pattern)):
            entries.append((wb_path, _compute_output_subdir(wb_path)))
    return entries


def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    input_path = Path(args.input)
    output_root = Path(args.output_dir)
    config = load_config(args.config)

    if input_path.is_file():
        workbook_entries: list[tuple[Path, Path]] = [(input_path, Path("."))]
    else:
        workbook_entries = _iter_workbooks_structured(input_path)
        if not workbook_entries:
            flat = _iter_workbooks(input_path)
            workbook_entries = [(wb, Path(".")) for wb in flat]

    if not workbook_entries:
        parser.error(f"No Excel workbooks found in: {input_path}")

    for workbook, output_subdir in workbook_entries:
        output_dir = output_root / output_subdir
        output_dir.mkdir(parents=True, exist_ok=True)
        output_name = f"{sanitize_name(config.canonical_name_for_document(workbook))}.xlsx"
        output_path = output_dir / output_name
        transform_workbook(workbook, output_path, config=config)
        print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
