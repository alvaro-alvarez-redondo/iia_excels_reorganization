from __future__ import annotations

import argparse
import re
from pathlib import Path

from .config import load_config
from .transformer import transform_workbook

_EXTRACTED_PAGES_RE = re.compile(r"^extracted_pages_(?P<year>\d{4})_\d{2}$", re.IGNORECASE)
_EXCEL_PATTERNS = ("*.xlsx", "*.xlsm")


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Reorganize historical Excel workbooks into a standardized structure.",
    )
    parser.add_argument(
        "input",
        help="Excel workbook file or directory containing workbook files.",
    )
    parser.add_argument(
        "output_dir",
        help="Directory where transformed workbooks will be written.",
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

        iia_extracted_pages_YYYY/
        └── iia_{subfolder}_YYYY/   (only when a subfolder sits between
                                     extracted_pages_* and the workbook file)

    If no ``extracted_pages_*`` segment is found the file is placed directly in
    the output root (relative path ``Path(".")``).
    """
    parts = workbook_path.parts
    for idx, part in enumerate(parts):
        match = _EXTRACTED_PAGES_RE.match(part)
        if match:
            year = match.group("year")
            parent_dir = f"iia_extracted_pages_{year}"
            # Parts between the extracted_pages directory and the file itself
            intermediate = parts[idx + 1 : -1]
            if intermediate:
                child_dir = f"iia_{intermediate[0]}_{year}"
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
        output_name = f"{config.canonical_name_for_document(workbook)}.xlsx"
        output_path = output_dir / output_name
        transform_workbook(workbook, output_path, config=config)
        print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
