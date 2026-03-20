"""Command-line workflow for reorganizing historical Excel workbooks."""

from __future__ import annotations
import argparse
import re
import shutil
import sys
from pathlib import Path
from typing import Callable, TypeAlias
from .config import load_config
from .core.transformer import (
    GeographyIndex,
    ProductIndex,
    UnitFootnoteDocumentIndex,
    transform_workbook,
)
from .utils.naming import sanitize_name
from .utils.text import derive_product_from_document

WorkbookEntry: TypeAlias = tuple[Path, Path]
WorkbookAction: TypeAlias = Callable[[WorkbookEntry], None]
TxtAction: TypeAlias = Callable[[], None]
_EXTRACTED_PAGES_RE = re.compile(
    r"^extracted_pages_(?P<year>\d{4})_\d{2}$",
    re.IGNORECASE,
)
_EXCEL_SUFFIXES = frozenset({".xlsx", ".xlsm"})
DEFAULT_INPUT_DIR = Path("data/raw_inputs")
DEFAULT_OUTPUT_DIR = Path("data/10-raw_imports")
HEMISPHERE_INDEX_FILENAME = "unique_hemisphere_values.txt"
CONTINENT_INDEX_FILENAME = "unique_continent_values.txt"
COUNTRY_INDEX_FILENAME = "unique_country_values.txt"
GEOGRAPHY_INDEX_FILENAME = "unique_geography_values.txt"
PRODUCT_INDEX_FILENAME = "unique_product_values.txt"
UNIT_FOOTNOTE_DOCUMENT_INDEX_FILENAME = "final_docs_with_unit_footnotes.txt"
PROJECT_ROOT = Path(__file__).resolve().parents[3]
LISTS_DIR = PROJECT_ROOT / "data" / "lists"


def build_parser() -> argparse.ArgumentParser:
    """Build and return the CLI argument parser."""
    parser = argparse.ArgumentParser(
        description=(
            "Reorganize historical Excel workbooks into a standardized structure."
        ),
    )
    parser.add_argument(
        "input",
        nargs="?",
        default=str(DEFAULT_INPUT_DIR),
        help=(
            "Excel workbook file or directory containing workbook files. "
            "Defaults to the 'data/raw_inputs/' folder in the current directory. "
            'Quote the path when it contains spaces: "data/raw_inputs".'
        ),
    )
    parser.add_argument(
        "output_dir",
        nargs="?",
        default=str(DEFAULT_OUTPUT_DIR),
        help=(
            "Directory where transformed workbooks will be written. "
            "Defaults to '10-raw_imports/' in the current directory."
        ),
    )
    parser.add_argument(
        "--config",
        help=(
            "Path to YAML configuration for categories, aliases, filters, "
            "and unit overrides."
        ),
    )
    return parser


def _compute_output_subdir(workbook_path: Path) -> Path:
    """Return the relative output subdirectory for *workbook_path*."""
    parts = workbook_path.parts
    for idx, part in enumerate(parts):
        match = _EXTRACTED_PAGES_RE.match(part)
        if match is None:
            continue
        year = match.group("year")
        parent_dir = Path(f"iia_extracted_pages_{year}")
        intermediate = parts[idx + 1 : -1]
        if intermediate:
            child_dir = sanitize_name(f"iia_{intermediate[0]}_{year}")
            return parent_dir / child_dir
        if idx > 0:
            topic = parts[idx - 1]
            child_dir = sanitize_name(f"iia_{topic}_{year}")
            return parent_dir / child_dir
        return parent_dir
    return Path(".")


def _iter_workbooks(path: Path) -> list[Path]:
    """Return Excel workbooks under *path* using a non-recursive scan."""
    if path.is_file():
        return [path]
    return sorted(
        candidate
        for candidate in path.iterdir()
        if candidate.is_file() and candidate.suffix.lower() in _EXCEL_SUFFIXES
    )


def _iter_workbooks_structured(root: Path) -> list[WorkbookEntry]:
    """Walk *root* recursively and return ``(workbook_path, output_subdir)`` pairs."""
    workbook_paths = sorted(
        candidate
        for candidate in root.rglob("*")
        if candidate.is_file() and candidate.suffix.lower() in _EXCEL_SUFFIXES
    )
    return [
        (workbook_path, _compute_output_subdir(workbook_path))
        for workbook_path in workbook_paths
    ]


def _ensure_workspace(input_path: Path, output_root: Path) -> None:
    """Create or reset the input/output workspace directories."""
    if not input_path.exists() and input_path.suffix == "":
        input_path.mkdir(parents=True, exist_ok=True)
    if output_root.exists():
        shutil.rmtree(output_root)
    output_root.mkdir(parents=True, exist_ok=True)
    LISTS_DIR.mkdir(parents=True, exist_ok=True)


def _render_progress_bar(
    label: str,
    current: int,
    total: int,
    width: int = 24,
) -> str:
    """Return a single-line progress bar string."""
    normalized_total = max(total, 1)
    completed = min(width, int(width * current / normalized_total))
    percent = int(100 * current / normalized_total)
    bar = "█" * completed + "·" * (width - completed)
    return f"{label:<21} │{bar}│ {percent:>3}% ({current}/{normalized_total})"


def _run_progress(
    label: str,
    items: list[WorkbookEntry],
    action: WorkbookAction,
) -> None:
    """Run *action* on each item in *items* while updating a progress bar."""
    total = len(items)
    sys.stdout.write(_render_progress_bar(label, 0, total))
    sys.stdout.flush()
    for index, item in enumerate(items, start=1):
        action(item)
        sys.stdout.write("\r" + _render_progress_bar(label, index, total))
        sys.stdout.flush()
    print()


def _run_txt_progress(label: str, actions: list[tuple[str, TxtAction]]) -> None:
    """Run TXT generation actions while updating a dedicated progress bar."""
    total = len(actions)
    sys.stdout.write(_render_progress_bar(label, 0, total))
    sys.stdout.flush()
    for index, (_, action) in enumerate(actions, start=1):
        action()
        sys.stdout.write("\r" + _render_progress_bar(label, index, total))
        sys.stdout.flush()
    print()


def main() -> None:
    """Entry point for the ``iia-excel-reorg`` command-line tool."""
    parser = build_parser()
    args = parser.parse_args()
    input_path = Path(args.input)
    output_root = Path(args.output_dir)
    _ensure_workspace(input_path, output_root)
    config = load_config(args.config)
    if input_path.is_file():
        workbook_entries: list[WorkbookEntry] = [(input_path, Path("."))]
    else:
        workbook_entries = _iter_workbooks_structured(input_path)
        if not workbook_entries:
            workbook_entries = [
                (path, Path(".")) for path in _iter_workbooks(input_path)
            ]
    if not workbook_entries:
        print(f"No Excel workbooks found in: {input_path}")
        print(
            "Created workspace folders if needed. Add source Excel files there "
            "and run again."
        )
        return
    geography_index = GeographyIndex()
    product_index = ProductIndex()
    unit_footnote_document_index = UnitFootnoteDocumentIndex()

    def prepare_output(entry: WorkbookEntry) -> None:
        """Create the output subdirectory for *entry* if needed."""
        _, output_subdir = entry
        (output_root / output_subdir).mkdir(parents=True, exist_ok=True)

    def transform_entry(entry: WorkbookEntry) -> None:
        """Transform a single workbook and write it into the output tree."""
        workbook_path, output_subdir = entry
        output_dir = output_root / output_subdir
        output_name = (
            f"{sanitize_name(config.canonical_name_for_document(workbook_path))}.xlsx"
        )
        output_path = output_dir / output_name
        transform_workbook(
            workbook_path,
            output_path,
            config=config,
            geography_index=geography_index,
            unit_footnote_document_index=unit_footnote_document_index,
        )
        product_index.add_product(derive_product_from_document(output_path.name))

    _run_progress("Reorganizing folders", workbook_entries, prepare_output)
    _run_progress("Reorganizing excels", workbook_entries, transform_entry)
    _run_txt_progress(
        "Generating txt lists",
        [
            (
                GEOGRAPHY_INDEX_FILENAME,
                lambda: geography_index.write_txt(LISTS_DIR / GEOGRAPHY_INDEX_FILENAME),
            ),
            (
                PRODUCT_INDEX_FILENAME,
                lambda: product_index.write_txt(LISTS_DIR / PRODUCT_INDEX_FILENAME),
            ),
            (
                UNIT_FOOTNOTE_DOCUMENT_INDEX_FILENAME,
                lambda: unit_footnote_document_index.write_txt(
                    LISTS_DIR / UNIT_FOOTNOTE_DOCUMENT_INDEX_FILENAME
                ),
            ),
        ],
    )


if __name__ == "__main__":
    main()
