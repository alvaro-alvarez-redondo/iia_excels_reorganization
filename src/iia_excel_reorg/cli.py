from __future__ import annotations

import argparse
from pathlib import Path

from .config import load_config
from .transformer import transform_workbook



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



def _iter_workbooks(path: Path) -> list[Path]:
    if path.is_file():
        return [path]
    patterns = ("*.xlsx", "*.xlsm")
    workbooks: list[Path] = []
    for pattern in patterns:
        workbooks.extend(sorted(path.glob(pattern)))
    return workbooks



def main() -> None:
    parser = build_parser()
    args = parser.parse_args()

    input_path = Path(args.input)
    output_dir = Path(args.output_dir)
    config = load_config(args.config)

    workbooks = _iter_workbooks(input_path)
    if not workbooks:
        parser.error(f"No Excel workbooks found in: {input_path}")

    output_dir.mkdir(parents=True, exist_ok=True)
    for workbook in workbooks:
        output_name = f"{config.canonical_name_for_document(workbook)}.xlsx"
        output_path = output_dir / output_name
        transform_workbook(workbook, output_path, config=config)
        print(f"Wrote {output_path}")


if __name__ == "__main__":
    main()
