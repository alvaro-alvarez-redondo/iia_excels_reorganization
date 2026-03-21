"""Minimal OOXML workbook reader/writer used by the transformation pipeline."""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from functools import lru_cache, reduce
from itertools import groupby
from pathlib import Path
from typing import TypeAlias
from collections.abc import Sequence
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

CellScalar: TypeAlias = str | int | float | None

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"
_REF_RE = re.compile(r"([A-Za-z]+)(\d+)")

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


@dataclass(slots=True)
class CellData:
    """Lightweight container for a single spreadsheet cell's value and fill color."""

    value: CellScalar = None
    fill_rgb: str | None = None


@dataclass(slots=True)
class SheetData:
    """In-memory representation of a single worksheet."""

    name: str
    cells: dict[tuple[int, int], CellData] = field(default_factory=dict)
    _max_row: int = 0
    _max_column: int = 0

    def set_cell(
        self,
        row: int,
        column: int,
        value: CellScalar,
        fill_rgb: str | None = None,
    ) -> None:
        """Write *value* and optional *fill_rgb* at ``(row, column)``."""
        self.cells[(row, column)] = CellData(
            value=value,
            fill_rgb=_normalize_rgb(fill_rgb),
        )
        if row > self._max_row:
            self._max_row = row
        if column > self._max_column:
            self._max_column = column

    def get_cell(self, row: int, column: int) -> CellData:
        """Return the :class:`CellData` at ``(row, column)``, or an empty cell."""
        return self.cells.get((row, column), CellData())

    def set_row(
        self,
        row: int,
        values: Sequence[CellScalar],
        fills: Sequence[str | None] | None = None,
        *,
        start_column: int = 1,
    ) -> None:
        """Write a row of values and optional fills starting at *start_column*.

        Vectorized implementation: builds all ``(row, col) -> CellData`` pairs
        in a single dict-comprehension and merges them with ``dict.update``,
        replacing the element-wise ``for offset, value in enumerate(values)``
        append loop.  The max-row / max-column bookkeeping is resolved with
        vectorized ``max()`` calls over the new column indices.
        """
        n = len(values)
        if not n:
            return
        normalized_fills = fills or ()
        end_column = start_column + n - 1
        new_cells = {
            (row, start_column + offset): CellData(
                value=value,
                fill_rgb=_normalize_rgb(
                    normalized_fills[offset] if offset < len(normalized_fills) else None
                ),
            )
            for offset, value in enumerate(values)
        }
        self.cells.update(new_cells)
        if row > self._max_row:
            self._max_row = row
        if end_column > self._max_column:
            self._max_column = end_column


    @property
    def max_row(self) -> int:
        """Return the 1-based index of the last occupied row."""
        return self._max_row

    @property
    def max_column(self) -> int:
        """Return the 1-based index of the last occupied column."""
        return self._max_column


@dataclass(slots=True)
class WorkbookData:
    """In-memory representation of an OOXML workbook."""

    sheets: list[SheetData]


def _normalize_rgb(fill_rgb: str | None) -> str | None:
    """Normalize a hex color string to the 8-character ARGB OOXML form."""
    if not fill_rgb:
        return None
    normalized_rgb = fill_rgb.strip().upper()
    return f"FF{normalized_rgb}" if len(normalized_rgb) == 6 else normalized_rgb


def read_workbook(path: str | Path) -> WorkbookData:
    """Parse an ``.xlsx`` or ``.xlsm`` file into :class:`WorkbookData`.

    Vectorized: both the per-sheet cell-parsing loop and the outer sheet-loading
    loop are replaced with a single map + bulk dict-update approach, eliminating
    two levels of explicit ``for`` iteration.
    """
    workbook_path = Path(path)
    with ZipFile(workbook_path) as archive:
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        relationships_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        sheet_targets = _resolve_sheet_targets(workbook_root, relationships_root)
        shared_strings = _read_shared_strings(archive)
        fill_map = _read_fill_map(archive)

        def _load_sheet(name_target: tuple[str, str]) -> SheetData:
            """Parse one worksheet XML into a :class:`SheetData` instance."""
            name, target = name_target
            sheet_root = ET.fromstring(archive.read(f"xl/{target}"))
            # Build all cells in one dict comprehension; the nested
            # ``for row, col in (_split_ref(...),)`` is a let-binding idiom that
            # avoids a temporary variable without adding another loop level.
            bulk_cells: dict[tuple[int, int], CellData] = {
                (row, col): CellData(
                    value=_read_cell_value(cell, shared_strings),
                    fill_rgb=_normalize_rgb(fill_map.get(int(cell.attrib.get("s", "0")))),
                )
                for cell in sheet_root.findall(f".//{{{MAIN_NS}}}c")
                for row, col in (_split_ref(cell.attrib["r"]),)
            }
            sheet = SheetData(name=name)
            sheet.cells.update(bulk_cells)
            if bulk_cells:
                sheet._max_row = max(r for r, _ in bulk_cells)
                sheet._max_column = max(c for _, c in bulk_cells)
            return sheet

        return WorkbookData(sheets=list(map(_load_sheet, sheet_targets)))


def write_workbook(path: str | Path, workbook: WorkbookData) -> Path:
    """Serialize *workbook* to an OOXML ``.xlsx`` file and return the path.

    Vectorized: the per-sheet write loop is replaced with a list comprehension
    that renders and stores all sheets in one pass.
    """
    output_path = Path(path)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    fill_styles = _collect_fill_styles(workbook)

    with ZipFile(output_path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _render_content_types(workbook))
        archive.writestr("_rels/.rels", _render_root_relationships())
        archive.writestr("xl/workbook.xml", _render_workbook(workbook))
        archive.writestr(
            "xl/_rels/workbook.xml.rels",
            _render_workbook_relationships(workbook),
        )
        archive.writestr("xl/styles.xml", _render_styles(fill_styles))
        [
            archive.writestr(
                f"xl/worksheets/sheet{index}.xml",
                _render_sheet(sheet, fill_styles),
            )
            for index, sheet in enumerate(workbook.sheets, start=1)
        ]
    return output_path


def _resolve_sheet_targets(
    workbook_root: ET.Element,
    rels_root: ET.Element,
) -> list[tuple[str, str]]:
    """Return ``(sheet_name, target_path)`` pairs from workbook relationships."""
    relationships = {
        relation.attrib["Id"]: relation.attrib["Target"]
        for relation in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    return [
        (sheet.attrib["name"], relationships[sheet.attrib[f"{{{REL_NS}}}id"]])
        for sheet in workbook_root.findall(f".//{{{MAIN_NS}}}sheet")
    ]


def _read_shared_strings(archive: ZipFile) -> list[str]:
    """Extract the shared-strings table from *archive*, or return an empty list."""
    if "xl/sharedStrings.xml" not in archive.NameToInfo:
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    return [
        "".join(node.text or "" for node in shared_item.findall(f".//{{{MAIN_NS}}}t"))
        for shared_item in root.findall(f"{{{MAIN_NS}}}si")
    ]


def _read_fill_map(archive: ZipFile) -> dict[int, str | None]:
    """Build a mapping from cell-style index to fill RGB color string.

    Vectorized: both the fill-color extraction and the style-index mapping are
    expressed as single comprehensions, eliminating two explicit ``for`` loops.
    """
    if "xl/styles.xml" not in archive.NameToInfo:
        return {}
    root = ET.fromstring(archive.read("xl/styles.xml"))

    def _fill_rgb(fill: ET.Element) -> str | None:
        fg = fill.find(f".//{{{MAIN_NS}}}fgColor")
        return fg.attrib.get("rgb") if fg is not None else None

    fill_colors: list[str | None] = [
        _fill_rgb(fill)
        for fill in root.findall(f".//{{{MAIN_NS}}}fills/{{{MAIN_NS}}}fill")
    ]
    return {
        index: fill_colors[fill_id] if (fill_id := int(xf.attrib.get("fillId", "0"))) < len(fill_colors) else None
        for index, xf in enumerate(root.findall(f".//{{{MAIN_NS}}}cellXfs/{{{MAIN_NS}}}xf"))
    }


def _read_cell_value(cell: ET.Element, shared_strings: list[str]) -> CellScalar:
    """Parse the value of a ``<c>`` element."""
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        text = cell.find(f".//{{{MAIN_NS}}}t")
        return (text.text or "") if text is not None else ""

    value_node = cell.find(f"{{{MAIN_NS}}}v")
    if value_node is None:
        return None
    raw_value = value_node.text or ""
    if cell_type == "s":
        return shared_strings[int(raw_value)]
    if cell_type == "str":
        return raw_value
    try:
        if raw_value.isdigit() or (raw_value.startswith("-") and raw_value[1:].isdigit()):
            return int(raw_value)
        return float(raw_value)
    except ValueError:
        return raw_value


def _split_ref(ref: str) -> tuple[int, int]:
    """Parse an OOXML cell reference into ``(row, column)`` indices."""
    match = _REF_RE.match(ref)
    if match is None:
        raise ValueError(f"Invalid cell reference: {ref!r}")
    return int(match.group(2)), _column_index_from_letters(match.group(1))


@lru_cache(maxsize=256)
def _column_index_from_letters(letters: str) -> int:
    """Convert a column-letter string to a 1-based column index."""
    return reduce(lambda acc, char: acc * 26 + (ord(char) - 64), letters.upper(), 0)


@lru_cache(maxsize=256)
def _column_letters(index: int) -> str:
    """Convert a 1-based column index to its OOXML letter string."""
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _collect_fill_styles(workbook: WorkbookData) -> dict[str, int]:
    """Return a mapping of RGB color to fill-style index.

    Vectorized implementation: replace the nested ``for sheet … for cell``
    loop with a single generator expression that streams all cell fill colors
    across the workbook, then builds the de-duplication dict in one pass using
    ``dict.fromkeys`` (O(n) hash-map construction).
    """
    unique_fills = dict.fromkeys(
        rgb
        for sheet in workbook.sheets
        for cell in sheet.cells.values()
        if (rgb := cell.fill_rgb)
    )
    return {rgb: idx for idx, rgb in enumerate(unique_fills, start=2)}


def _render_content_types(workbook: WorkbookData) -> bytes:
    """Render the ``[Content_Types].xml`` OOXML part as UTF-8 bytes.

    Vectorized: per-sheet ``<Override>`` elements are appended in one
    ``root.extend()`` call driven by a generator expression.
    """
    root = ET.Element(f"{{{CONTENT_NS}}}Types")
    ET.SubElement(
        root,
        f"{{{CONTENT_NS}}}Default",
        Extension="rels",
        ContentType="application/vnd.openxmlformats-package.relationships+xml",
    )
    ET.SubElement(
        root,
        f"{{{CONTENT_NS}}}Default",
        Extension="xml",
        ContentType="application/xml",
    )
    ET.SubElement(
        root,
        f"{{{CONTENT_NS}}}Override",
        PartName="/xl/workbook.xml",
        ContentType=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
        ),
    )
    ET.SubElement(
        root,
        f"{{{CONTENT_NS}}}Override",
        PartName="/xl/styles.xml",
        ContentType=(
            "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
        ),
    )
    root.extend(
        ET.Element(
            f"{{{CONTENT_NS}}}Override",
            PartName=f"/xl/worksheets/sheet{index}.xml",
            ContentType=(
                "application/vnd.openxmlformats-officedocument"
                ".spreadsheetml.worksheet+xml"
            ),
        )
        for index in range(1, len(workbook.sheets) + 1)
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _render_root_relationships() -> bytes:
    """Render the ``_rels/.rels`` OOXML part pointing to ``xl/workbook.xml``."""
    root = ET.Element(f"{{{PKG_REL_NS}}}Relationships")
    ET.SubElement(
        root,
        f"{{{PKG_REL_NS}}}Relationship",
        Id="rId1",
        Type=(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
            "officeDocument"
        ),
        Target="xl/workbook.xml",
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _render_workbook(workbook: WorkbookData) -> bytes:
    """Render the ``xl/workbook.xml`` OOXML part listing all sheets.

    Vectorized: per-sheet ``<sheet>`` elements are built by a helper and
    appended via ``sheets_element.extend()``, replacing an explicit ``for`` loop.
    """
    root = ET.Element(f"{{{MAIN_NS}}}workbook")
    sheets_element = ET.SubElement(root, f"{{{MAIN_NS}}}sheets")

    def _make_sheet_element(index_sheet: tuple[int, SheetData]) -> ET.Element:
        index, sheet = index_sheet
        element = ET.Element(
            f"{{{MAIN_NS}}}sheet",
            name=sheet.name,
            sheetId=str(index),
        )
        element.set(f"{{{REL_NS}}}id", f"rId{index}")
        return element

    sheets_element.extend(map(_make_sheet_element, enumerate(workbook.sheets, start=1)))
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _render_workbook_relationships(workbook: WorkbookData) -> bytes:
    """Render the workbook relationship part for sheets and styles.

    Vectorized: per-sheet ``<Relationship>`` elements are appended in one
    ``root.extend()`` call driven by a generator expression.
    """
    root = ET.Element(f"{{{PKG_REL_NS}}}Relationships")
    root.extend(
        ET.Element(
            f"{{{PKG_REL_NS}}}Relationship",
            Id=f"rId{index}",
            Type=(
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
                "worksheet"
            ),
            Target=f"worksheets/sheet{index}.xml",
        )
        for index in range(1, len(workbook.sheets) + 1)
    )
    ET.SubElement(
        root,
        f"{{{PKG_REL_NS}}}Relationship",
        Id=f"rId{len(workbook.sheets) + 1}",
        Type=(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/"
            "styles"
        ),
        Target="styles.xml",
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _render_styles(fill_styles: dict[str, int]) -> bytes:
    """Render the ``xl/styles.xml`` OOXML part for unique fill colors."""
    root = ET.Element(f"{{{MAIN_NS}}}styleSheet")
    ET.SubElement(root, f"{{{MAIN_NS}}}numFmts", count="0")
    fonts = ET.SubElement(root, f"{{{MAIN_NS}}}fonts", count="1")
    font = ET.SubElement(fonts, f"{{{MAIN_NS}}}font")
    ET.SubElement(font, f"{{{MAIN_NS}}}sz", val="11")
    ET.SubElement(font, f"{{{MAIN_NS}}}name", val="Calibri")
    ET.SubElement(font, f"{{{MAIN_NS}}}family", val="2")

    fills = ET.SubElement(root, f"{{{MAIN_NS}}}fills", count=str(2 + len(fill_styles)))
    ET.SubElement(
        ET.SubElement(fills, f"{{{MAIN_NS}}}fill"),
        f"{{{MAIN_NS}}}patternFill",
        patternType="none",
    )
    ET.SubElement(
        ET.SubElement(fills, f"{{{MAIN_NS}}}fill"),
        f"{{{MAIN_NS}}}patternFill",
        patternType="gray125",
    )

    def _make_fill_element(rgb: str) -> ET.Element:
        fill = ET.Element(f"{{{MAIN_NS}}}fill")
        pattern = ET.SubElement(fill, f"{{{MAIN_NS}}}patternFill", patternType="solid")
        ET.SubElement(pattern, f"{{{MAIN_NS}}}fgColor", rgb=rgb)
        ET.SubElement(pattern, f"{{{MAIN_NS}}}bgColor", indexed="64")
        return fill

    fills.extend(
        _make_fill_element(rgb)
        for rgb, _ in sorted(fill_styles.items(), key=lambda item: item[1])
    )

    borders = ET.SubElement(root, f"{{{MAIN_NS}}}borders", count="1")
    ET.SubElement(borders, f"{{{MAIN_NS}}}border")
    ET.SubElement(root, f"{{{MAIN_NS}}}cellStyleXfs", count="1").append(
        ET.Element(
            f"{{{MAIN_NS}}}xf",
            numFmtId="0",
            fontId="0",
            fillId="0",
            borderId="0",
        )
    )

    cell_xfs = ET.SubElement(root, f"{{{MAIN_NS}}}cellXfs", count=str(1 + len(fill_styles)))
    ET.SubElement(
        cell_xfs,
        f"{{{MAIN_NS}}}xf",
        numFmtId="0",
        fontId="0",
        fillId="0",
        borderId="0",
        xfId="0",
    )
    cell_xfs.extend(
        ET.Element(
            f"{{{MAIN_NS}}}xf",
            numFmtId="0",
            fontId="0",
            fillId=str(fill_id),
            borderId="0",
            xfId="0",
            applyFill="1",
        )
        for _, fill_id in sorted(fill_styles.items(), key=lambda item: item[1])
    )

    cell_styles = ET.SubElement(root, f"{{{MAIN_NS}}}cellStyles", count="1")
    ET.SubElement(
        cell_styles,
        f"{{{MAIN_NS}}}cellStyle",
        name="Normal",
        xfId="0",
        builtinId="0",
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _render_sheet(sheet: SheetData, fill_styles: dict[str, int]) -> bytes:
    """Render a worksheet as an OOXML ``xl/worksheets/sheetN.xml`` part.

    Vectorized: the outer row loop and inner cell loop are replaced by a local
    helper consumed through ``sheet_data.extend()``.  The inner cell iteration
    is replaced by a list comprehension (side-effects on the row element).
    """
    root = ET.Element(f"{{{MAIN_NS}}}worksheet")
    if sheet.max_row and sheet.max_column:
        ET.SubElement(
            root,
            f"{{{MAIN_NS}}}dimension",
            ref=f"A1:{_column_letters(sheet.max_column)}{sheet.max_row}",
        )
    sheet_data = ET.SubElement(root, f"{{{MAIN_NS}}}sheetData")

    def _make_row_element(row_group: tuple) -> ET.Element:
        """Build one ``<row>`` element with all its non-null ``<c>`` children."""
        row_index, row_iter = row_group
        row_element = ET.Element(f"{{{MAIN_NS}}}row", r=str(row_index))
        [
            _render_cell(row_element, row_index, col_index, cell, fill_styles)
            for (_, col_index), cell in row_iter
            if cell.value is not None
        ]
        return row_element

    sorted_cells = sorted(sheet.cells.items())
    sheet_data.extend(
        map(_make_row_element, groupby(sorted_cells, key=lambda item: item[0][0]))
    )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _render_cell(
    row_element: ET.Element,
    row_index: int,
    column_index: int,
    cell: CellData,
    fill_styles: dict[str, int],
) -> None:
    """Render a single OOXML cell element into *row_element*."""
    attributes = {"r": f"{_column_letters(column_index)}{row_index}"}
    style_id = 0
    if cell.fill_rgb and cell.fill_rgb in fill_styles:
        style_id = fill_styles[cell.fill_rgb] - 1
    if style_id:
        attributes["s"] = str(style_id)

    if isinstance(cell.value, str):
        attributes["t"] = "inlineStr"
        cell_element = ET.SubElement(row_element, f"{{{MAIN_NS}}}c", attributes)
        inline_string = ET.SubElement(cell_element, f"{{{MAIN_NS}}}is")
        text = ET.SubElement(inline_string, f"{{{MAIN_NS}}}t")
        if cell.value.strip() != cell.value:
            text.set(f"{{{XML_NS}}}space", "preserve")
        text.text = cell.value
        return

    cell_element = ET.SubElement(row_element, f"{{{MAIN_NS}}}c", attributes)
    ET.SubElement(cell_element, f"{{{MAIN_NS}}}v").text = str(cell.value)
