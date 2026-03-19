from __future__ import annotations

import re
from dataclasses import dataclass, field
from functools import reduce
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CONTENT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
XML_NS = "http://www.w3.org/XML/1998/namespace"

# Pre-compiled pattern to split an OOXML cell reference (e.g. "AB12") into
# its column letters and row number in a single pass.
_REF_RE = re.compile(r"([A-Za-z]+)(\d+)")

ET.register_namespace("", MAIN_NS)
ET.register_namespace("r", REL_NS)


@dataclass(slots=True)
class CellData:
    """Lightweight container for a single spreadsheet cell's value and fill colour."""

    value: str | int | float | None = None
    fill_rgb: str | None = None


@dataclass(slots=True)
class SheetData:
    """In-memory representation of a single worksheet."""

    name: str
    cells: dict[tuple[int, int], CellData] = field(default_factory=dict)

    def set_cell(self, row: int, column: int, value: str | int | float | None, fill_rgb: str | None = None) -> None:
        """Write *value* (and optional *fill_rgb*) at ``(row, column)``."""
        self.cells[(row, column)] = CellData(value=value, fill_rgb=_normalize_rgb(fill_rgb))

    def get_cell(self, row: int, column: int) -> CellData:
        """Return the :class:`CellData` at ``(row, column)``, or an empty cell."""
        return self.cells.get((row, column), CellData())

    @property
    def max_row(self) -> int:
        """1-based index of the last occupied row, or 0 when the sheet is empty."""
        return max((row for row, _ in self.cells), default=0)

    @property
    def max_column(self) -> int:
        """1-based index of the last occupied column, or 0 when the sheet is empty."""
        return max((column for _, column in self.cells), default=0)


@dataclass(slots=True)
class WorkbookData:
    """In-memory representation of an OOXML workbook."""

    sheets: list[SheetData]


def _normalize_rgb(fill_rgb: str | None) -> str | None:
    """Normalise a hex colour string to the 8-character ARGB form used by OOXML.

    A 6-character RGB value is prefixed with ``FF`` (fully opaque).  ``None``
    and empty strings are returned as ``None``.
    """
    if not fill_rgb:
        return None
    rgb = fill_rgb.strip().upper()
    if len(rgb) == 6:
        return f"FF{rgb}"
    return rgb


def read_workbook(path: str | Path) -> WorkbookData:
    """Parse an ``.xlsx`` / ``.xlsm`` file and return its contents as :class:`WorkbookData`."""
    path = Path(path)
    with ZipFile(path) as archive:
        workbook_root = ET.fromstring(archive.read("xl/workbook.xml"))
        rels_root = ET.fromstring(archive.read("xl/_rels/workbook.xml.rels"))
        sheet_targets = _resolve_sheet_targets(workbook_root, rels_root)
        shared_strings = _read_shared_strings(archive)
        fill_map = _read_fill_map(archive)
        sheets: list[SheetData] = []
        for name, target in sheet_targets:
            sheet_root = ET.fromstring(archive.read(f"xl/{target}"))
            sheet = SheetData(name=name)
            for cell in sheet_root.findall(f".//{{{MAIN_NS}}}c"):
                ref = cell.attrib["r"]
                row, column = _split_ref(ref)
                style_idx = int(cell.attrib.get("s", "0"))
                fill_rgb = fill_map.get(style_idx)
                value = _read_cell_value(cell, shared_strings)
                sheet.set_cell(row, column, value, fill_rgb=fill_rgb)
            sheets.append(sheet)
        return WorkbookData(sheets=sheets)


def write_workbook(path: str | Path, workbook: WorkbookData) -> Path:
    """Serialise *workbook* to an OOXML ``.xlsx`` file at *path* and return the path."""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)

    fill_styles = _collect_fill_styles(workbook)

    with ZipFile(path, "w", compression=ZIP_DEFLATED) as archive:
        archive.writestr("[Content_Types].xml", _render_content_types(workbook))
        archive.writestr("_rels/.rels", _render_root_relationships())
        archive.writestr("xl/workbook.xml", _render_workbook(workbook))
        archive.writestr("xl/_rels/workbook.xml.rels", _render_workbook_relationships(workbook))
        archive.writestr("xl/styles.xml", _render_styles(fill_styles))
        for idx, sheet in enumerate(workbook.sheets, start=1):
            archive.writestr(f"xl/worksheets/sheet{idx}.xml", _render_sheet(sheet, fill_styles))
    return path


def _resolve_sheet_targets(workbook_root: ET.Element, rels_root: ET.Element) -> list[tuple[str, str]]:
    rels = {
        rel.attrib["Id"]: rel.attrib["Target"]
        for rel in rels_root.findall(f"{{{PKG_REL_NS}}}Relationship")
    }
    sheets = []
    for sheet in workbook_root.findall(f".//{{{MAIN_NS}}}sheet"):
        rel_id = sheet.attrib[f"{{{REL_NS}}}id"]
        sheets.append((sheet.attrib["name"], rels[rel_id]))
    return sheets


def _read_shared_strings(archive: ZipFile) -> list[str]:
    if "xl/sharedStrings.xml" not in archive.namelist():
        return []
    root = ET.fromstring(archive.read("xl/sharedStrings.xml"))
    strings: list[str] = []
    for si in root.findall(f"{{{MAIN_NS}}}si"):
        texts = [node.text or "" for node in si.findall(f".//{{{MAIN_NS}}}t")]
        strings.append("".join(texts))
    return strings


def _read_fill_map(archive: ZipFile) -> dict[int, str | None]:
    if "xl/styles.xml" not in archive.namelist():
        return {}
    root = ET.fromstring(archive.read("xl/styles.xml"))
    fills = []
    for fill in root.findall(f".//{{{MAIN_NS}}}fills/{{{MAIN_NS}}}fill"):
        fg = fill.find(f".//{{{MAIN_NS}}}fgColor")
        fills.append(fg.attrib.get("rgb") if fg is not None else None)
    style_map: dict[int, str | None] = {}
    for idx, xf in enumerate(root.findall(f".//{{{MAIN_NS}}}cellXfs/{{{MAIN_NS}}}xf")):
        fill_id = int(xf.attrib.get("fillId", "0"))
        style_map[idx] = fills[fill_id] if fill_id < len(fills) else None
    return style_map


def _read_cell_value(cell: ET.Element, shared_strings: list[str]) -> str | int | float | None:
    cell_type = cell.attrib.get("t")
    if cell_type == "inlineStr":
        text = cell.find(f".//{{{MAIN_NS}}}t")
        return (text.text or "") if text is not None else ""
    value_node = cell.find(f"{{{MAIN_NS}}}v")
    if value_node is None:
        return None
    raw = value_node.text or ""
    if cell_type == "s":
        return shared_strings[int(raw)]
    if cell_type == "str":
        return raw
    try:
        return int(raw) if raw.isdigit() or (raw.startswith("-") and raw[1:].isdigit()) else float(raw)
    except ValueError:
        return raw


def _split_ref(ref: str) -> tuple[int, int]:
    """Parse an OOXML cell reference (e.g. ``"AB12"``) into ``(row, column)`` indices.

    Uses a pre-compiled regex instead of a character-by-character loop.
    """
    m = _REF_RE.match(ref)
    if m is None:
        raise ValueError(f"Invalid cell reference: {ref!r}")
    return int(m.group(2)), _column_index_from_letters(m.group(1))


def _column_index_from_letters(letters: str) -> int:
    """Convert a column letter string (e.g. ``"AB"``) to a 1-based column index.

    Uses :func:`functools.reduce` for a concise, loop-free implementation.
    Each character contributes its 1-based alphabetic position (A=1 … Z=26),
    so ``ord(c) - 64`` maps ``'A'`` → 1, ``'B'`` → 2, …, ``'Z'`` → 26.
    """
    return reduce(lambda acc, c: acc * 26 + (ord(c) - 64), letters.upper(), 0)


def _column_letters(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _collect_fill_styles(workbook: WorkbookData) -> dict[str, int]:
    fills: dict[str, int] = {}
    next_fill_id = 2
    for sheet in workbook.sheets:
        for cell in sheet.cells.values():
            if cell.fill_rgb and cell.fill_rgb not in fills:
                fills[cell.fill_rgb] = next_fill_id
                next_fill_id += 1
    return fills


def _render_content_types(workbook: WorkbookData) -> bytes:
    root = ET.Element(f"{{{CONTENT_NS}}}Types")
    ET.SubElement(root, f"{{{CONTENT_NS}}}Default", Extension="rels", ContentType="application/vnd.openxmlformats-package.relationships+xml")
    ET.SubElement(root, f"{{{CONTENT_NS}}}Default", Extension="xml", ContentType="application/xml")
    ET.SubElement(root, f"{{{CONTENT_NS}}}Override", PartName="/xl/workbook.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml")
    ET.SubElement(root, f"{{{CONTENT_NS}}}Override", PartName="/xl/styles.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml")
    for idx, _sheet in enumerate(workbook.sheets, start=1):
        ET.SubElement(root, f"{{{CONTENT_NS}}}Override", PartName=f"/xl/worksheets/sheet{idx}.xml", ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _render_root_relationships() -> bytes:
    root = ET.Element(f"{{{PKG_REL_NS}}}Relationships")
    ET.SubElement(root, f"{{{PKG_REL_NS}}}Relationship", Id="rId1", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument", Target="xl/workbook.xml")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _render_workbook(workbook: WorkbookData) -> bytes:
    root = ET.Element(f"{{{MAIN_NS}}}workbook")
    sheets_el = ET.SubElement(root, f"{{{MAIN_NS}}}sheets")
    for idx, sheet in enumerate(workbook.sheets, start=1):
        sheet_el = ET.SubElement(sheets_el, f"{{{MAIN_NS}}}sheet", name=sheet.name, sheetId=str(idx))
        sheet_el.set(f"{{{REL_NS}}}id", f"rId{idx}")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _render_workbook_relationships(workbook: WorkbookData) -> bytes:
    root = ET.Element(f"{{{PKG_REL_NS}}}Relationships")
    for idx, _sheet in enumerate(workbook.sheets, start=1):
        ET.SubElement(root, f"{{{PKG_REL_NS}}}Relationship", Id=f"rId{idx}", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet", Target=f"worksheets/sheet{idx}.xml")
    ET.SubElement(root, f"{{{PKG_REL_NS}}}Relationship", Id=f"rId{len(workbook.sheets) + 1}", Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles", Target="styles.xml")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _render_styles(fill_styles: dict[str, int]) -> bytes:
    root = ET.Element(f"{{{MAIN_NS}}}styleSheet")
    ET.SubElement(root, f"{{{MAIN_NS}}}numFmts", count="0")
    fonts = ET.SubElement(root, f"{{{MAIN_NS}}}fonts", count="1")
    font = ET.SubElement(fonts, f"{{{MAIN_NS}}}font")
    ET.SubElement(font, f"{{{MAIN_NS}}}sz", val="11")
    ET.SubElement(font, f"{{{MAIN_NS}}}name", val="Calibri")
    ET.SubElement(font, f"{{{MAIN_NS}}}family", val="2")

    fills = ET.SubElement(root, f"{{{MAIN_NS}}}fills", count=str(2 + len(fill_styles)))
    ET.SubElement(ET.SubElement(fills, f"{{{MAIN_NS}}}fill"), f"{{{MAIN_NS}}}patternFill", patternType="none")
    ET.SubElement(ET.SubElement(fills, f"{{{MAIN_NS}}}fill"), f"{{{MAIN_NS}}}patternFill", patternType="gray125")
    for rgb, _fill_id in sorted(fill_styles.items(), key=lambda item: item[1]):
        fill = ET.SubElement(fills, f"{{{MAIN_NS}}}fill")
        pattern = ET.SubElement(fill, f"{{{MAIN_NS}}}patternFill", patternType="solid")
        ET.SubElement(pattern, f"{{{MAIN_NS}}}fgColor", rgb=rgb)
        ET.SubElement(pattern, f"{{{MAIN_NS}}}bgColor", indexed="64")

    borders = ET.SubElement(root, f"{{{MAIN_NS}}}borders", count="1")
    ET.SubElement(borders, f"{{{MAIN_NS}}}border")
    ET.SubElement(root, f"{{{MAIN_NS}}}cellStyleXfs", count="1").append(ET.Element(f"{{{MAIN_NS}}}xf", numFmtId="0", fontId="0", fillId="0", borderId="0"))

    cell_xfs = ET.SubElement(root, f"{{{MAIN_NS}}}cellXfs", count=str(1 + len(fill_styles)))
    ET.SubElement(cell_xfs, f"{{{MAIN_NS}}}xf", numFmtId="0", fontId="0", fillId="0", borderId="0", xfId="0")
    for _rgb, fill_id in sorted(fill_styles.items(), key=lambda item: item[1]):
        ET.SubElement(cell_xfs, f"{{{MAIN_NS}}}xf", numFmtId="0", fontId="0", fillId=str(fill_id), borderId="0", xfId="0", applyFill="1")

    cell_styles = ET.SubElement(root, f"{{{MAIN_NS}}}cellStyles", count="1")
    ET.SubElement(cell_styles, f"{{{MAIN_NS}}}cellStyle", name="Normal", xfId="0", builtinId="0")
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def _render_sheet(sheet: SheetData, fill_styles: dict[str, int]) -> bytes:
    root = ET.Element(f"{{{MAIN_NS}}}worksheet")
    if sheet.max_row and sheet.max_column:
        ET.SubElement(root, f"{{{MAIN_NS}}}dimension", ref=f"A1:{_column_letters(sheet.max_column)}{sheet.max_row}")
    sheet_data = ET.SubElement(root, f"{{{MAIN_NS}}}sheetData")

    rows: dict[int, list[tuple[int, CellData]]] = {}
    for (row_idx, col_idx), cell in sorted(sheet.cells.items()):
        rows.setdefault(row_idx, []).append((col_idx, cell))

    for row_idx in sorted(rows):
        row_el = ET.SubElement(sheet_data, f"{{{MAIN_NS}}}row", r=str(row_idx))
        for col_idx, cell in sorted(rows[row_idx], key=lambda item: item[0]):
            if cell.value is None:
                continue
            attrs = {"r": f"{_column_letters(col_idx)}{row_idx}"}
            style_id = 0
            if cell.fill_rgb and cell.fill_rgb in fill_styles:
                style_id = fill_styles[cell.fill_rgb] - 1
            if style_id:
                attrs["s"] = str(style_id)
            if isinstance(cell.value, str):
                attrs["t"] = "inlineStr"
                cell_el = ET.SubElement(row_el, f"{{{MAIN_NS}}}c", attrs)
                is_el = ET.SubElement(cell_el, f"{{{MAIN_NS}}}is")
                t_el = ET.SubElement(is_el, f"{{{MAIN_NS}}}t")
                if cell.value.strip() != cell.value:
                    t_el.set(f"{{{XML_NS}}}space", "preserve")
                t_el.text = cell.value
            else:
                cell_el = ET.SubElement(row_el, f"{{{MAIN_NS}}}c", attrs)
                ET.SubElement(cell_el, f"{{{MAIN_NS}}}v").text = str(cell.value)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
