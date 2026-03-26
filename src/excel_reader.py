"""Read .xlsx files and produce ReportData — XML-aware extraction.

Uses the raw XML inside .xlsx (which is a zip of XML files) to:
- Resolve cell styles (bold, borders, fills) for header/section detection
- Detect merged cells for title areas
- Extract structured financial data with proper field names
"""

from __future__ import annotations

import io
import logging
import re
import zipfile
from datetime import date, datetime
from typing import Any
from xml.etree import ElementTree as ET

import openpyxl
from openpyxl.cell.cell import MergedCell

from src.models import ReportData, TableData

logger = logging.getLogger("excel_reader")

MAX_ROWS_PER_SHEET = 10_000

# OOXML namespaces
NS = {
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
}


# ── XML style parsing ─────────────────────────────────────────────────

class WorkbookStyles:
    """Parse xl/styles.xml to resolve cell formatting."""

    def __init__(self, styles_xml: bytes):
        root = ET.fromstring(styles_xml)

        # Parse fonts — extract bold flag
        self.fonts: list[dict] = []
        fonts_el = root.find("s:fonts", NS)
        if fonts_el is not None:
            for font_el in fonts_el.findall("s:font", NS):
                bold = font_el.find("s:b", NS) is not None
                sz_el = font_el.find("s:sz", NS)
                size = float(sz_el.get("val", "10")) if sz_el is not None else 10.0
                name_el = font_el.find("s:name", NS)
                name = name_el.get("val", "") if name_el is not None else ""
                self.fonts.append({"bold": bold, "size": size, "name": name})

        # Parse fills — extract background colors
        self.fills: list[str | None] = []
        fills_el = root.find("s:fills", NS)
        if fills_el is not None:
            for fill_el in fills_el.findall("s:fill", NS):
                pattern = fill_el.find("s:patternFill", NS)
                if pattern is not None:
                    fg = pattern.find("s:fgColor", NS)
                    if fg is not None and fg.get("rgb"):
                        self.fills.append(fg.get("rgb"))
                    else:
                        self.fills.append(None)
                else:
                    self.fills.append(None)

        # Parse borders — detect which have bottom borders (underlines / total lines)
        self.borders: list[dict] = []
        borders_el = root.find("s:borders", NS)
        if borders_el is not None:
            for border_el in borders_el.findall("s:border", NS):
                bottom = border_el.find("s:bottom", NS)
                top = border_el.find("s:top", NS)
                has_bottom = bottom is not None and bottom.get("style") is not None
                has_top = top is not None and top.get("style") is not None
                bottom_style = bottom.get("style", "") if bottom is not None else ""
                self.borders.append({
                    "has_bottom": has_bottom,
                    "has_top": has_top,
                    "bottom_style": bottom_style,
                })

        # Parse cellXfs — the style index lookup
        self.cell_xfs: list[dict] = []
        xfs_el = root.find("s:cellXfs", NS)
        if xfs_el is not None:
            for xf in xfs_el.findall("s:xf", NS):
                self.cell_xfs.append({
                    "font_id": int(xf.get("fontId", "0")),
                    "fill_id": int(xf.get("fillId", "0")),
                    "border_id": int(xf.get("borderId", "0")),
                    "num_fmt_id": int(xf.get("numFmtId", "0")),
                })

    def is_bold(self, style_index: int) -> bool:
        if style_index < 0 or style_index >= len(self.cell_xfs):
            return False
        font_id = self.cell_xfs[style_index]["font_id"]
        if font_id < len(self.fonts):
            return self.fonts[font_id]["bold"]
        return False

    def font_size(self, style_index: int) -> float:
        if style_index < 0 or style_index >= len(self.cell_xfs):
            return 10.0
        font_id = self.cell_xfs[style_index]["font_id"]
        if font_id < len(self.fonts):
            return self.fonts[font_id]["size"]
        return 10.0

    def has_fill(self, style_index: int) -> bool:
        if style_index < 0 or style_index >= len(self.cell_xfs):
            return False
        fill_id = self.cell_xfs[style_index]["fill_id"]
        if fill_id < len(self.fills):
            return self.fills[fill_id] is not None
        return False

    def has_bottom_border(self, style_index: int) -> bool:
        if style_index < 0 or style_index >= len(self.cell_xfs):
            return False
        border_id = self.cell_xfs[style_index]["border_id"]
        if border_id < len(self.borders):
            return self.borders[border_id]["has_bottom"]
        return False

    def bottom_border_style(self, style_index: int) -> str:
        if style_index < 0 or style_index >= len(self.cell_xfs):
            return ""
        border_id = self.cell_xfs[style_index]["border_id"]
        if border_id < len(self.borders):
            return self.borders[border_id]["bottom_style"]
        return ""


def _parse_shared_strings(ss_xml: bytes) -> list[str]:
    """Parse xl/sharedStrings.xml into a list of strings by index."""
    root = ET.fromstring(ss_xml)
    strings = []
    for si in root.findall("s:si", NS):
        # Concatenate all <t> elements (handles rich text with multiple <r> elements)
        parts = []
        for t in si.iter("{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t"):
            if t.text:
                parts.append(t.text)
        strings.append("".join(parts))
    return strings


def _parse_sheet_xml(
    sheet_xml: bytes,
    shared_strings: list[str],
    styles: WorkbookStyles,
) -> list[list[dict]]:
    """Parse a worksheet XML and return rows of cell dicts.

    Each cell dict has: value, style_index, is_bold, col_letter, row_num.
    """
    root = ET.fromstring(sheet_xml)
    rows: list[list[dict]] = []

    sheet_data = root.find("s:sheetData", NS)
    if sheet_data is None:
        return rows

    for row_el in sheet_data.findall("s:row", NS):
        cells = []
        for c_el in row_el.findall("s:c", NS):
            ref = c_el.get("r", "")
            style_idx = int(c_el.get("s", "0"))
            cell_type = c_el.get("t", "")
            v_el = c_el.find("s:v", NS)

            value: Any = None
            if v_el is not None and v_el.text is not None:
                if cell_type == "s":
                    # Shared string
                    idx = int(v_el.text)
                    value = shared_strings[idx] if idx < len(shared_strings) else ""
                elif cell_type == "b":
                    value = v_el.text == "1"
                else:
                    # Number or date
                    try:
                        value = float(v_el.text)
                        if value == int(value):
                            value = int(value)
                    except ValueError:
                        value = v_el.text

            cells.append({
                "ref": ref,
                "value": value,
                "style_index": style_idx,
                "is_bold": styles.is_bold(style_idx),
                "font_size": styles.font_size(style_idx),
                "has_fill": styles.has_fill(style_idx),
                "has_bottom_border": styles.has_bottom_border(style_idx),
                "bottom_border_style": styles.bottom_border_style(style_idx),
            })

        if cells:
            rows.append(cells)

    return rows


def _parse_merge_cells(sheet_xml: bytes) -> list[tuple[str, str]]:
    """Extract merged cell ranges from sheet XML."""
    root = ET.fromstring(sheet_xml)
    merges = []
    mc_el = root.find("s:mergeCells", NS)
    if mc_el is not None:
        for merge in mc_el.findall("s:mergeCell", NS):
            ref = merge.get("ref", "")
            if ":" in ref:
                start, end = ref.split(":")
                merges.append((start, end))
    return merges


# ── Value normalization ────────────────────────────────────────────────

def _normalize_value(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float):
        if value == int(value):
            return int(value)
        return value
    return value


def _is_label_like(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, str):
        s = value.strip()
        if not s or len(s) > 80:
            return False
        try:
            float(s.replace(",", ""))
            return False
        except ValueError:
            return True
    return False


# ── XML-aware extraction ──────────────────────────────────────────────

def _detect_header_row_xml(rows: list[list[dict]]) -> int:
    """Find the header row using XML style info.

    Priority:
    1. Row where most cells are bold (strong signal)
    2. Row with the most non-null cells (fallback)
    """
    if not rows:
        return 0

    scan_limit = min(15, len(rows))
    best_bold_idx = -1
    max_bold_count = 0
    best_count_idx = 0
    max_cell_count = 0

    for idx in range(scan_limit):
        bold_count = sum(1 for c in rows[idx] if c["is_bold"] and c["value"] is not None)
        cell_count = sum(1 for c in rows[idx] if c["value"] is not None)

        if bold_count > max_bold_count:
            max_bold_count = bold_count
            best_bold_idx = idx

        if cell_count > max_cell_count:
            max_cell_count = cell_count
            best_count_idx = idx

    # Prefer bold-based detection if we found bold cells
    if max_bold_count >= 2:
        logger.debug(f"    Header detected by bold formatting at row {best_bold_idx} ({max_bold_count} bold cells)")
        return best_bold_idx

    logger.debug(f"    Header detected by cell count at row {best_count_idx} ({max_cell_count} cells)")
    return best_count_idx


def _extract_title_from_xml(rows: list[list[dict]], header_idx: int, merges: list[tuple[str, str]]) -> str | None:
    """Extract title from preamble rows using formatting clues."""
    for idx in range(min(header_idx, 5)):
        row = rows[idx]
        for cell in row:
            if cell["value"] is None:
                continue
            val = str(cell["value"]).strip()
            if not val:
                continue
            # Bold + large font = likely title
            if cell["is_bold"] and cell["font_size"] >= 11:
                return val
            # First non-empty cell in row 0 is a reasonable fallback
            if idx == 0:
                return val
    return None


def _extract_preamble_fields_xml(rows: list[list[dict]], header_idx: int) -> dict[str, Any]:
    """Extract key-value pairs from rows above the header using XML info."""
    fields: dict[str, Any] = {}

    for row in rows[:header_idx]:
        values = [(c["value"], c["is_bold"]) for c in row if c["value"] is not None]
        if not values:
            continue

        # Pattern: "Label: Value" in a single cell
        for val, is_bold in values:
            text = str(val).strip()
            if ":" in text:
                parts = text.split(":", 1)
                label = parts[0].strip()
                value = parts[1].strip()
                if label and value and _is_label_like(label):
                    fields[label] = _normalize_value(value)

        # Pattern: Bold label cell followed by non-bold value cell
        all_cells = [c for c in row if c["value"] is not None]
        for i in range(len(all_cells) - 1):
            if all_cells[i]["is_bold"] and not all_cells[i + 1]["is_bold"]:
                label = str(all_cells[i]["value"]).strip()
                if _is_label_like(label):
                    fields[label] = _normalize_value(all_cells[i + 1]["value"])

    return fields


_TOTAL_KEYWORDS = re.compile(
    r"^(total|net|gross|sub.?total|balance|loss|profit|surplus|deficit|"
    r"cash and cash equivalents|share capital|accumulated|revenue|"
    r"operating|financing|investing)",
    re.IGNORECASE,
)


def _extract_summary_fields_xml(
    rows: list[list[dict]],
    header_idx: int,
    sheet_name: str,
) -> dict[str, Any]:
    """Extract fields from total/summary rows, detected by bold + bottom borders."""
    fields: dict[str, Any] = {}
    data_rows = rows[header_idx + 1:]

    for row in data_rows:
        # Find label cell (first non-empty string cell)
        label = None
        for c in row:
            if c["value"] is not None and isinstance(c["value"], str) and c["value"].strip():
                label = c["value"].strip()
                break

        if label is None:
            continue

        # A summary row is identified by: keyword match OR (bold label + bottom border on value cells)
        is_keyword = bool(_TOTAL_KEYWORDS.match(label))
        is_bold_label = any(c["is_bold"] for c in row if c["value"] is not None and isinstance(c["value"], str))
        has_border = any(c["has_bottom_border"] for c in row if isinstance(c["value"], (int, float)))

        if not is_keyword and not (is_bold_label and has_border):
            continue

        # Find the first numeric value
        for c in row:
            if isinstance(c["value"], (int, float)) and c["value"] != 0:
                key = f"{label} ({sheet_name})"
                fields[key] = _normalize_value(c["value"])
                break

    return fields


# ── Legacy fallback (openpyxl-based) ──────────────────────────────────

def _cell_value(cell) -> Any:
    if isinstance(cell, MergedCell):
        sheet = cell.parent
        for merge_range in sheet.merged_cells.ranges:
            if cell.coordinate in merge_range:
                top_left = sheet.cell(merge_range.min_row, merge_range.min_col)
                return top_left.value
        return None
    return cell.value


def _detect_key_value_sheet(rows: list[list[Any]]) -> dict[str, Any] | None:
    """If a sheet has 2 columns and column A looks like labels, return as key-value dict."""
    if not rows:
        return None
    for row in rows:
        non_empty = [v for v in row if v is not None]
        if len(non_empty) > 2:
            return None
    label_count = sum(1 for row in rows if len(row) >= 1 and _is_label_like(row[0]))
    if label_count < len(rows) * 0.6:
        return None
    fields = {}
    for row in rows:
        if len(row) >= 2 and _is_label_like(row[0]):
            key = str(row[0]).strip().rstrip(":")
            value = _normalize_value(row[1])
            if value is not None:
                fields[key] = value
    return fields if fields else None


# ── Main extraction ───────────────────────────────────────────────────

def extract_report_data(file_bytes: bytes, filename: str = "uploaded.xlsx") -> ReportData:
    """Read an .xlsx file using XML-aware parsing and produce a ReportData instance."""

    # Parse XML from the zip
    styles = None
    shared_strings: list[str] = []
    sheet_xmls: dict[str, bytes] = {}
    sheet_names: list[str] = []

    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            # Parse styles
            if "xl/styles.xml" in zf.namelist():
                styles = WorkbookStyles(zf.read("xl/styles.xml"))
                logger.debug(f"  Parsed styles: {len(styles.fonts)} fonts, {len(styles.cell_xfs)} cell formats")

            # Parse shared strings
            if "xl/sharedStrings.xml" in zf.namelist():
                shared_strings = _parse_shared_strings(zf.read("xl/sharedStrings.xml"))
                logger.debug(f"  Parsed {len(shared_strings)} shared strings")

            # Parse workbook for sheet names
            if "xl/workbook.xml" in zf.namelist():
                wb_root = ET.fromstring(zf.read("xl/workbook.xml"))
                sheets_el = wb_root.find("s:sheets", NS)
                if sheets_el is not None:
                    for s_el in sheets_el.findall("s:sheet", NS):
                        sheet_names.append(s_el.get("name", ""))

            # Read sheet XMLs
            for i, name in enumerate(sheet_names):
                sheet_path = f"xl/worksheets/sheet{i + 1}.xml"
                if sheet_path in zf.namelist():
                    sheet_xmls[name] = zf.read(sheet_path)

    except zipfile.BadZipFile:
        logger.warning("Not a valid zip/xlsx — falling back to openpyxl")
        return _extract_with_openpyxl(file_bytes, filename)

    if not styles or not sheet_xmls:
        logger.warning("Could not parse XML — falling back to openpyxl")
        return _extract_with_openpyxl(file_bytes, filename)

    # Process each sheet with XML-aware parsing
    fields: dict[str, Any] = {}
    tables: list[TableData] = []
    total_rows = 0
    title = filename.rsplit(".", 1)[0] if "." in filename else filename

    for sheet_name in sheet_names:
        if sheet_name not in sheet_xmls:
            continue

        logger.debug(f"  Processing sheet (XML): '{sheet_name}'")
        sheet_xml = sheet_xmls[sheet_name]

        # Parse rows with style info
        xml_rows = _parse_sheet_xml(sheet_xml, shared_strings, styles)
        if not xml_rows:
            logger.debug(f"    Sheet '{sheet_name}': empty, skipping")
            continue

        total_rows += len(xml_rows)
        logger.debug(f"    Sheet '{sheet_name}': {len(xml_rows)} rows")

        # Parse merge cells
        merges = _parse_merge_cells(sheet_xml)

        # Try to detect simple 2-column key-value sheets first (use plain values)
        plain_rows = [[c["value"] for c in row] for row in xml_rows]
        kv = _detect_key_value_sheet(plain_rows)
        if kv:
            logger.debug(f"    Sheet '{sheet_name}' detected as key-value ({len(kv)} pairs)")
            for k, v in kv.items():
                logger.debug(f"      {k}: {v!r}")
            fields.update(kv)
            continue

        # Detect header row using XML styles
        header_idx = _detect_header_row_xml(xml_rows)
        logger.debug(f"    Sheet '{sheet_name}': header at row index {header_idx}")

        # Extract title from first sheet preamble
        if not tables and not fields:
            extracted_title = _extract_title_from_xml(xml_rows, header_idx, merges)
            if extracted_title:
                title = extracted_title

        # Extract preamble key-value fields
        preamble = _extract_preamble_fields_xml(xml_rows, header_idx)
        if preamble:
            logger.debug(f"    Sheet '{sheet_name}': {len(preamble)} preamble fields")
            for k, v in preamble.items():
                logger.debug(f"      {k}: {v!r}")
            fields.update(preamble)

        # Extract summary/total row fields
        summary = _extract_summary_fields_xml(xml_rows, header_idx, sheet_name)
        if summary:
            logger.debug(f"    Sheet '{sheet_name}': {len(summary)} summary fields")
            for k, v in summary.items():
                logger.debug(f"      {k}: {v!r}")
            fields.update(summary)

        # Build table data
        header_cells = xml_rows[header_idx]
        headers = []
        for i, c in enumerate(header_cells):
            if c["value"] is not None and isinstance(c["value"], str):
                headers.append(c["value"].strip())
            else:
                headers.append(f"Column {i + 1}")

        data_rows_plain = []
        for row in xml_rows[header_idx + 1:]:
            data_rows_plain.append([_normalize_value(c["value"]) for c in row])

        if data_rows_plain:
            tables.append(TableData(name=sheet_name, headers=headers, rows=data_rows_plain))

    metadata: dict[str, Any] = {
        "source_file": filename,
        "sheet_names": sheet_names,
        "total_rows": total_rows,
        "parse_method": "xml",
    }

    return ReportData(title=title, fields=fields, tables=tables, metadata=metadata)


def _extract_with_openpyxl(file_bytes: bytes, filename: str) -> ReportData:
    """Fallback extraction using openpyxl (no XML style info)."""
    try:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True, read_only=True)
    except Exception as e:
        raise ValueError(f"Cannot read Excel file: {e}")

    fields: dict[str, Any] = {}
    tables: list[TableData] = []
    total_rows = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        all_rows: list[list[Any]] = []
        row_count = 0
        for row in ws.iter_rows():
            if row_count >= MAX_ROWS_PER_SHEET:
                break
            values = [_cell_value(cell) for cell in row]
            if any(v is not None for v in values):
                all_rows.append([_normalize_value(v) for v in values])
            row_count += 1

        if not all_rows:
            continue
        total_rows += len(all_rows)

        kv = _detect_key_value_sheet(all_rows)
        if kv:
            fields.update(kv)
            continue

        first_row = all_rows[0]
        all_strings = all(isinstance(v, str) for v in first_row if v is not None)
        if all_strings and len(all_rows) > 1:
            headers = [str(v) if v is not None else f"Column {i+1}" for i, v in enumerate(first_row)]
            data_rows = all_rows[1:]
        else:
            headers = [f"Column {i+1}" for i in range(len(first_row))]
            data_rows = all_rows
        tables.append(TableData(name=sheet_name, headers=headers, rows=data_rows))

    wb.close()
    metadata = {
        "source_file": filename,
        "sheet_names": list(wb.sheetnames),
        "total_rows": total_rows,
        "parse_method": "openpyxl_fallback",
    }
    title = filename.rsplit(".", 1)[0] if "." in filename else filename
    return ReportData(title=title, fields=fields, tables=tables, metadata=metadata)
