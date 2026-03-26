"""Scan a .docx template for placeholder fields using XML-level parsing.

Parses word/document.xml directly to find:
- XXX+ placeholder patterns at the run level (preserving formatting context)
- Traditional placeholder patterns ([●], ___, [TBD], etc.)
- Empty table cells adjacent to labels
- Suggests fuzzy mapping from placeholders to ReportData fields
"""

from __future__ import annotations

import io
import logging
import re
import zipfile
from typing import Any
from xml.etree import ElementTree as ET

from docx import Document

from src.models import PlaceholderInfo, ReportData

logger = logging.getLogger("template_scanner")

# OOXML namespaces
WML = "http://schemas.microsoft.com/office/word/2012/wordml"
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W = f"{{{W_NS}}}"

# Placeholder patterns
PLACEHOLDER_PATTERNS = [
    re.compile(r'\[●\]'),
    re.compile(r'\[[^\]]{1,50}\]'),       # [Name], [date], [TBD], etc.
    re.compile(r'_{3,}'),                  # ___ underlines
    re.compile(r'DD[/\-]MM[/\-]YYYY', re.IGNORECASE),
    re.compile(r'\[TBD\]', re.IGNORECASE),
]

# XXX+ pattern — 4 or more consecutive X's
XXX_PATTERN = re.compile(r'X{4,}')


def _is_placeholder(text: str) -> bool:
    """Check if text contains a placeholder pattern."""
    if XXX_PATTERN.search(text):
        return True
    return any(p.search(text) for p in PLACEHOLDER_PATTERNS)


def _extract_label_from_context(full_text: str, placeholder_match: str) -> str:
    """Derive a field label from text surrounding a placeholder."""
    idx = full_text.find(placeholder_match)
    if idx > 0:
        before = full_text[:idx].strip().rstrip(":").strip()
        if before:
            return before
    cleaned = full_text.replace(placeholder_match, "").strip().rstrip(":").strip()
    return cleaned if cleaned else placeholder_match


# ── XML-level scanning ────────────────────────────────────────────────

def _scan_xml_for_placeholders(doc_xml: bytes) -> list[dict]:
    """Parse word/document.xml to find placeholder runs with formatting context.

    Returns a list of dicts with: text, element_index, context, run_bold, font_name, font_size.
    """
    root = ET.fromstring(doc_xml)
    body = root.find(f"{W}body")
    if body is None:
        return []

    placeholders = []
    element_index = 0

    for child in body:
        element_index += 1
        tag = child.tag.split("}")[-1] if "}" in child.tag else child.tag

        if tag == "p":
            # It's a paragraph — collect all run texts and check for placeholders
            runs = child.findall(f".//{W}r")
            full_text_parts = []
            run_details = []

            for run in runs:
                t_el = run.find(f"{W}t")
                if t_el is not None and t_el.text:
                    text = t_el.text
                    full_text_parts.append(text)

                    # Extract run formatting from rPr
                    rpr = run.find(f"{W}rPr")
                    is_bold = False
                    font_name = ""
                    font_size = ""

                    if rpr is not None:
                        if rpr.find(f"{W}b") is not None:
                            is_bold = True
                        fonts = rpr.find(f"{W}rFonts")
                        if fonts is not None:
                            font_name = fonts.get(f"{W}ascii", fonts.get("w:ascii", ""))
                        sz = rpr.find(f"{W}sz")
                        if sz is not None:
                            font_size = sz.get(f"{W}val", sz.get("val", ""))

                    run_details.append({
                        "text": text,
                        "bold": is_bold,
                        "font": font_name,
                        "size": font_size,
                    })

            full_text = "".join(full_text_parts)

            # Check each run for XXX+ placeholder pattern
            for rd in run_details:
                if XXX_PATTERN.search(rd["text"]):
                    placeholders.append({
                        "text": rd["text"].strip(),
                        "element_index": element_index,
                        "location": "paragraph",
                        "context": full_text[:120],
                        "bold": rd["bold"],
                        "font": rd["font"],
                        "size": rd["size"],
                        "pattern": "XXX",
                    })

            # Also check full paragraph text for other patterns
            for pattern in PLACEHOLDER_PATTERNS:
                for match in pattern.finditer(full_text):
                    placeholders.append({
                        "text": match.group(),
                        "element_index": element_index,
                        "location": "paragraph",
                        "context": full_text[:120],
                        "bold": False,
                        "font": "",
                        "size": "",
                        "pattern": "standard",
                    })

        elif tag == "tbl":
            # It's a table — scan cells
            for ri, tr in enumerate(child.findall(f".//{W}tr"), 1):
                cells_text = []
                cells_runs = []
                for tc in tr.findall(f"{W}tc"):
                    # Collect all text in the cell
                    cell_parts = []
                    cell_has_placeholder = False
                    for t_el in tc.iter(f"{W}t"):
                        if t_el.text:
                            cell_parts.append(t_el.text)
                            if XXX_PATTERN.search(t_el.text) or any(p.search(t_el.text) for p in PLACEHOLDER_PATTERNS):
                                cell_has_placeholder = True
                    cell_text = "".join(cell_parts).strip()
                    cells_text.append(cell_text)
                    cells_runs.append(cell_has_placeholder)

                # Check for placeholder cells and empty cells with adjacent labels
                for ci, (cell_text, has_ph) in enumerate(zip(cells_text, cells_runs)):
                    if has_ph and cell_text:
                        placeholders.append({
                            "text": cell_text,
                            "element_index": element_index,
                            "location": "table_cell",
                            "row": ri,
                            "column": ci + 1,
                            "context": f"Row {ri}: {' | '.join(cells_text[:4])}",
                            "bold": False,
                            "font": "",
                            "size": "",
                            "pattern": "XXX" if XXX_PATTERN.search(cell_text) else "standard",
                        })
                    elif not cell_text and ci > 0 and cells_text[ci - 1]:
                        # Empty cell next to a label — potential placeholder
                        label = cells_text[ci - 1].rstrip(":").strip()
                        if label and _is_label_like_text(label):
                            placeholders.append({
                                "text": "(empty)",
                                "element_index": element_index,
                                "location": "table_cell",
                                "row": ri,
                                "column": ci + 1,
                                "context": f"Row {ri}: {' | '.join(cells_text[:4])}",
                                "bold": False,
                                "font": "",
                                "size": "",
                                "pattern": "empty_cell",
                                "inferred_label": label,
                            })

    return placeholders


def _is_label_like_text(text: str) -> bool:
    """Check if text looks like a field label."""
    s = text.strip()
    if not s or len(s) > 80:
        return False
    try:
        float(s.replace(",", ""))
        return False
    except ValueError:
        return True


# ── Public API ────────────────────────────────────────────────────────

def scan_template(file_bytes: bytes) -> list[PlaceholderInfo]:
    """Scan a .docx template using XML-level parsing and return all detected placeholders."""
    placeholders: list[PlaceholderInfo] = []
    seen_fields: set[str] = set()

    # Try XML-level scanning first
    xml_results = []
    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            if "word/document.xml" in zf.namelist():
                doc_xml = zf.read("word/document.xml")
                xml_results = _scan_xml_for_placeholders(doc_xml)
                logger.info(f"XML scan found {len(xml_results)} raw placeholder hits")
    except (zipfile.BadZipFile, ET.ParseError) as e:
        logger.warning(f"XML parsing failed, falling back to python-docx: {e}")

    if xml_results:
        # Deduplicate and build PlaceholderInfo list
        for hit in xml_results:
            # Determine field name
            if "inferred_label" in hit:
                label = hit["inferred_label"]
            else:
                label = _extract_label_from_context(hit["context"], hit["text"])

            # Skip if we've already seen this label
            if label in seen_fields:
                continue
            seen_fields.add(label)

            ph = PlaceholderInfo(
                field_name=label,
                location=hit["location"],
                element_index=hit["element_index"],
                row=hit.get("row"),
                column=hit.get("column"),
                current_value=hit["text"],
                context=hit["context"],
            )
            placeholders.append(ph)

        logger.info(f"Found {len(placeholders)} unique placeholders via XML parsing")
        return placeholders

    # Fallback to python-docx based scanning
    return _scan_with_python_docx(file_bytes)


def _scan_with_python_docx(file_bytes: bytes) -> list[PlaceholderInfo]:
    """Fallback scanner using python-docx (same as original implementation)."""
    doc = Document(io.BytesIO(file_bytes))
    placeholders: list[PlaceholderInfo] = []
    seen_fields: set[str] = set()

    element_index = 0
    for kind, obj in _walk_body(doc):
        element_index += 1

        if kind == "paragraph":
            text = obj.text.strip()
            if not text:
                continue

            all_patterns = PLACEHOLDER_PATTERNS + [XXX_PATTERN]
            for pattern in all_patterns:
                for match in pattern.finditer(text):
                    matched = match.group()
                    label = _extract_label_from_context(text, matched)
                    if label in seen_fields:
                        continue
                    seen_fields.add(label)
                    placeholders.append(PlaceholderInfo(
                        field_name=label,
                        location="paragraph",
                        element_index=element_index,
                        current_value=matched,
                        context=text[:100],
                    ))

        elif kind == "table":
            for ri, row in enumerate(obj.rows, 1):
                cells = [cell.text.strip() for cell in row.cells]
                for ci, cell_text in enumerate(cells, 1):
                    if not cell_text:
                        if ci == 2 and len(cells) >= 2 and cells[0] and not _is_placeholder(cells[0]):
                            label = cells[0].rstrip(":").strip()
                            if label and label not in seen_fields:
                                seen_fields.add(label)
                                placeholders.append(PlaceholderInfo(
                                    field_name=label,
                                    location="table_cell",
                                    element_index=element_index,
                                    row=ri, column=ci,
                                    current_value="(empty)",
                                    context=f"Row {ri}: {' | '.join(cells[:3])}",
                                ))
                        continue

                    if _is_placeholder(cell_text):
                        if ci > 1 and cells[ci - 2]:
                            label = cells[ci - 2].rstrip(":").strip()
                        else:
                            label = _extract_label_from_context(cell_text, cell_text)
                        if label in seen_fields:
                            continue
                        seen_fields.add(label)
                        placeholders.append(PlaceholderInfo(
                            field_name=label,
                            location="table_cell",
                            element_index=element_index,
                            row=ri, column=ci,
                            current_value=cell_text,
                            context=f"Row {ri}: {' | '.join(cells[:3])}",
                        ))

    logger.info(f"Found {len(placeholders)} placeholders via python-docx fallback")
    return placeholders


def _walk_body(doc: Document):
    """Walk document body yielding (type, object) in order."""
    from docx.table import Table as DocxTable
    from docx.text.paragraph import Paragraph

    para_map = {p._element: p for p in doc.paragraphs}
    table_map = {t._element: t for t in doc.tables}

    for child in doc.element.body:
        if child in para_map:
            yield ("paragraph", para_map[child])
        elif child in table_map:
            yield ("table", table_map[child])


# ── Mapping ───────────────────────────────────────────────────────────

def _normalize_key(s: str) -> str:
    return re.sub(r'[^a-z0-9]', '', s.lower())


def suggest_mapping(
    placeholders: list[PlaceholderInfo],
    report: ReportData,
) -> dict[str, Any]:
    """Fuzzy-match template placeholders to ReportData fields."""
    normalized_fields: dict[str, tuple[str, Any]] = {}
    for key, value in report.fields.items():
        norm = _normalize_key(key)
        normalized_fields[norm] = (key, value)

    mapping: dict[str, Any] = {}

    for ph in placeholders:
        norm_ph = _normalize_key(ph.field_name)

        # Exact normalized match
        if norm_ph in normalized_fields:
            mapping[ph.field_name] = normalized_fields[norm_ph][1]
            continue

        # Substring match
        matched = False
        for norm_key, (orig_key, value) in normalized_fields.items():
            if norm_ph in norm_key or norm_key in norm_ph:
                mapping[ph.field_name] = value
                matched = True
                break

        if not matched:
            mapping[ph.field_name] = None

    return mapping
