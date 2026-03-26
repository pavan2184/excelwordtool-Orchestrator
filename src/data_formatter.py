"""Convert ReportData into structured plain text for the word-tool."""

from __future__ import annotations

from datetime import date, datetime
from typing import Any

from src.models import ReportData, TableData


def _format_value(value: Any) -> str:
    """Format a single value for display."""
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.strftime("%Y-%m-%d")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float):
        return f"{value:g}"
    return str(value)


def format_report_data_as_text(report: ReportData) -> str:
    """Convert a ReportData object into structured plain text."""
    source = report.metadata.get("source_file", report.title)
    lines: list[str] = [f"SOURCE DATA (extracted from: {source})", ""]

    if report.fields:
        lines.append("FIELDS:")
        for key, value in report.fields.items():
            lines.append(f"- {key}: {_format_value(value)}")
        lines.append("")

    for table in report.tables:
        if not table.headers and not table.rows:
            continue

        lines.append(f"TABLE: {table.name}")

        all_rows = [table.headers] + table.rows if table.headers else table.rows
        if not all_rows:
            continue

        num_cols = max(len(row) for row in all_rows)
        col_widths = [0] * num_cols
        formatted_rows: list[list[str]] = []

        for row in all_rows:
            formatted = [_format_value(row[i]) if i < len(row) else "" for i in range(num_cols)]
            formatted_rows.append(formatted)
            for i, cell in enumerate(formatted):
                col_widths[i] = max(col_widths[i], len(cell))

        for formatted in formatted_rows:
            cells = [cell.ljust(col_widths[i]) for i, cell in enumerate(formatted)]
            lines.append("| " + " | ".join(cells) + " |")

        lines.append("")

    text = "\n".join(lines).strip()
    if not report.fields and not report.tables:
        return f"SOURCE DATA (extracted from: {source})\n\nNo data found."
    return text


def report_data_to_dict(report: ReportData) -> dict:
    """Serialize ReportData to a JSON-safe dict."""
    return {
        "title": report.title,
        "fields": report.fields,
        "tables": [
            {"name": t.name, "headers": t.headers, "rows": t.rows}
            for t in report.tables
        ],
        "metadata": report.metadata,
    }


def report_data_from_dict(data: dict) -> ReportData:
    """Deserialize a JSON dict into a ReportData object."""
    if not isinstance(data, dict):
        raise ValueError("ReportData must be a JSON object")

    title = data.get("title", "")
    fields = data.get("fields", {})
    if not isinstance(fields, dict):
        raise ValueError("'fields' must be a dict")

    raw_tables = data.get("tables", [])
    if not isinstance(raw_tables, list):
        raise ValueError("'tables' must be a list")

    tables: list[TableData] = []
    for i, t in enumerate(raw_tables):
        if not isinstance(t, dict):
            raise ValueError(f"tables[{i}] must be a dict")
        tables.append(TableData(
            name=str(t.get("name", f"Table {i + 1}")),
            headers=t.get("headers", []),
            rows=t.get("rows", []),
        ))

    metadata = data.get("metadata", {})
    if not isinstance(metadata, dict):
        metadata = {}

    return ReportData(title=title, fields=fields, tables=tables, metadata=metadata)
