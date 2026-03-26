"""Shared data models for the orchestrator."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any


@dataclass
class TableData:
    name: str
    headers: list[str]
    rows: list[list[Any]]


@dataclass
class ReportData:
    title: str
    fields: dict[str, Any]
    tables: list[TableData]
    metadata: dict[str, Any] = field(default_factory=dict)


@dataclass
class PlaceholderInfo:
    field_name: str
    location: str             # "paragraph" or "table_cell"
    element_index: int
    row: int | None = None
    column: int | None = None
    current_value: str = ""
    context: str = ""

    def to_dict(self) -> dict:
        d = {
            "field_name": self.field_name,
            "location": self.location,
            "element_index": self.element_index,
            "current_value": self.current_value,
            "context": self.context,
        }
        if self.row is not None:
            d["row"] = self.row
        if self.column is not None:
            d["column"] = self.column
        return d
