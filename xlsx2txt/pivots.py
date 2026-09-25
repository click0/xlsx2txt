"""Pivot tables: their definitions and caches are kept as indented XML files.

Layout in the model::

    sheet["pivotTables"] = [{"table": "pivotTable1.xml", "cache": "pivotCache1"}]
    model["pivots"] = {
        "pivotTable1.xml": "<pivotTableDefinition ...>",
        "pivotCache1.xml": "<pivotCacheDefinition ...>",
        "pivotCache1_records.xml": "<pivotCacheRecords ...>",   # optional
    }
"""

import re
import xml.etree.ElementTree as ET
from typing import Any

from openpyxl.pivot.cache import CacheDefinition
from openpyxl.pivot.record import RecordList
from openpyxl.pivot.table import TableDefinition
from openpyxl.xml.functions import fromstring, tostring

_NAMESPACES = {
    "": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "x14": "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main",
    "xr": "http://schemas.microsoft.com/office/spreadsheetml/2014/revision",
}
for _prefix, _uri in _NAMESPACES.items():
    ET.register_namespace(_prefix, _uri)


def pretty_xml(element) -> str:
    """Serialize an openpyxl tree as indented XML (one element per line)."""
    tree = ET.fromstring(tostring(element))
    ET.indent(tree, space="  ")
    return ET.tostring(tree, encoding="unicode") + "\n"


def export_pivots(worksheets) -> dict[str, Any]:
    """Collect pivot tables of all sheets.

    Returns ``{"sheets": {title: [entry, ...]}, "files": {name: xml}}``.
    Caches shared by several pivot tables are stored once.
    """
    files: dict[str, str] = {}
    sheets: dict[str, list[dict[str, str]]] = {}
    # openpyxl creates a separate cache object per sheet even when the file
    # has one shared cache, so identical caches are matched by content.
    cache_names: dict[tuple[str, str | None], str] = {}
    table_count = 0
    for ws in worksheets:
        for pivot in getattr(ws, "_pivots", []):
            table_count += 1
            entry = {"table": f"pivotTable{table_count}.xml"}
            files[entry["table"]] = pretty_xml(pivot.to_tree())
            cache = pivot.cache
            if cache is not None:
                definition = pretty_xml(cache.to_tree())
                records = pretty_xml(cache.records.to_tree()) if cache.records is not None else None
                key = (definition, records)
                if key not in cache_names:
                    name = f"pivotCache{len(cache_names) + 1}"
                    cache_names[key] = name
                    files[f"{name}.xml"] = definition
                    if records is not None:
                        files[f"{name}_records.xml"] = records
                entry["cache"] = cache_names[key]
            sheets.setdefault(ws.title, []).append(entry)
    return {"sheets": sheets, "files": files}


class PivotBuilder:
    """Recreates pivot tables on import, sharing caches between tables."""

    def __init__(self, files: dict[str, str]):
        self._files = files
        self._caches: dict[str, CacheDefinition] = {}

    def _read(self, name: str) -> str:
        if name not in self._files:
            raise ValueError(f"missing pivot file data/pivots/{name}")
        return self._files[name]

    def cache(self, name: str) -> CacheDefinition:
        if name not in self._caches:
            cache = CacheDefinition.from_tree(fromstring(self._read(f"{name}.xml")))
            records = self._files.get(f"{name}_records.xml")
            if records is not None:
                cache.records = RecordList.from_tree(fromstring(records))
            self._caches[name] = cache
        return self._caches[name]

    def table(self, entry: dict[str, str]) -> TableDefinition:
        pivot = TableDefinition.from_tree(fromstring(self._read(entry["table"])))
        if entry.get("cache"):
            pivot.cache = self.cache(entry["cache"])
        return pivot


def describe(xml: str) -> str | None:
    """'name at D1:E4' for summaries."""
    name = re.search(r'<pivotTableDefinition[^>]*\bname="([^"]*)"', xml)
    ref = re.search(r'<location[^>]*\bref="([^"]*)"', xml)
    if not name:
        return None
    return name.group(1) + (f" at {ref.group(1)}" if ref else "")
