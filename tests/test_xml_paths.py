"""No XML element or attribute of an example may vanish in a round-trip.

The other tests compare what openpyxl or xlsx2txt read back; this one looks
at the raw parts, so it also catches what neither of them reads (it found the
lost chart style and rounded corners). Every path found in an original part
must still exist somewhere in the same kind of part after import, unless it
is listed below as a known, harmless difference.
"""

import re
import zipfile
from collections import Counter
from pathlib import Path
from xml.etree import ElementTree

import pytest

from xlsx2txt import export_model
from xlsx2txt.importer import import_model

EXAMPLES = sorted((Path(__file__).resolve().parent.parent / "examples").glob("*.xlsx"))

# (part, path) patterns that may disappear, with the reason.
ALLOWED = [
    # An empty <left/> and a missing one both mean "no line".
    (r"xl/styles\.xml", r"/styleSheet/borders/border/(left|right|top|bottom|diagonal)"),
    # openpyxl writes an empty <workbookProtection/>; restored files leave it out.
    (r"xl/workbook\.xml", r"/workbook/workbookProtection"),
    # openpyxl writes a copy of a shared pivot cache per table; it is stored once.
    (r"xl/pivotCache/_rels/pivotCacheDefinitionN\.xml\.rels", r".*"),
]


def _local(tag):
    return tag.rsplit("}", 1)[-1]


def _paths(xml):
    found = Counter()

    def walk(element, prefix):
        path = f"{prefix}/{_local(element.tag)}"
        found[path] += 1
        for name in element.attrib:
            found[f"{path}@{_local(name)}"] += 1
        for child in element:
            walk(child, path)

    walk(ElementTree.fromstring(xml), "")
    return found


def _collect(path):
    parts = {}
    with zipfile.ZipFile(path) as archive:
        for name in archive.namelist():
            if name.endswith((".xml", ".rels", ".vml")):
                key = re.sub(r"\d+(?=\.xml(\.rels)?$)", "N", name)  # compare like parts
                parts.setdefault(key, Counter()).update(_paths(archive.read(name)))
    return parts


@pytest.mark.parametrize("path", EXAMPLES, ids=lambda p: p.stem)
def test_no_xml_path_vanishes(path, tmp_path):
    restored = tmp_path / path.name
    import_model(export_model(path), restored)
    original, copy = _collect(path), _collect(restored)
    vanished = sorted(
        (part, xml_path)
        for part, paths in original.items()
        for xml_path in paths
        if not copy.get(part, {}).get(xml_path)
        and not any(re.fullmatch(p, part) and re.fullmatch(x, xml_path) for p, x in ALLOWED)
    )
    assert vanished == []
