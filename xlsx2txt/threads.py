"""Threaded comments (Excel 365 conversations) and their authors.

openpyxl keeps only the legacy note Excel writes next to each thread, so the
threads (``xl/threadedComments/*.xml``) and the people who wrote them
(``xl/persons/person.xml``) are read and written here. They are stored as
JSON: one record per comment with the cell, author, date and text.
"""

import html
import re
from typing import Any

from xlsx2txt.xmlfrag import self_contained, split_children

THREADS_NS = "http://schemas.microsoft.com/office/spreadsheetml/2018/threadedcomments"
MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
THREAD_REL = "http://schemas.microsoft.com/office/2017/10/relationships/threadedComment"
PERSON_REL = "http://schemas.microsoft.com/office/2017/10/relationships/person"
THREAD_TYPE = "application/vnd.ms-excel.threadedcomments+xml"
PERSON_TYPE = "application/vnd.ms-excel.person+xml"

_ATTR = re.compile(r"([\w:.-]+)=(?:\"([^\"]*)\"|'([^']*)')")
_THREAD_ORDER = ["ref", "id", "parentId", "personId", "dT", "done"]


def _attrs(tag: str) -> dict[str, str]:
    return {name: html.unescape(a if a or not b else b)
            for name, a, b in _ATTR.findall(tag) if not name.startswith("xmlns")}


def _start_tag(fragment: str) -> str:
    return re.match(r"<[^>]*>", fragment).group(0)


def read_threads(xml: str) -> list[dict[str, Any]]:
    """Comments of one sheet, in file order."""
    root, children = split_children(xml)
    result = []
    for child in children:
        if not re.match(r"<(?:\w+:)?threadedComment\b", child):
            continue
        attrs = _attrs(_start_tag(child))
        record = {key: attrs.pop(key) for key in _THREAD_ORDER if key in attrs}
        record.update(attrs)
        _, parts = split_children(child) if not child.endswith("/>") else ("", [])
        extra = []
        for part in parts:
            text = re.fullmatch(r"<(?:\w+:)?text>(.*)</(?:\w+:)?text>", part, re.S)
            if text:
                record["text"] = html.unescape(text.group(1))
            elif part.startswith("<") and not re.match(r"<(?:\w+:)?text\s*/>", part):
                extra.append(self_contained(part, root + _start_tag(child)))
            else:
                record["text"] = ""
        if extra:
            record["xml"] = "".join(extra)  # mentions and other details, kept as is
        result.append(record)
    return result


def write_threads(threads: list[dict[str, Any]]) -> str:
    items = []
    for record in threads:
        attrs = "".join(f' {key}="{html.escape(str(value))}"' for key, value in record.items()
                        if key not in ("text", "xml"))
        text = html.escape(record.get("text", ""), quote=False)
        items.append(f"<threadedComment{attrs}><text>{text}</text>{record.get('xml', '')}</threadedComment>")
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<ThreadedComments xmlns="{THREADS_NS}" xmlns:x="{MAIN_NS}">{"".join(items)}</ThreadedComments>')


def read_persons(xml: str) -> list[dict[str, str]]:
    _, children = split_children(xml)
    return [_attrs(_start_tag(child)) for child in children if re.match(r"<(?:\w+:)?person\b", child)]


def write_persons(persons: list[dict[str, str]]) -> str:
    items = "".join(
        "<person" + "".join(f' {key}="{html.escape(str(value))}"' for key, value in person.items()) + "/>"
        for person in persons
    )
    return ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            f'<personList xmlns="{THREADS_NS}" xmlns:x="{MAIN_NS}">{items}</personList>')


def person_names(persons: list[dict[str, str]]) -> dict[str, str]:
    return {p.get("id", ""): p.get("displayName", "") for p in persons}
