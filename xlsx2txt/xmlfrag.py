"""Work with fragments of XML parts as text, keeping them byte for byte.

Parts that openpyxl does not understand are kept as XML fragments cut out of
the original part; parsing and re-serializing them would change prefixes and
namespace declarations (and break ``mc:Ignorable``).
"""

import re

# Tags of an XML part (comments and processing instructions included so
# they can be skipped when tracking nesting).
_TAG = re.compile(r"<!--.*?-->|<\?.*?\?>|<!\[CDATA\[.*?\]\]>|<(/?)([\w.:-]+)((?:[^>\"']|\"[^\"]*\"|'[^']*')*?)(/?)>",
                  re.S)
_XMLNS = re.compile(r"\bxmlns(?::([\w.-]+))?=(\"[^\"]*\"|'[^']*')")


def split_children(xml: str) -> tuple[str, list[str]]:
    """Split an XML part into its root start tag and its top-level elements."""
    depth = 0
    root = ""
    start = None
    children = []
    for match in _TAG.finditer(xml):
        closing, name, _, self_closing = match.groups()
        if name is None:
            continue
        if closing:
            depth -= 1
            if depth == 1 and start is not None:
                children.append(xml[start:match.end()])
                start = None
        else:
            if depth == 0:
                root = match.group(0)
            elif depth == 1:
                start = match.start()
                if self_closing:
                    children.append(match.group(0))
                    start = None
            if not self_closing:
                depth += 1
    return root, children


def self_contained(fragment: str, root: str) -> str:
    """Declare on a fragment the namespaces it inherits from the root element."""
    first_tag = re.match(r"<[^>]*?(?=/?>)", fragment).group(0)
    declared = {prefix or "" for prefix, _ in _XMLNS.findall(first_tag)}
    used = set(re.findall(r"</?([\w.-]+):", fragment))
    used |= set(re.findall(r"\s([\w.-]+):[\w.-]+=", fragment)) - {"xmlns"}
    for value in re.findall(r"\bIgnorable=\"([^\"]*)\"", fragment):
        used |= set(value.split())
    if re.search(r"<(?![\w.-]+:)[\w.-]+[\s/>]", fragment):
        used.add("")
    additions = [
        f" xmlns:{prefix}={value}" if prefix else f" xmlns={value}"
        for prefix, value in _XMLNS.findall(root)
        if (prefix or "") in used - declared
    ]
    return first_tag + "".join(additions) + fragment[len(first_tag):]
