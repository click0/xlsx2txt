"""Extraction of VBA source code for review (requires the optional oletools)."""

from pathlib import Path
from typing import Dict, List, Tuple, Union

from xlsx2txt.names import safe_file_name


def extract_vba_sources(path: Union[str, Path]) -> Tuple[Dict[str, str], List[str]]:
    """Return ``({module file name: source code}, warnings)``.

    The sources are informational: importing always restores the original
    ``vbaProject.bin``, so macros must still be edited in Excel.
    """
    try:
        from oletools.olevba import VBA_Parser
    except ImportError:
        return {}, ["VBA source code is not exported: install oletools (pip install xlsx2txt[vba])"]

    sources: Dict[str, str] = {}
    parser = VBA_Parser(str(path))
    try:
        if not parser.detect_vba_macros():
            return {}, []
        for _, _, vba_filename, code in parser.extract_macros():
            if isinstance(code, bytes):
                code = code.decode("utf-8", errors="replace")
            name = safe_file_name(vba_filename or "module.bas")
            base, dot, ext = name.rpartition(".")
            candidate, counter = name, 2
            while candidate.lower() in {n.lower() for n in sources}:
                candidate = f"{base}_{counter}.{ext}" if dot else f"{name}_{counter}"
                counter += 1
            sources[candidate] = code.replace("\r\n", "\n")
    except Exception as exc:  # oletools raises many parser-specific errors
        return {}, [f"VBA source code could not be extracted: {exc}"]
    finally:
        parser.close()
    return sources, []
