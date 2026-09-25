"""Filesystem-safe file names."""

import re

_UNSAFE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_RESERVED = {"con", "prn", "aux", "nul"} | {f"com{i}" for i in range(1, 10)} | {f"lpt{i}" for i in range(1, 10)}


def safe_file_name(name: str, default: str = "sheet") -> str:
    """Replace characters that are invalid on common filesystems."""
    base = _UNSAFE.sub("_", name).strip(" .") or default
    stem = base.split(".", 1)[0]
    if stem.lower() in _RESERVED:
        base = f"_{base}"
    return base
