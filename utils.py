"""
utils.py
--------
Small, reusable helper functions shared across modules.
"""

import re


def to_int(value, default=0) -> int:
    try:
        s = str(value).strip()
        return default if (not s or s.startswith("#")) else int(float(s))
    except Exception:
        return default


def to_float(value, default=0.0) -> float:
    try:
        s = str(value).strip()
        return default if (not s or s.startswith("#")) else float(s)
    except Exception:
        return default


def split_cell(value) -> list[str]:
    """Split a pdfplumber cell that packs multiple values with \\n."""
    return [v.strip() for v in str(value or "").split("\n") if v.strip()]


def strip_lot_suffix(text: str) -> str:
    """Remove trailing '(xN)' from lot display strings."""
    return re.sub(r"\s*\(x\d+\)\s*$", "", str(text).strip())


def parse_pn(pn: str) -> tuple[str, int | None]:
    """
    Split a dash-revision PN into its base and revision number.
    '147712-03' → ('147712', 3)
    '150517'    → ('150517', None)
    """
    m = re.match(r"^(.+?)-(\d+)$", pn.strip())
    return (m.group(1), int(m.group(2))) if m else (pn.strip(), None)


def with_rev(base: str, rev: str) -> str:
    """
    Return 'base-rev' zero-padded to 2 digits e.g. '150517-02'.
    Returns bare base if rev is invalid.
    """
    try:
        return f"{base}-{int(str(rev).strip()):02d}"
    except Exception:
        return base
