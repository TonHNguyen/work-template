"""
config.py
---------
All user-configurable settings in one place.
Change paths, anchor PN, rev-sensitive parts, or PN overrides here — nowhere else.
"""

import os
from openpyxl.styles import PatternFill, Font

# All paths are relative to THIS file's directory so the app works
# regardless of which folder you run it from.
_HERE = os.path.dirname(os.path.abspath(__file__))

def _path(*parts):
    return os.path.join(_HERE, *parts)

# ── File paths ────────────────────────────────────────────────────────────────
TEMPLATE_PATH = _path("templates", "DOC-012743.pdf")
EXPORT_PATH   = _path("data", "export_all_c80eb8e5ef3040379d61be5c46e8cd83.xlsx")
OUTPUT_DIR    = _path("outputs")

# ── Anchor PN ─────────────────────────────────────────────────────────────────
ANCHOR_PN = "LBL-F5-01"

# ── Rev-sensitive parts ───────────────────────────────────────────────────────
# Only these base PNs require an exact revision match.
# All other parts will use the highest available revision in the export.
REV_SENSITIVE = {"147712"}

# ── Manual PN overrides ───────────────────────────────────────────────────────
# Parts not yet in the PDF but valid replacements in the export.
# Format:  "pdf_pn" -> ["replacement_pn_1", ...]
# Delete a line once the PDF is updated.
PN_OVERRIDES = {
    "159604-01": ["160275-02"],   # PROPELLER ASSY, T-MOTOR, 4-PLY, CENTER, CCW
    "159603-01": ["160274-02"],   # PROPELLER ASSY, T-MOTOR, 4-PLY, CENTER, CW
    "159602-01": ["160273-02"],   # PROPELLER ASSY, T-MOTOR, 4-PLY, OFF-CENTER, CCW
    "159601-01": ["160272-02"],   # PROPELLER ASSY, T-MOTOR, 4-PLY, OFF-CENTER, CW
}

# ── Cell highlight colours (Missing Components sheet) ─────────────────────────
FILL_AMBER = PatternFill("solid", fgColor="FFD966")   # OLD REV
FILL_RED   = PatternFill("solid", fgColor="FF7575")   # NOT FOUND
FONT_BOLD  = Font(bold=True)