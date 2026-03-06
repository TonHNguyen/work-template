"""
matcher.py
----------
Loads the export workbook and matches PDF requirements against it.

Matching priority (non-rev-sensitive parts):
  1. Preferred versioned PN from PDF        e.g. '159604-01'
  2. Manual overrides from config           e.g. '160275-02'
  3. Highest revision of that base in the export
  4. Alternate PNs from PDF effectivity rows

Rev-sensitive parts (e.g. 147712):
  Exact match only — triggers OLD REV if a lower revision is found.
"""

import re
import pandas as pd
from config import ANCHOR_PN, REV_SENSITIVE, PN_OVERRIDES
from utils  import parse_pn, to_float, to_int


# =============================================================================
# EXPORT LOADING
# =============================================================================

def load_export(export_path: str) -> pd.DataFrame:
    df = pd.read_excel(export_path, dtype=str).fillna("")
    required = ["ProductNo", "Serial #", "Lot #",
                "Component Description", "Quantity", "Parent Serial #"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise SystemExit(f"ERROR: Export missing columns: {missing}")
    df["ProductNo"] = df["ProductNo"].str.strip()
    return df


def get_parent_serial(df: pd.DataFrame) -> str:
    rows = df[df["ProductNo"] == ANCHOR_PN]
    if rows.empty:
        raise ValueError(f"Anchor PN '{ANCHOR_PN}' not found in export.")
    parents = rows["Parent Serial #"].str.strip().replace("", pd.NA).dropna()
    if parents.empty:
        raise ValueError(f"Anchor PN '{ANCHOR_PN}' has no Parent Serial #.")
    return parents.value_counts().idxmax()


# =============================================================================
# MATCHING
# =============================================================================

def _pick_highest_rev(base: str, installed: dict) -> str:
    """Return the highest 'base-XX' variant present in the export."""
    pattern    = re.compile(r"^" + re.escape(base) + r"-(\d+)$")
    candidates = [(int(m.group(1)), pn)
                  for pn in installed
                  if (m := pattern.match(pn))]
    if not candidates:
        return base if base in installed else ""
    return max(candidates)[1]


def find_match(part: dict, installed: dict) -> str:
    """
    Return the best matching PN from the export, or '' if nothing found.

    Priority
    --------
    1. Preferred versioned PN from PDF    e.g. '159604-01'
    2. Manual overrides from PN_OVERRIDES e.g. '160275-02'
    3. Highest revision of base PN in export
    4. Alternate PNs from PDF effectivity rows

    Rev-sensitive parts: exact match only — no fallback.
    """
    base      = part["base"]
    preferred = part["preferred"]

    # Rev-sensitive: exact match only
    if base in REV_SENSITIVE:
        return preferred if preferred in installed else ""

    # 1. Preferred PN from PDF
    if preferred in installed:
        return preferred

    # 2. Manual overrides (config.py PN_OVERRIDES)
    for override_pn in PN_OVERRIDES.get(preferred, []):
        if override_pn in installed:
            return override_pn

    # 3. Highest revision of base PN in export
    highest = _pick_highest_rev(base, installed)
    if highest and highest in installed:
        return highest

    # 4. Alternate PNs from PDF effectivity rows
    for alt_base, alt_pref in part["alts"]:
        if alt_pref in installed:
            return alt_pref
        alt_highest = _pick_highest_rev(alt_base, installed)
        if alt_highest and alt_highest in installed:
            return alt_highest

    return ""


def find_old_rev(preferred_pn: str, installed: dict) -> str:
    """
    For REV_SENSITIVE parts only.
    Returns the highest revision lower than required that exists in the export.
    """
    base, req_rev = parse_pn(preferred_pn)
    if base not in REV_SENSITIVE or req_rev is None:
        return ""
    pattern = re.compile(r"^" + re.escape(base) + r"-(\d+)$")
    older   = [(int(m.group(1)), ipn)
               for ipn in installed
               if (m := pattern.match(ipn)) and int(m.group(1)) < req_rev]
    return max(older, default=(None, ""))[1] if older else ""


# =============================================================================
# DESCRIPTION & SN/LOT DETECTION
# =============================================================================

def get_best_desc(rows: pd.DataFrame) -> str:
    """Return the most common non-empty Component Description for these rows."""
    if rows is None or rows.empty:
        return ""
    s = rows["Component Description"].str.strip().replace("", pd.NA).dropna()
    return s.value_counts().idxmax() if not s.empty else ""


def detect_snlot(rows: pd.DataFrame) -> str:
    """
    Determine tracking type from the actual export data.
    Any non-empty Lot # → LOT, otherwise → SN.
    """
    return "LOT" if rows["Lot #"].str.strip().ne("").any() else "SN"


# =============================================================================
# ID ALLOCATION
# =============================================================================

def allocate_sns(rows: pd.DataFrame, qty_req: int,
                 sn_pool: dict, matched_pn: str) -> list[str]:
    """
    Pull exactly qty_req serial numbers from the pool for this PN.
    Pool is initialised once per matched PN so serials are never reused.
    """
    if matched_pn not in sn_pool:
        sn_pool[matched_pn] = sorted(
            {str(x).strip() for x in rows["Serial #"] if str(x).strip()}
        )
    pool = sn_pool[matched_pn]
    return [pool.pop(0) for _ in range(min(qty_req, len(pool)))]


def allocate_lots(rows: pd.DataFrame) -> list[str]:
    """
    Return one display string per distinct lot number.
    Same lot → summed into one row  e.g. 'LOT-A (x3)'
    Diff lots → separate rows       e.g. ['LOT-A (x2)', 'LOT-B (x1)']
    """
    lot_rows = rows[rows["Lot #"].str.strip() != ""].copy()
    if lot_rows.empty:
        return []
    lot_rows["Lot_clean"] = lot_rows["Lot #"].str.strip()
    lot_rows["Qty_num"]   = lot_rows["Quantity"].apply(to_float)
    per_lot = lot_rows.groupby("Lot_clean")["Qty_num"].sum()
    return [f"{lot} (x{int(q)})" for lot, q in per_lot.items()]
