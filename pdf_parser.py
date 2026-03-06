"""
pdf_parser.py
-------------
Reads the drone configuration PDF and extracts part requirements.

Each part is returned as a dict:
  base       — bare base PN from PDF          e.g. '150517'
  preferred  — base + highest PDF revision    e.g. '150517-02'
  desc       — description text
  qty        — quantity required
  alts       — list of (bare, preferred) tuples for alternate PNs
"""

import pdfplumber
from config import REV_SENSITIVE
from utils  import split_cell, with_rev, to_int


def _detect_table_type(header: list) -> str:
    first = str(header[0] or "").strip().lower()
    if first.startswith("description"):
        return "drone"
    if first.startswith("part"):
        return "battery"
    return "unknown"


def _resolve_pn_rev_pairs(pns_raw: str, revs_raw: str) -> list[tuple[str, str]]:
    """
    Pair each PN with its revision from the PDF cell.

    Handles two layouts:
      A) Same number of PNs and revs  → pair by index, keep highest rev per PN
         e.g. PNs='156976\\n156976'  Revs='2\\n1'  → [('156976', '2')]

      B) Fewer PNs than revs (one PN, multiple revs)
         → all revs belong to the single PN, pick the highest
         e.g. PNs='150298'  Revs='1\\n2'  → [('150298', '2')]

    Returns list of (base_pn, rev_str), first entry is primary.
    """
    pns  = split_cell(pns_raw)
    revs = split_cell(revs_raw)
    if not pns:
        return []

    # Layout B: one PN, multiple revs
    if len(pns) == 1 and len(revs) > 1:
        best_rev = max(revs, key=lambda r: int(r) if r.isdigit() else 0)
        return [(pns[0], best_rev)]

    # Layout A: pair by index
    pairs = [(pn, revs[i] if i < len(revs) else "0") for i, pn in enumerate(pns)]

    # Keep highest rev per PN (handles duplicate PNs with different revs)
    best: dict[str, tuple[str, int]] = {}
    for pn, rev in pairs:
        ri = int(rev) if rev.isdigit() else 0
        if pn not in best or ri > best[pn][1]:
            best[pn] = (rev, ri)

    # Return in original order of first appearance
    seen, result = set(), []
    for pn, _ in pairs:
        if pn not in seen:
            seen.add(pn)
            rev_str, _ = best[pn]
            result.append((pn, rev_str))
    return result


def _parse_table(table: list[list],
                 col_pn: int, col_rev: int,
                 col_qty: int, col_desc: int) -> list[dict]:
    """Parse one PDF table into a list of part dicts."""
    parts = []

    for row in table[1:]:   # skip header row
        def cell(i):
            try:
                return str(row[i] or "").strip()
            except IndexError:
                return ""

        desc = cell(col_desc)
        if not desc or desc.lower() == "description":
            continue

        pairs = _resolve_pn_rev_pairs(cell(col_pn), cell(col_rev))
        if not pairs:
            continue

        base0, rev0 = pairs[0]
        alts        = pairs[1:]

        parts.append({
            "base":      base0,
            "preferred": with_rev(base0, rev0),
            "desc":      desc,
            "qty":       to_int(cell(col_qty)),
            "alts":      [(b, with_rev(b, r)) for b, r in alts],
        })

    return parts


def read_pdf(pdf_path: str) -> tuple[list[dict], dict]:
    """
    Parse the PDF template.

    Returns
    -------
    parts   : list of part dicts (deduplicated by base PN)
    aliases : dict  base_pn -> [preferred_pn, alt_pref1, alt_pref2, ...]
    """
    all_parts: list[dict] = []
    aliases:   dict       = {}

    with pdfplumber.open(pdf_path) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            for table in (page.extract_tables() or []):
                if not table:
                    continue

                kind = _detect_table_type(table[0])

                if kind == "drone":
                    # Description(0) | Qty(1) | Part Num.(2) | Rev.(3) | ...
                    parsed = _parse_table(table, col_pn=2, col_rev=3,
                                          col_qty=1, col_desc=0)
                elif kind == "battery":
                    # Part Num.(0) | Rev.(1) | Qty(2) | Description(3) | ...
                    parsed = _parse_table(table, col_pn=0, col_rev=1,
                                          col_qty=2, col_desc=3)
                else:
                    continue

                print(f"  [PDF] Page {page_num} ({kind}): {len(parsed)} parts")
                all_parts.extend(parsed)

    # Deduplicate — keep first occurrence per base PN
    seen, dedup = set(), []
    for p in all_parts:
        if p["base"] not in seen:
            seen.add(p["base"])
            dedup.append(p)
            if p["alts"]:
                aliases[p["base"]] = [p["preferred"]] + [pref for _, pref in p["alts"]]

    print(f"\n  [PDF] {len(dedup)} unique parts | {len(aliases)} with alternates\n")
    return dedup, aliases
