"""
main.py
-------
Entry point. Run this file to process the BOM.

    python main.py

Output files are written to the outputs/ folder named after the drone's
parent serial number:
    {serial}_BOM.xlsx
    {serial}_part_list.xlsx
    {serial}_registration.xlsx
"""

import os
import re

from config     import TEMPLATE_PATH, EXPORT_PATH, OUTPUT_DIR
from pdf_parser import read_pdf
from matcher    import (load_export, get_parent_serial,
                        find_match, find_old_rev,
                        get_best_desc, detect_snlot,
                        allocate_sns, allocate_lots)
from writers    import write_bom, write_part_list, write_registration
from utils      import to_int


def main():
    os.makedirs(OUTPUT_DIR, exist_ok=True)

    # ── 1. Parse PDF template ─────────────────────────────────────────────────
    print("Reading PDF template …")
    parts, aliases = read_pdf(TEMPLATE_PATH)
    print(f"Total parts to check: {len(parts)}")

    # ── 2. Load export ────────────────────────────────────────────────────────
    df            = load_export(EXPORT_PATH)
    parent_serial = get_parent_serial(df)
    installed     = {pn: grp for pn, grp in df.groupby("ProductNo")}
    print(f"Parent Serial       : {parent_serial}")
    print(f"Unique PNs in export: {len(installed)}\n")

    # ── 3. Match each part ────────────────────────────────────────────────────
    sn_pool = {}
    results = []

    for part in parts:
        qty_req = part["qty"]
        if qty_req == 0:
            continue

        matched = find_match(part, installed)

        if matched:
            rows     = installed[matched]
            exp_desc = get_best_desc(rows)
            snlot    = detect_snlot(rows)

            if snlot == "SN":
                ids     = allocate_sns(rows, qty_req, sn_pool, matched)
                qty_act = len(ids)
            else:
                ids     = allocate_lots(rows)
                qty_act = sum(
                    to_int(m.group(1))
                    for s in ids
                    if (m := re.search(r"\(x(\d+)\)", s))
                )

            status = "SATISFIED" if qty_act >= qty_req else "NOT SATISFIED"

        else:
            old_pn = find_old_rev(part["preferred"], installed)

            if old_pn:
                rows     = installed[old_pn]
                exp_desc = get_best_desc(rows)
                snlot    = detect_snlot(rows)
                ids      = (allocate_sns(rows, qty_req, sn_pool, old_pn)
                            if snlot == "SN" else allocate_lots(rows))
                qty_act  = len(ids)
                matched  = old_pn
                status   = "OLD REV"
                print(f"  [OLD REV] {part['preferred']}  →  found {old_pn}")
            else:
                ids, qty_act, exp_desc, snlot = [], 0, "", "SN"
                status = "NOT FOUND"

        results.append({
            "pn":         part["preferred"],
            "matched_pn": matched,
            "desc":       exp_desc or part["desc"],
            "snlot":      snlot,
            "qty_req":    qty_req,
            "qty_act":    qty_act,
            "ids_text":   "\n".join(ids),
            "status":     status,
        })

    # ── 4. Print summary ──────────────────────────────────────────────────────
    counts = {s: sum(1 for r in results if r["status"] == s)
              for s in ("SATISFIED", "NOT SATISFIED", "OLD REV", "NOT FOUND")}
    print(f"\nSATISFIED    : {counts['SATISFIED']}")
    print(f"NOT SATISFIED: {counts['NOT SATISFIED']}")
    print(f"OLD REV      : {counts['OLD REV']}")
    print(f"NOT FOUND    : {counts['NOT FOUND']}")

    # ── 5. Write output files ─────────────────────────────────────────────────
    base = f"{OUTPUT_DIR}/{parent_serial}"
    write_bom(f"{base}_BOM.xlsx", df, results, parent_serial)
    write_part_list(f"{base}_part_list.xlsx", results, parent_serial)
    write_registration(f"{base}_registration.xlsx", parent_serial)

    print("\nDone.")


if __name__ == "__main__":
    main()
