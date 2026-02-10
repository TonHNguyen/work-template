from openpyxl import load_workbook, Workbook
import pandas as pd
import os

TEMPLATE_PATH = "templates/template.xlsx"
EXPORT_PATH = "data/export_all_c80eb8e5ef3040379d61be5c46e8cd83.xlsx"
ALIASES_PATH = "data/aliases.xlsx"


def to_int_or_zero(v):
    if v is None:
        return 0
    if isinstance(v, str):
        s = v.strip()
        if s == "" or s.startswith("#"):
            return 0
        try:
            return int(float(s))
        except ValueError:
            return 0
    try:
        return int(v)
    except Exception:
        return 0


def read_requirements():
    wb = load_workbook(TEMPLATE_PATH, keep_vba=True, data_only=True)
    ws = wb["44807"]

    requirements = []
    row = 2
    while True:
        pn = ws[f"A{row}"].value
        if pn is None or str(pn).strip() == "":
            break

        pn = str(pn).strip()
        qty_required = to_int_or_zero(ws[f"C{row}"].value)

        sn_lot = ws[f"E{row}"].value
        sn_lot = str(sn_lot).strip().upper() if sn_lot else ""
        sn_lot = sn_lot if sn_lot in ("SN", "LOT") else "SN"

        requirements.append((pn, qty_required, sn_lot))
        row += 1

    return requirements


def load_export():
    return pd.read_excel(EXPORT_PATH, dtype=str).fillna("")


def load_aliases():
    # expects: col A primary, cols B..N alternates
    df = pd.read_excel(ALIASES_PATH, sheet_name="Aliases", dtype=str).fillna("")
    aliases = {}

    for _, row in df.iterrows():
        primary = str(row.iloc[0]).strip()
        if not primary:
            continue

        allowed = [primary]
        for v in row.iloc[1:]:
            v = str(v).strip()
            if v:
                allowed.append(v)

        # de-dup while preserving order
        seen = set()
        allowed = [x for x in allowed if not (x in seen or seen.add(x))]
        aliases[primary] = allowed

    return aliases


def match_one_pn(installed_rows: pd.DataFrame, sn_lot_type: str):
    if installed_rows is None or installed_rows.empty:
        return 0, []

    if sn_lot_type == "SN":
        ids = [str(x).strip() for x in installed_rows["Serial #"].tolist() if str(x).strip()]
        uniq = sorted(set(ids))
        return len(uniq), uniq

    # LOT can repeat: count rows with non-empty Lot #
    lots = [str(x).strip() for x in installed_rows["Lot #"].tolist() if str(x).strip()]
    qty_actual = len(lots)
    display = sorted(set(lots))  # just for readability
    return qty_actual, display


def resolve_match_pn(primary_pn: str, installed_by_pn: dict, aliases: dict):
    allowed = aliases.get(primary_pn, [primary_pn])
    for candidate in allowed:
        if candidate in installed_by_pn:
            return candidate
    return ""

def pick_description(rows_for_match: pd.DataFrame) -> str:
    if rows_for_match is None or rows_for_match.empty:
        return ""
    s = rows_for_match["Component Description"].astype(str).str.strip()
    s = s[s != ""]
    if s.empty:
        return ""
    return s.value_counts().idxmax()


def write_required_parts_xlsx(output_path: str, results: list[dict]):
    wb = Workbook()
    ws = wb.active
    ws.title = "RequiredParts"

    headers = [
        "MatchedPartNo",
        "Component Description",
        "SN/LOT",
        "QTY Required",
        "QTY Actual",
        "IDs",
        "Status",
    ]
    ws.append(headers)

    for r in results:
        ws.append([
            r["matched_pn"],
            r["description"],
            r["snlot"],
            r["qty_required"],
            r["qty_actual"],
            r["ids_text"],
            r["status"],
        ])

    wb.save(output_path)
    print("Saved:", output_path)


if __name__ == "__main__":
    reqs = read_requirements()
    print("Total requirements:", len(reqs))

    df = load_export()

    # validate export columns
    for c in ["ProductNo", "Serial #", "Lot #", "Component Description"]:
        if c not in df.columns:
            print("ERROR: export missing column:", c)
            print("Available columns:", list(df.columns))
            raise SystemExit(1)

    # build global index
    df["ProductNo_clean"] = df["ProductNo"].astype(str).str.strip()
    installed_by_pn = {pn: g for pn, g in df.groupby("ProductNo_clean")}
    print("Unique ProductNo count:", len(installed_by_pn))

    # load aliases (use sheet_name=0 if you don't want tab-name issues)
    aliases = load_aliases()
    print("Aliases loaded:", len(aliases))

    os.makedirs("outputs", exist_ok=True)

    results = []
    for primary_pn, qty_req, snlot in reqs:
        matched_pn = resolve_match_pn(primary_pn, installed_by_pn, aliases)
        rows_for_match = installed_by_pn.get(matched_pn) if matched_pn else None

        qty_act, ids = match_one_pn(rows_for_match, snlot)
        desc = pick_description(rows_for_match)

        # status logic (keep it simple)
        if qty_req > 0:
            status = "SATISFIED" if qty_act >= qty_req else "NOT SATISFIED"
        else:
            status = "FOUND" if qty_act > 0 else "MISSING"

        ids_text = "\n".join(ids)  # newline-separated in the Excel cell

        results.append({
            "matched_pn": matched_pn if matched_pn else "",
            "description": desc,
            "snlot": snlot,
            "qty_required": qty_req,
            "qty_actual": qty_act,
            "ids_text": ids_text,
            "status": status,
        })

    output_file = "outputs/required_parts.xlsx"
    write_required_parts_xlsx(output_file, results)

