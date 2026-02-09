import pandas as pd

df = pd.read_excel("data/export_all_65c4618b186b4d068cc944cf7f28a71a.xlsx", dtype=str).fillna("")

# Pick the columns you want
cols_needed = ["ProductNo", "Component Description", "Serial #", "Lot #"]
missing = [c for c in cols_needed if c not in df.columns]
if missing:
    print("Missing columns:", missing)
    print("Available columns:", list(df.columns))
    raise SystemExit(1)

# Compute SerialOrLot
df["SerialOrLot"] = df["Serial #"].str.strip()
mask_blank = df["SerialOrLot"].eq("")
df.loc[mask_blank, "SerialOrLot"] = df.loc[mask_blank, "Lot #"].str.strip()

# Keep only what you want to print
out = df[["ProductNo", "Component Description", "SerialOrLot"]]

# Print first N rows
print(out.head(50).to_string(index=False))
