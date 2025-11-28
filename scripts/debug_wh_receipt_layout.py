import os
import pandas as pd

PATH = r"G:\Logistics\6_Reporting\1_PVS\WH Receipt FY25.xlsx"
SHEET = "Daily PVS"

print("[DEBUG] Inspecting:", PATH)
print("[DEBUG] Sheet:", SHEET)

if not os.path.exists(PATH):
    print("[ERROR] File not found")
    raise SystemExit(1)

# Read a small portion with no header so we see the raw layout
print("[DEBUG] Loading sheet head (no header)...")
df = pd.read_excel(PATH, sheet_name=SHEET, header=None, nrows=30)
print("[DEBUG] Shape:", df.shape)

# Show first 10 rows and first 40 columns
print("[DEBUG] Top-left block (rows 0-9, cols 0-19):")
print(df.iloc[0:10, 0:20])

# Locate any cells containing 'Target (LTP input)'
mask = df.eq("Target (LTP input)")
positions = list(zip(*mask.to_numpy().nonzero()))
print("[DEBUG] Positions of 'Target (LTP input)':", positions[:20])

# Show row(s) that contain that label
for (r, c) in positions[:5]:
    print(f"[DEBUG] Row {r} around 'Target (LTP input)':")
    print(df.iloc[r, 0:25])

print("[DONE]")
