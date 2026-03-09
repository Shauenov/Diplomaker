"""Check 3D sheet columns around ОН 4.4 area to verify index positions."""
import pandas as pd
import sys

SRC = r"2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (2).xlsx"

# Read 3D-1 sheet
df = pd.read_excel(SRC, sheet_name="3D-1", header=None)

print(f"Sheet columns: {len(df.columns)}")
print()

# Check rows 1-3 (subject names and hours) for columns 110-160
print("=== 3D-1: Columns 110-160, Rows 1-3 ===")
for col in range(110, min(160, len(df.columns))):
    vals = []
    for r in [1, 2, 3]:
        v = df.iloc[r, col] if r < len(df) else ""
        if pd.notna(v) and str(v).strip():
            vals.append(f"R{r}={str(v).strip()[:80]}")
    if vals:
        print(f"  Col {col}: {' | '.join(vals)}")

print()
# Also check the specific column indices from SUBJECT_COLUMNS["3D"]
print("=== Checking specific SUBJECT_COLUMNS indices ===")
check_cols = [117, 120, 123, 126, 130, 133, 136, 140, 143, 149, 150, 151]
for col in check_cols:
    if col < len(df.columns):
        r1 = str(df.iloc[1, col]).strip()[:60] if pd.notna(df.iloc[1, col]) else "—"
        r2 = str(df.iloc[2, col]).strip()[:60] if pd.notna(df.iloc[2, col]) else "—"
        r3 = str(df.iloc[3, col]).strip()[:60] if pd.notna(df.iloc[3, col]) else "—"
        # Check first student data row (row 5)
        r5 = str(df.iloc[5, col]).strip()[:30] if 5 < len(df) and pd.notna(df.iloc[5, col]) else "—"
        print(f"  Col {col}: R1=[{r1}] R2=[{r2}] R3=[{r3}] R5=[{r5}]")

# Also dump ALL columns from 50-70 to check БМ area
print()
print("=== 3D-1: Columns 50-70 (БМ area), Rows 1-3 ===")
for col in range(50, min(70, len(df.columns))):
    vals = []
    for r in [1, 2, 3]:
        v = df.iloc[r, col] if r < len(df) else ""
        if pd.notna(v) and str(v).strip():
            vals.append(f"R{r}={str(v).strip()[:80]}")
    if vals:
        print(f"  Col {col}: {' | '.join(vals)}")
