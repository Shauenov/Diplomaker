"""Detailed column mapping check for 3D sheet around ОН 4.3-4.5 area."""
import pandas as pd

SRC = r"2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (2).xlsx"
df = pd.read_excel(SRC, sheet_name="3D-1", header=None)

# Dump rows 1-5 for columns 115-145 to see exact sub-column structure
print("=== Rows 0-5, Cols 115-145 ===")
for col in range(115, min(145, len(df.columns))):
    vals = []
    for r in range(6):
        v = df.iloc[r, col]
        if pd.notna(v):
            s = str(v).strip()[:50]
            vals.append(f"R{r}=[{s}]")
        else:
            vals.append(f"R{r}=[-]")
    print(f"  Col {col}: {' '.join(vals)}")

print()
# Also check row 4 for ALL subject columns (54-140) to see "п б цэ трад" pattern
print("=== Row 4 (label row) for cols 54-145 ===")
labels = []
for col in range(54, min(145, len(df.columns))):
    v = df.iloc[4, col]
    if pd.notna(v) and str(v).strip():
        labels.append(f"Col{col}={str(v).strip()}")
print("  " + " | ".join(labels))

# Count how many sub-columns per subject (distance between "п" markers)
print()
print("=== Subject starts (col label 'п' in row 4) ===")
subj_starts = []
for col in range(2, len(df.columns)):
    v = df.iloc[4, col]
    if pd.notna(v) and str(v).strip().lower() == 'п':
        r2 = str(df.iloc[2, col]).strip()[:50] if pd.notna(df.iloc[2, col]) else ""
        r1 = str(df.iloc[1, col]).strip()[:50] if pd.notna(df.iloc[1, col]) else ""
        name = r2 if r2 else r1
        print(f"  Col {col}: {name}")
        subj_starts.append(col)

# Show gaps
print()
print("=== Gaps between subject start columns ===")
for i in range(1, len(subj_starts)):
    gap = subj_starts[i] - subj_starts[i-1]
    if gap != 4:
        print(f"  ** NON-4 GAP: Col {subj_starts[i-1]} → Col {subj_starts[i]} = {gap} columns")
