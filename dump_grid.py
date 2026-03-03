import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_excel('local_test_copy.xlsx', sheet_name='3D-1', header=None)

cols = range(71, 77)
print(f"     | " + " | ".join([f"Col {c:<8}" for c in cols]))
print("-" * 80)
for r in range(0, 8):
    row_vals = []
    for c in cols:
        val = str(df.iloc[r, c]).replace('\n', ' ')
        if len(val) > 10: val = val[:7] + "..."
        row_vals.append(f"{val:<12}")
    print(f"R{r:<3} | " + " | ".join(row_vals))
