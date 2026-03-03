import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

df = pd.read_excel('local_test_copy.xlsx', sheet_name='3D-1', header=None)
print("=== 3D-1 Columns 70 to 130 ===")
for c in range(70, 130, 4):
    print(f"Col {c}:")
    print(f"  R1: {str(df.iloc[1, c])[:60].replace(chr(10), ' ')}")
    print(f"  R2: {str(df.iloc[2, c])[:60].replace(chr(10), ' ')}")
    print(f"  R3: {str(df.iloc[3, c])[:60].replace(chr(10), ' ')}")
