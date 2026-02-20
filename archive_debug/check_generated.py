import pandas as pd

file = "Diplomas_Batch/Бакирам Балнұр Бегімжанқызы_KZ.xlsx"
df = pd.read_excel(file, sheet_name="Бет 1", header=None)
print("--- Sheet: Бет 1 ---")
print(df.iloc[18:25, 0:8])

pd.set_option('display.max_columns', None)
df4 = pd.read_excel(file, sheet_name="Бет 4", header=None)
print("\n--- Sheet: Бет 4 (Full Columns) ---")
# Find row where Col 0 is 62
row62 = df4[df4[0] == 62]
print(row62)
if not row62.empty:
    print("\nSubject 62 values:")
    for i, val in enumerate(row62.iloc[0].tolist()):
        print(f"  Col {i}: {repr(val)}")
