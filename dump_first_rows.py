import pandas as pd
import openpyxl

wb = openpyxl.load_workbook('local_test_copy.xlsx', read_only=True, data_only=True)
ws = wb['3D-1']
data = []
for row in ws.iter_rows(max_row=10, max_col=10, values_only=True):
    data.append(row)

df = pd.DataFrame(data)
with open("dump.txt", "w", encoding="utf-8") as f:
    for i in range(len(df)):
        f.write(f"Index {i}: {df.iloc[i].values}\n")
