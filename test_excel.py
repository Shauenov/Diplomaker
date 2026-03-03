import pandas as pd
df = pd.read_excel('2025-2026 диплом ба?алары (ТОЛЫ?) ?ызыл диплом жазыл?ан со??ысы точно (1).xlsx', sheet_name='3D-1', header=None)
print("Row 2:")
print(df.iloc[2].dropna().tolist()[:10])
print("Row 3:")
print(df.iloc[3].dropna().tolist()[:10])
