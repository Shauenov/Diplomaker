# -*- coding: utf-8 -*-
import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')
try:
    df = pd.read_excel('2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx', sheet_name='3Ғ-1', header=None)
    cols = [i for i, v in enumerate(df.iloc[2]) if str(v) != 'nan' and '10.2' in str(v)]
    print('Columns with 10.2:', cols)
    if cols:
        print("Cols 180 to 205:")
        for i in range(2, 8):
            row_vals = [str(df.iloc[i, c]) for c in range(180, 205)]
            print(f"Row {i}:", " | ".join(row_vals))
except Exception as e:
    print("Error:", e)
