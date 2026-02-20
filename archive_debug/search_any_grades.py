import pandas as pd
import numpy as np

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def search_any_grades():
    df = pd.read_excel(SOURCE_FILE, sheet_name=SHEET_NAME, header=None)
    for r in range(df.shape[0]):
        # Check Columns 2 to 200 for any float values > 0
        row_vals = df.iloc[r, 2:200]
        # Ignore strings, nans, and zeros
        numeric_vals = [v for v in row_vals if isinstance(v, (int, float)) and v > 0]
        if len(numeric_vals) > 5:
            print(f"Row {r} ({df.iloc[r, 1]}) has {len(numeric_vals)} numeric grades! Sample: {numeric_vals[:5]}")

if __name__ == "__main__":
    search_any_grades()
