import pandas as pd

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"

def search_value():
    xl = pd.ExcelFile(SOURCE_FILE)
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                val = df.iloc[r, c]
                if isinstance(val, (int, float)) and val >= 90 and val <= 100:
                    print(f"Found {val} in Sheet '{sheet}' at Row {r}, Col {c}")

if __name__ == "__main__":
    search_value()
