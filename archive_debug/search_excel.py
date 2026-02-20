import pandas as pd

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def search_string():
    df = pd.read_excel(SOURCE_FILE, sheet_name=SHEET_NAME, header=None)
    for r in range(df.shape[0]):
        val = str(df.iloc[r, 1])
        if "Бакирам" in val or "Болат Хамида" in val:
            print(f"Found student at Row {r}: {val}")
            # Check if any columns have numbers
            non_empty_cols = []
            for c in range(2, df.shape[1]):
                v = df.iloc[r, c]
                if pd.notna(v) and str(v).strip() != "" and str(v) != "0" and str(v) != "0.0":
                    non_empty_cols.append(c)
            print(f"  Non-empty columns: {non_empty_cols}")

if __name__ == "__main__":
    search_string()
