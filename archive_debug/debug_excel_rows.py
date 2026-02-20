import pandas as pd

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3D-1"

def debug_rows():
    df = pd.read_excel(SOURCE_FILE, sheet_name=SHEET_NAME, header=None, nrows=10)
    print(f"Row 3 (54-60): {df.iloc[3].tolist()[54:60]}")
    print(f"Row 2 (54-60): {df.iloc[2].tolist()[54:60]}")

if __name__ == "__main__":
    debug_rows()
