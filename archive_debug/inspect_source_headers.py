import pandas as pd

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def inspect():
    df = pd.read_excel(SOURCE_FILE, sheet_name=SHEET_NAME, header=None, nrows=10)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', 1000)
    print("--- First 10 rows ---")
    print(df)

if __name__ == "__main__":
    inspect()
