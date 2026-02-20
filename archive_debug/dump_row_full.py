import pandas as pd

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def dump_row_full():
    df = pd.read_excel(SOURCE_FILE, sheet_name=SHEET_NAME, header=None)
    row = df.iloc[11]
    with open("row_11_dump.txt", "w", encoding="utf-8") as f:
        for i, val in enumerate(row):
            f.write(f"Col {i}: {val}\n")

if __name__ == "__main__":
    dump_row_full()
