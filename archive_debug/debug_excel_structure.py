import pandas as pd

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def main():
    print(f"Reading {SOURCE_FILE}...")
    try:
        df = pd.read_excel(SOURCE_FILE, sheet_name=SHEET_NAME, header=None)
    except FileNotFoundError:
        print("File not found.")
        return

    print("--- First 10 Rows ---")
    for i in range(10):
        row = df.iloc[i].tolist()
        # Print only non-nan values to reduce noise
        clean_row = [str(x)[:20] for x in row if not pd.isna(x) and str(x).strip() != ""]
        print(f"Row {i}: {clean_row}")

if __name__ == "__main__":
    main()
