import pandas as pd

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"

def list_sheets():
    xl = pd.ExcelFile(SOURCE_FILE)
    print(f"Sheets: {xl.sheet_names}")

if __name__ == "__main__":
    list_sheets()
