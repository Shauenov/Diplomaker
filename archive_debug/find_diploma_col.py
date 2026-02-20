import openpyxl
import re

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def find_diploma_col():
    wb = openpyxl.load_workbook(SOURCE_FILE, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    
    # Regex for Diploma ID: e.g. KZ 1234567 or similar
    pattern = re.compile(r"^[A-Za-zА-Яа-я]{2}.*\d+")
    
    found_cols = {}
    
    print("Scanning for Diploma ID patterns...")
    for r in range(1, 20): # Scan first 20 rows
        for c in range(1, 100):
            val = ws.cell(row=r, column=c).value
            if val and isinstance(val, str):
                if pattern.match(val):
                    print(f"Match at Row {r}, Col {c}: {val}")
                    found_cols[c] = found_cols.get(c, 0) + 1
                    
    print("\nColumn Matches:")
    for c, count in found_cols.items():
        print(f"Col {c}: {count} matches")

if __name__ == "__main__":
    find_diploma_col()
