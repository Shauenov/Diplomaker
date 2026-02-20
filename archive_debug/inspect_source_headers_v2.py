import openpyxl

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def inspect():
    wb = openpyxl.load_workbook(SOURCE_FILE, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    
    print("--- Inspecting Row 1-4 for Headers ---")
    for r in range(1, 5):
        print(f"Row {r}:")
        for c in range(1, 200): # Check first 200 cols
            val = ws.cell(row=r, column=c).value
            if val:
                print(f"  Col {c}: {val}")

if __name__ == "__main__":
    inspect()
