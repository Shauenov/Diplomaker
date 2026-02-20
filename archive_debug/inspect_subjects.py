import openpyxl

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def list_subjects():
    wb = openpyxl.load_workbook(SOURCE_FILE, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    
    # Rows 1-5 might contain headers
    print("--- Scanning Rows 1-5 for Subject Names ---")
    for r in range(1, 6):
        print(f"Row {r}:")
        for c in range(1, 200):
            val = ws.cell(row=r, column=c).value
            if val and isinstance(val, str) and ("Факультатив" in val or "ағылшын" in val or "түрік" in val or "кәсіпкерлік" in val):
                print(f"  Col {c}: {val}")

if __name__ == "__main__":
    list_subjects()
