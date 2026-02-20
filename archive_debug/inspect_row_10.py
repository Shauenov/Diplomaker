import openpyxl

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def inspect_row():
    wb = openpyxl.load_workbook(SOURCE_FILE, read_only=True, data_only=True)
    ws = wb[SHEET_NAME]
    
    target_row = 10
    print(f"--- Inspecting Row {target_row} ---")
    for c in range(1, 300):
        val = ws.cell(row=target_row, column=c).value
        # Filter out likely grades (numbers 0-5 or 0-100)
        is_grade = False
        try:
            if isinstance(val, (int, float)) and val <= 100:
                is_grade = True
        except: pass
        
        if val and not is_grade:
            print(f"Col {c}: {val}")

if __name__ == "__main__":
    inspect_row()
