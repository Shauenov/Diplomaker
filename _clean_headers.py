"""Clean leftover static text from D_RU and D_KZ templates (rows 9, 12 col A)."""
import openpyxl, io, shutil
from pathlib import Path

templates = [
    Path("templates/Diplom_D_RU_Template(4).xlsx"),
    Path("templates/Diplom_D_KZ_Template(4).xlsx"),
    Path("templates/Diplom_F_RU_Template (4).xlsx"),
    Path("templates/Diplom_F_KZ_Template (4).xlsx"),
]

for t in templates:
    if not t.exists():
        continue
    with open(t, 'rb') as f:
        buf = io.BytesIO(f.read())
    wb = openpyxl.load_workbook(buf)
    ws = wb.worksheets[0]
    
    cleaned = []
    for r in range(1, 15):
        for c in range(1, 9):  # A-H
            cell = ws.cell(r, c)
            if cell.value is not None:
                cleaned.append(f"  Row {r} Col {chr(64+c)}: '{cell.value}' -> cleared")
                cell.value = None
    
    if cleaned:
        print(f"\n{t.name}:")
        for line in cleaned:
            print(line)
        wb.save(str(t))
    else:
        print(f"\n{t.name}: clean")
    wb.close()

print("\nDone!")
