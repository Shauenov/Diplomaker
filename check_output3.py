import openpyxl
import sys
sys.stdout.reconfigure(encoding='utf-8')

def check(file_path):
    print(f"\nChecking {file_path}")
    try:
        wb = openpyxl.load_workbook(file_path)
    except FileNotFoundError:
        print("File not found.")
        return
        
    for ws_name in wb.sheetnames:
        ws = wb[ws_name]
        for row in range(1, 65):
            val = ws.cell(row=row, column=2).value
            if val and isinstance(val, str) and any(x in val for x in ['ПМ', 'КМ', 'ОН', 'БМ', 'Ф', 'РО', 'практика']):
                h = ws.cell(row=row, column=3).value
                c = ws.cell(row=row, column=4).value
                tr = ws.cell(row=row, column=8).value
                print(f"Row {row:02d}: {str(val)[:30]:<30} | H: {h} C: {c} Trad: {tr}")

check("output/3D-1_Асанов Бекарыс Максатович_RU.xlsx")
check("output/3Ғ-1_Аймахан Балауса Абайханқызы_RU.xlsx")
check("output/3D-1_Асанов Бекарыс Максатович_KZ.xlsx")
