import pandas as pd
import openpyxl
from src.parser import parse_excel_sheet
from src.utils import normalize_key

df = pd.read_excel('local_test_copy.xlsx', sheet_name='3D-1')
students = parse_excel_sheet(df, '3D-1', start_row=5)
s = students[0]
grades = s['grades']

print("=== Parsed Excel Keys ===")
for k in grades.keys():
    if k.startswith('он41') or k.startswith('он4'):
        print(f"EXCEL KEY: '{k}'")

wb = openpyxl.load_workbook("templates/Diplom_D_KZ_Template(4).xlsx")
ws = wb['Бет 3']
print("\n=== Template Keys ===")
for r in range(1, 40):
    val = ws.cell(row=r, column=2).value
    if val and 'ОН 4' in str(val):
        nkey = normalize_key(val)
        print(f"TEMPL KEY: '{nkey}'")
        if nkey in grades:
            print("  -> MATCHED!")
        else:
            print("  -> NO MATCH")
