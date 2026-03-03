import openpyxl
import os

templates = [
    r"c:\Users\user\OneDrive\Рабочий стол\template\templates\diploma_v4 (1).xlsx",
    r"c:\Users\user\OneDrive\Рабочий стол\template\templates\diploma_ru_template.xlsx",
    r"c:\Users\user\OneDrive\Рабочий стол\template\templates\Diplom_IT_KZ_Template.xlsx",
    r"c:\Users\user\OneDrive\Рабочий стол\template\templates\Diplom_IT_RU_Template.xlsx"
]

for t in templates:
    if os.path.exists(t):
        wb = openpyxl.load_workbook(t)
        ws = wb.active
        widths = {col: ws.column_dimensions[col].width for col in "ABCDEFGHIJKLMNOPQRSTUVWXYZ" if ws.column_dimensions[col].width}
        print(f"Template: {os.path.basename(t)}")
        print(widths)
        print("-" * 40)
    else:
        print(f"Not found: {os.path.basename(t)}")
