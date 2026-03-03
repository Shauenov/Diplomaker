import openpyxl
import os

templates = [
    r"c:\Users\user\OneDrive\Рабочий стол\template\templates\diploma_v4 (1).xlsx",
    r"c:\Users\user\OneDrive\Рабочий стол\template\templates\diploma_ru_template.xlsx",
    r"c:\Users\user\OneDrive\Рабочий стол\template\templates\Diplom_IT_KZ_Template.xlsx",
    r"c:\Users\user\OneDrive\Рабочий стол\template\templates\Diplom_IT_RU_Template.xlsx"
]

target_widths_left = {
    'A': 3.51, 'B': 24.07, 'C': 4.51, 'D': 8.03, 'E': 3.51, 'F': 6.27, 'G': 6.27, 'H': 9.03
}
target_widths_right = {
    'J': 3.51, 'K': 24.07, 'L': 4.51, 'M': 8.03, 'N': 3.51, 'O': 6.27, 'P': 6.27, 'Q': 9.03
}

for t in templates:
    if os.path.exists(t):
        print(f"Updating {os.path.basename(t)}...")
        wb = openpyxl.load_workbook(t)
        ws = wb.active
        
        for col, width in target_widths_left.items():
            ws.column_dimensions[col].width = width
            
        for col, width in target_widths_right.items():
            ws.column_dimensions[col].width = width
            
        # Optional separator
        ws.column_dimensions['I'].width = 3.0
            
        wb.save(t)
        print("Done.")
    else:
        print(f"Not found: {os.path.basename(t)}")
