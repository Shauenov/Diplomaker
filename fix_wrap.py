import openpyxl
import os
from openpyxl.styles import Alignment

templates = [
    r"c:\Users\user\OneDrive\Рабочий стол\template\templates\diploma_v4 (1).xlsx",
    r"c:\Users\user\OneDrive\Рабочий стол\template\templates\diploma_ru_template.xlsx",
    r"c:\Users\user\OneDrive\Рабочий стол\template\templates\Diplom_IT_KZ_Template.xlsx",
    r"c:\Users\user\OneDrive\Рабочий стол\template\templates\Diplom_IT_RU_Template.xlsx"
]

for t in templates:
    if os.path.exists(t):
        print(f"Fixing wrap text in {os.path.basename(t)}...")
        wb = openpyxl.load_workbook(t)
        
        for ws in wb.worksheets:
            for row in ws.iter_rows():
                # For each cell that has some data, enforce wrap_text
                for cell in row:
                    if cell.value is not None:
                        current_alignment = cell.alignment
                        if current_alignment:
                            cell.alignment = Alignment(
                                horizontal=current_alignment.horizontal,
                                vertical='center', 
                                text_rotation=current_alignment.text_rotation,
                                wrap_text=True,
                                shrink_to_fit=current_alignment.shrink_to_fit,
                                indent=current_alignment.indent
                            )
                        else:
                            cell.alignment = Alignment(wrap_text=True, vertical='center')
        wb.save(t)
        print("Done.")
    else:
        print(f"Not found: {os.path.basename(t)}")
