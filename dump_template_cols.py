import openpyxl

wb = openpyxl.load_workbook('templates/Diplom_D_KZ_Template(4).xlsx')
ws = wb.worksheets[0]

with open('dump_template.txt', 'w', encoding='utf-8') as f:
    for row in range(12, 16):
        f.write(f"Row {row}:\n")
        for col in range(1, 10):
            val = ws.cell(row=row, column=col).value
            val_str = str(val).replace('\n', ' ') if val else ""
            if val_str:
                f.write(f"  Col {col}: {val_str[:30]}\n")
