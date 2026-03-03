"""Fix continuous numbering in column A across all 4 pages for all templates."""
import sys, io, openpyxl, os

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

templates = [
    'templates/Diplom_F_KZ_Template (4).xlsx',
    'templates/Diplom_F_RU_Template (4).xlsx',
    'templates/Diplom_D_KZ_Template(4).xlsx',
    'templates/Diplom_D_RU_Template(4).xlsx',
]

for tmpl in templates:
    wb = openpyxl.load_workbook(tmpl)
    global_num = 0
    for i, ws in enumerate(wb.worksheets):
        start_row = 19 if i == 0 else 2
        for r in range(start_row, ws.max_row + 1):
            b = ws.cell(r, 2).value
            if b and isinstance(b, str) and b.strip():
                global_num += 1
                ws.cell(r, 1).value = global_num
    wb.save(tmpl)
    wb.close()
    print(f'Fixed {tmpl}: {global_num} subjects numbered 1-{global_num}')

# Verify
for tmpl in templates:
    wb = openpyxl.load_workbook(tmpl)
    short = os.path.basename(tmpl)
    for i, ws in enumerate(wb.worksheets):
        start_row = 19 if i == 0 else 2
        nums = []
        for r in range(start_row, ws.max_row + 1):
            b = ws.cell(r, 2).value
            a = ws.cell(r, 1).value
            if b and isinstance(b, str) and b.strip():
                nums.append(str(a))
        if nums:
            print(f'  {short} {ws.title}: {nums[0]}-{nums[-1]}')
    wb.close()
print('Done!')
