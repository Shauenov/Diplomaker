import pandas as pd
from src.parser import parse_excel_sheet
from src.generator import DiplomaGenerator
from configs import get_config
from src.utils import normalize_key

df = pd.read_excel('mock_filled_grades.xlsx', sheet_name='3D-1')
students = parse_excel_sheet(df, '3D-1')
s = students[0]

res = get_config('3D', 'kz')
config, terms, template_name = res[:3]
gen = DiplomaGenerator(f'templates/{template_name}', 'output/test.xlsx', config, terms)

grades = s['grades']

for ws in gen.workbook.worksheets:
    print(f'\\n--- {ws.title} ---')
    for row in range(1, min(30, ws.max_row + 1)):
        cell_b = ws.cell(row=row, column=2)
        subj = cell_b.value
        if subj and isinstance(subj, str) and subj.strip():
            nkey = normalize_key(subj)
            grade = grades.get(nkey)
            print(f"Row {row:2d} | subj={subj.replace(chr(10), ' ')[:30]:30s} | match?={'YES' if grade else 'NO'}")
            if grade: print(f"          -> hours={grade.get('hours')}, credits={grade.get('credits')}")
