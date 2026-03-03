import pandas as pd
from src.parser import parse_excel_sheet

df = pd.read_excel('mock_filled_grades.xlsx', sheet_name='3D-1')
students = parse_excel_sheet(df, '3D-1')
grades = students[0]['grades']

print('Keys starting with он1:')
for k, v in grades.items():
    if k.startswith('он1'):
        print(f"{k[:30]:30s} hours={v.get('hours')} credits={v.get('credits')}")
