import pandas as pd
import openpyxl
import sys
from src.parser import parse_excel_sheet

# Force utf-8 for stdout
sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('local_test_copy.xlsx', read_only=True, data_only=True)
ws = wb['3D-1']
data = []
for row in ws.iter_rows(max_row=200, max_col=200, values_only=True):
    data.append(row)

df = pd.DataFrame(data)
students = parse_excel_sheet(df, '3D-1', start_row=5)
s1 = students[0]
print("Student:", s1['name'])
for subj, v in list(s1['grades'].items())[:5]:
    h = v.get('hours', 'MISSING_HOURS')
    c = v.get('credits', 'MISSING_CREDITS')
    t = v.get('subject_kz', 'MISSING')
    print(f"Subj: {subj[:10]} | H: {h} | C: {c} | Raw: {t[:15]}")
