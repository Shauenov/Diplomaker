import pandas as pd
import openpyxl
from src.parser import parse_excel_sheet
import sys

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def main():
    wb = openpyxl.load_workbook('local_test_copy.xlsx', read_only=True, data_only=True)
    ws = wb['3D-1']
    data = []
    for row in ws.iter_rows(max_row=200, max_col=200, values_only=True):
        data.append(row)
    
    df = pd.DataFrame(data)
    students = parse_excel_sheet(df, '3D-1', start_row=5)
    s = students[0]
    
    print(f"Student: {s['name']}")
    print("-" * 70)
    for subj, g in s['grades'].items():
        name = g.get('subject_kz', subj)
        h = g.get('hours', '')
        c = g.get('credits', '')
        print(f"{name[:45]:<45} | H: {str(h):<3} | C: {str(c):<3}")

if __name__ == "__main__":
    main()
