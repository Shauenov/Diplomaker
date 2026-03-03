import pandas as pd
import sys
sys.stdout.reconfigure(encoding='utf-8')
from src.parser import parse_excel_sheet

df = pd.read_excel('local_test_copy.xlsx', sheet_name='3D-1', header=None)
students = parse_excel_sheet(df, '3D-1', start_row=5)
s = students[0]

for kz_key, info in s['grades'].items():
    subj = info['subject_kz']
    if 'ОН' in subj or 'КМ' in subj or 'ПМ' in subj:
        h = info.get('hours', '')
        c = info.get('credits', '')
        pts = info.get('points', '')
        let = info.get('letter', '')
        gpa = info.get('gpa', '')
        tr = info.get('traditional_kz', '')
        print(f"{subj[:35]:<35} | H:{h:<5} C:{c:<3} | P:{pts:<3} L:{let} G:{gpa} T:{tr}")
