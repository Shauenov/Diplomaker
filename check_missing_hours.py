import pandas as pd
import openpyxl
from src.parser import parse_excel_sheet
import sys

# Force utf-8 for Windows console
if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def main():
    print("Loading excel 'local_test_copy.xlsx'...")
    wb = openpyxl.load_workbook('local_test_copy.xlsx', read_only=True, data_only=True)
    ws = wb['3D-1']
    data = []
    for row in ws.iter_rows(max_row=200, max_col=200, values_only=True):
        data.append(row)
    
    df = pd.DataFrame(data)
    students = parse_excel_sheet(df, '3D-1', start_row=5)
    
    s = students[0]
    print(f"\nStudent: {s['name']}")
    print("-" * 60)
    print(f"{'Subject':<40} | {'Type':<10} | {'Hours':<5} | {'Credits':<5}")
    print("-" * 60)
    
    import re
    
    for subj, grade_data in s['grades'].items():
        # Get raw subject names
        subj_kz = grade_data.get('subject_kz', '')
        
        # Check if it's a module
        is_mod = 'NO'
        if re.match(r'(БМ|КМ|ПМ|ОН|РО)\s*\.?\s*\d+', subj_kz):
            is_mod = 'YES'
            
        h = grade_data.get('hours', '')
        c = grade_data.get('credits', '')
        
        # Format for display
        display_name = subj_kz[:38] + ".." if len(subj_kz) > 40 else subj_kz
        print(f"{display_name:<40} | {is_mod:<10} | {str(h):<5} | {str(c):<5}")

if __name__ == "__main__":
    main()
