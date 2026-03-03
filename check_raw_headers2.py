import pandas as pd
import openpyxl
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
    for row in ws.iter_rows(max_row=5, max_col=200, values_only=True):
        data.append(row)
    
    df = pd.DataFrame(data)
    
    row_hours = None
    for i in range(5):
        val0 = str(df.iloc[i, 0]).lower() if pd.notna(df.iloc[i, 0]) else ''
        val1 = str(df.iloc[i, 1]).lower() if pd.notna(df.iloc[i, 1]) else ''
        if 'сағат' in val0 or 'часы' in val0 or 'сағат' in val1 or 'часы' in val1:
            row_hours = df.iloc[i]
            break

    print("=== ALL SUBJECTS ===")
    for col in range(2, df.shape[1], 4):
        subj2 = str(df.iloc[2, col]).strip() if pd.notna(df.iloc[2, col]) else ''
        subj1 = str(df.iloc[1, col]).strip() if pd.notna(df.iloc[1, col]) else ''
        
        name = subj2 if (subj2 and subj2 != 'nan' and 'сағат' not in subj2.lower()) else subj1
        if name and name != 'nan':
            name_d = name[:40].replace('\n', ' ')
            h = str(row_hours[col]).strip() if row_hours is not None and pd.notna(row_hours[col]) else ''
            print(f"Col {col:<3} | {name_d:<40} | H: {h}")

if __name__ == "__main__":
    main()
