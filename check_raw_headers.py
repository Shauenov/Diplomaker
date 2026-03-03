import pandas as pd
import openpyxl
import sys

if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

wb = openpyxl.load_workbook('local_test_copy.xlsx', read_only=True, data_only=True)
ws = wb['3D-1']
data = []
for row in ws.iter_rows(max_row=200, max_col=200, values_only=True):
    data.append(row)

df = pd.DataFrame(data)

# Let's find the header row (typically row 3 or 4 which we dynamic search)
row_hours = None
for i in range(10):
    val0 = str(df.iloc[i, 0]).lower() if pd.notna(df.iloc[i, 0]) else ''
    val1 = str(df.iloc[i, 1]).lower() if pd.notna(df.iloc[i, 1]) else ''
    if 'сағат' in val0 or 'часы' in val0 or 'сағат' in val1 or 'часы' in val1:
        row_hours = df.iloc[i]
        break

print("=== RAW HOURS/CREDITS ROW ===")
for col_idx in range(5, 40): # Adjust range to see subj columns
    subj_name = str(df.iloc[1, col_idx]).strip()
    if not subj_name or subj_name == 'nan':
         subj_name = str(df.iloc[2, col_idx]).strip()
    
    hour_val = str(row_hours[col_idx]).strip() if pd.notna(row_hours[col_idx]) else 'EMPTY'
    
    if subj_name and subj_name != 'nan':
        display_name = subj_name[:38] + ".." if len(subj_name) > 40 else subj_name
        print(f"Col {col_idx:<2}: {display_name:<40} | Raw H/C: {hour_val}")
