import pandas as pd
import os

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"
ROW_SUBJECT_NAMES = 1       # Excel Row 2
COL_START_SUBJECTS = 2      # Column C

base_path = r"c:\Users\user\OneDrive\Рабочий стол\template"
file_path = os.path.join(base_path, SOURCE_FILE)

print(f"Loading {file_path}...")
try:
    df = pd.read_excel(file_path, sheet_name=SHEET_NAME, header=None)
    
    row_names = df.iloc[ROW_SUBJECT_NAMES]
    print(f"Total columns: {df.shape[1]}")
    
    count = 0
    print("\n--- Scanning Columns ---")
    for col_idx in range(COL_START_SUBJECTS, df.shape[1], 4):
        raw_name = row_names[col_idx]
        print(f"Col {col_idx}: {repr(raw_name)}")
        
        if pd.isna(raw_name):
            print("  -> SKIPPED (NaN)")
            continue
            
        count += 1
        
    print(f"\nTotal subjects found: {count}")
    
except Exception as e:
    print(f"Error: {e}")
