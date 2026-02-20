import pandas as pd
import re

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

# Indices (Updated based on latest dump analysis)
ROW_SUBJECT_NAMES = 1
ROW_HOURS = 3

def parse_hours_credits(text):
    if not isinstance(text, str):
        return str(text), ""
    match = re.search(r"(\d+)с-(\d+(?:,\d+)?)к", text)
    if match:
        return match.group(1), match.group(2)
    return text, ""

def clean(text):
    if pd.isna(text): return "NaN"
    return str(text).replace('\n', ' | ')

def main():
    print(f"Reading {SOURCE_FILE}...")
    try:
        df = pd.read_excel(SOURCE_FILE, sheet_name=SHEET_NAME, header=None)
    except FileNotFoundError:
        print("File not found.")
        return

    print(f"--- Testing Row {ROW_SUBJECT_NAMES} (Names) & {ROW_HOURS} (Hours) ---")
    row_names = df.iloc[ROW_SUBJECT_NAMES]
    row_hours = df.iloc[ROW_HOURS]
    
    for col_idx in range(4, 150, 4):
        name = row_names[col_idx]
        hours_val = row_hours[col_idx]
        
        if pd.isna(name):
             # Try checking adjacent columns in case of merge issues?
             # But strictly following step 4:
             pass
        
        print(f"Col {col_idx}: {clean(name)} [Hours: {clean(hours_val)}]")

if __name__ == "__main__":
    main()
