import pandas as pd
import re

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

# Indices
ROW_SUBJECT_NAMES = 17
ROW_HOURS = 19

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

    print("--- Row 18 (Names) & Row 20 (Hours) Analysis ---")
    row_names = df.iloc[ROW_SUBJECT_NAMES]
    row_hours = df.iloc[ROW_HOURS]
    
    # Print first 20 valid columns to align logic
    count = 0
    for col_idx in range(4, df.shape[1]):
        val = row_names[col_idx]
        if not pd.isna(val):
            # Found a potential subject column
            hour_val = row_hours[col_idx]
            print(f"Col {col_idx}: {clean(val)} [Hours: {clean(hour_val)}]")
            
            # Check if grade headers exist in Col+0, +1, +2, +3
            # Row 21 (Index 21)
            # headers = [df.iloc[20, i] for i in range(col_idx, col_idx+4)]
            # print(f"    Grade Headers: {headers}")
            # count += 1
            # if count > 10: break

            # Check alignment: if we expect next subject at col_idx + 4
            # We will see if the loop (range(4, end, 4)) would work
            pass

    print("\n--- Testing Step-by-4 Logic ---")
    for col_idx in range(4, 100, 4):
        name = row_names[col_idx]
        if pd.isna(name):
            print(f"Col {col_idx}: [EMPTY]")
        else:
            print(f"Col {col_idx}: {clean(name)}")

if __name__ == "__main__":
    main()
