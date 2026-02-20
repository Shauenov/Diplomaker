import pandas as pd
import re
from batch_generate_it import get_student_row_data, clean_subject_name, parse_hours_credits

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def debug_data():
    df = pd.read_excel(SOURCE_FILE, sheet_name=SHEET_NAME, header=None)
    
    # Parse Subjects
    subject_columns = []
    row_names = df.iloc[2]
    row_hours = df.iloc[3]
    for col_idx in range(2, df.shape[1], 4):
        raw_name = row_names[col_idx]
        if pd.isna(raw_name): continue
        name_kz, name_ru = clean_subject_name(raw_name)
        hours, credits = parse_hours_credits(str(row_hours[col_idx]))
        subject_columns.append({"col_idx": col_idx, "name_kz": name_kz, "name_ru": name_ru, "hours": hours, "credits": credits})

    # Student 1 (Row 10)
    student_data = get_student_row_data(df, 9, subject_columns)
    print("--- Student Info ---")
    print(f"Name: {student_data['full_name']}")
    print(f"ID Clean: '{student_data['diploma_number_clean']}'")
    
    print("\n--- Grades Sample (First 5) ---")
    keys = list(student_data['grades'].keys())
    for k in keys[:10]:
        print(f"Key: '{k}' | Data: {student_data['grades'][k]}")

if __name__ == "__main__":
    debug_data()
