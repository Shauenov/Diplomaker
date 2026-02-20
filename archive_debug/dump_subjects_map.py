import pandas as pd
import re

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def clean_subject_name(text):
    if not isinstance(text, str):
        return str(text), str(text)
    parts = text.split('\n')
    if len(parts) >= 2:
        return parts[0].strip(), parts[1].strip()
    return text.strip(), text.strip()

def dump():
    df = pd.read_excel(SOURCE_FILE, sheet_name=SHEET_NAME, header=None)
    row_names = df.iloc[2]
    
    print("Column | KZ Name | RU Name")
    for col_idx in range(0, df.shape[1], 4):
        raw_name = row_names[col_idx] if col_idx < len(row_names) else None
        if pd.isna(raw_name): continue
        kz, ru = clean_subject_name(raw_name)
        print(f"{col_idx} | {kz} | {ru}")

if __name__ == "__main__":
    dump()
