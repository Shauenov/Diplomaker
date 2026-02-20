# -*- coding: utf-8 -*-
"""
diagnose_missing_hours.py
=========================
Check which subjects in the source file have missing hours/credits.
"""

import pandas as pd
import openpyxl

SOURCE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"

ROW_SUBJECT_NAMES = 1       # Excel Row 2
ROW_HOURS = 3               # Excel Row 4
COL_START_SUBJECTS = 2      # Column C

def parse_hours_credits(text):
    """Parse '72с-3к' into (72, 3)."""
    if not isinstance(text, str) or text.lower() == "nan" or not text.strip():
        return "", ""
    import re
    match = re.search(r"(\d+)с-(\d+(?:,\d+)?)к", text)
    if match:
        return match.group(1), match.group(2)
    return text, ""

def clean_subject_name(text):
    """Split 'NameKZ\nNameRU' into ('NameKZ', 'NameRU')."""
    if not isinstance(text, str):
        return str(text).strip(), str(text).strip()
    parts = text.split('\n')
    if len(parts) >= 2:
        return parts[0].strip().rstrip(':').strip(), parts[1].strip().rstrip(':').strip()
    return text.strip().rstrip(':').strip(), text.strip().rstrip(':').strip()

print(f"Analyzing: {SOURCE}\n")

for sheet_name in ["3Ғ-1", "3Ғ-2", "3Ғ-3", "3Ғ-4"]:
    print(f"=== Sheet: {sheet_name} ===")
    df = pd.read_excel(SOURCE, sheet_name=sheet_name, header=None)
    
    row_hours_data = df.iloc[ROW_HOURS]
    
    missing = []
    filled = []
    
    for col_idx in range(COL_START_SUBJECTS, df.shape[1], 4):
        raw_r2 = df.iloc[ROW_SUBJECT_NAMES, col_idx]
        raw_r3 = df.iloc[ROW_SUBJECT_NAMES + 1, col_idx] if ROW_SUBJECT_NAMES + 1 < len(df) else None
        
        r2_str = str(raw_r2).strip() if raw_r2 is not None and not pd.isna(raw_r2) else ""
        r3_str = str(raw_r3).strip() if raw_r3 is not None and not pd.isna(raw_r3) else ""
        
        if not r2_str and not r3_str:
            continue
        
        primary_raw = r3_str if r3_str else r2_str
        name_kz, name_ru = clean_subject_name(primary_raw)
        
        raw_hours_val = row_hours_data[col_idx]
        h_raw_str = str(raw_hours_val).strip() if pd.notna(raw_hours_val) else ""
        
        hours, credits = parse_hours_credits(h_raw_str)
        
        if not hours:
            missing.append((col_idx, name_kz[:50]))
        else:
            filled.append((col_idx, name_kz[:50], hours, credits))
    
    print(f"  Filled: {len(filled)}")
    print(f"  Missing hours: {len(missing)}")
    if missing:
        print("  Missing subjects:")
        for col, name in missing:
            print(f"    col {col:3d}: {name}")
    print()
