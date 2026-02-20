# -*- coding: utf-8 -*-
"""
compare_subject_names.py
Detailed subject name comparison
"""

import pandas as pd
import re

SOURCE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"

from generate_diploma_it_kz import PAGE1_SUBJECTS

ROW_SUBJECT_NAMES = 1
COL_START_SUBJECTS = 2

def clean_subject_name(text):
    if not isinstance(text, str):
        return str(text).strip(), str(text).strip()
    parts = text.split('\n')
    if len(parts) >= 2:
        return parts[0].strip().rstrip(':').strip(), parts[1].strip().rstrip(':').strip()
    return text.strip().rstrip(':').strip(), text.strip().rstrip(':').strip()

print("Template PAGE1 subjects:")
for i, subj in enumerate(PAGE1_SUBJECTS):
    print(f"  {i+1:2}. {subj[:70]}")

print("\n" + "="*80)
print("Source file subjects (first 20):")

df = pd.read_excel(SOURCE, sheet_name="3Ғ-1", header=None)

source_subjects = []
for col_idx in range(COL_START_SUBJECTS, df.shape[1], 4):
    raw_r2 = df.iloc[ROW_SUBJECT_NAMES, col_idx]
    raw_r3 = df.iloc[ROW_SUBJECT_NAMES + 1, col_idx] if ROW_SUBJECT_NAMES + 1 < len(df) else None
    
    r2_str = str(raw_r2).strip() if raw_r2 is not None and not pd.isna(raw_r2) else ""
    r3_str = str(raw_r3).strip() if raw_r3 is not None and not pd.isna(raw_r3) else ""
    
    if not r2_str and not r3_str:
        continue
    
    primary_raw = r3_str if r3_str else r2_str
    name_kz, name_ru = clean_subject_name(primary_raw)
    source_subjects.append((name_kz, name_ru))

for i, (kz, ru) in enumerate(source_subjects[:20]):
    print(f"  {i+1:2}. KZ: {kz[:60]}")
    print(f"        RU: {ru[:60]}\n")
