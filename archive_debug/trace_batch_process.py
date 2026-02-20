# -*- coding: utf-8 -*-
"""
trace_batch_process.py
Show exactly what subjects batch_generate_it will find and store in grades_data
"""

import pandas as pd
import re
from generate_diploma_it_kz import PAGE2_SUBJECTS, PAGE3_SUBJECTS, PAGE4_SUBJECTS

SOURCE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
ROW_SUBJECT_NAMES = 1
ROW_SUB_NAMES = ROW_SUBJECT_NAMES + 1
COL_START_SUBJECTS = 2

def clean_subject_name(text):
    if not isinstance(text, str):
        return str(text).strip(), str(text).strip()
    parts = text.split('\n')
    if len(parts) >= 2:
        return parts[0].strip().rstrip(':').strip(), parts[1].strip().rstrip(':').strip()
    return text.strip().rstrip(':').strip(), text.strip().rstrip(':').strip()

def normalize_key(text):
    if not text:
        return ""
    t = str(text).lower()
    t = t.replace(".", "").replace(",", "").replace(":", "")
    t = t.replace(" ", "")
    t = re.sub(r'([a-zа-я]+)0+([1-9]+)', r'\1\2', t)
    return t.strip()

# All ON/RO subjects in template
TEMPLATE_ON_RU = []
for subj in PAGE2_SUBJECTS + PAGE3_SUBJECTS + PAGE4_SUBJECTS:
    if (str(subj).startswith("ОН ") or str(subj).startswith("РО ")):
        TEMPLATE_ON_RU.append(subj)

print(f"Template ОН/РО subjects: {len(TEMPLATE_ON_RU)}\n")

# Read source
df = pd.read_excel(SOURCE, sheet_name="3Ғ-1", header=None)

print("Subjects that batch_generate_it.py will STORE in grades_data:\n")

source_found = 0
for col_idx in range(COL_START_SUBJECTS, df.shape[1], 4):
    raw_r2 = df.iloc[ROW_SUBJECT_NAMES, col_idx]
    raw_r3 = df.iloc[ROW_SUB_NAMES, col_idx] if ROW_SUB_NAMES < len(df) else None
    
    r2_str = str(raw_r2).strip() if raw_r2 is not None and not pd.isna(raw_r2) else ""
    r3_str = str(raw_r3).strip() if raw_r3 is not None and not pd.isna(raw_r3) else ""
    
    if not r2_str and not r3_str:
        continue
    
    primary_raw = r3_str if r3_str else r2_str
    name_kz, name_ru = clean_subject_name(primary_raw)
    source_found += 1
    
    print(f"{source_found:2}. KZ: {name_kz[:60]}")
    print(f"    RU: {name_ru[:60]}")
    print()

print(f"=== Total source subjects to store: {source_found} ===")
print()
print("Plus hardcoded ELECTIVES:")
ELECTIVES_KZ = [
    "Ф1 Факультативтік ағылшын тілі",
    "Ф2 Факультативтік түрік тілі",
    "Ф3 Факультативтік кәсіпкерлік қызмет негіздері",
]
for e in ELECTIVES_KZ:
    print(f"  - {e}")

print(f"\nTOTAL in grades_data: {source_found} + 3 electives = {source_found + 3}")
