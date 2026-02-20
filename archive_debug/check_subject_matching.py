# -*- coding: utf-8 -*-
"""
check_subject_matching.py
=========================
Compare subject names from source file with template definitions.
Highlight subjects that exist in template but not found in source.
"""

import pandas as pd
import re

SOURCE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"

# Import template subjects
from generate_diploma_it_kz import PAGE1_SUBJECTS, PAGE2_SUBJECTS, PAGE3_SUBJECTS, PAGE4_SUBJECTS

TEMPLATE_ALL = PAGE1_SUBJECTS + PAGE2_SUBJECTS + PAGE3_SUBJECTS + PAGE4_SUBJECTS

ROW_SUBJECT_NAMES = 1
ROW_HOURS = 3
COL_START_SUBJECTS = 2

def normalize_key(text):
    if not text:
        return ""
    t = str(text).lower()
    t = t.replace(".", "").replace(",", "").replace(":", "")
    t = t.replace(" ", "")
    t = re.sub(r'([a-zа-я]+)0+([1-9]+)', r'\1\2', t)
    return t.strip()

def clean_subject_name(text):
    if not isinstance(text, str):
        return str(text).strip(), str(text).strip()
    parts = text.split('\n')
    if len(parts) >= 2:
        return parts[0].strip().rstrip(':').strip(), parts[1].strip().rstrip(':').strip()
    return text.strip().rstrip(':').strip(), text.strip().rstrip(':').strip()

print("Extracting subjects from source file...\n")

df = pd.read_excel(SOURCE, sheet_name="3Ғ-1", header=None)
row_names = df.iloc[ROW_SUBJECT_NAMES]

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
    source_subjects.append(name_kz)

print(f"Source subjects: {len(source_subjects)}")
print(f"Template subjects: {len(TEMPLATE_ALL)}\n")

# Normalize and match
source_norm = {normalize_key(s): s for s in source_subjects}
template_norm = {normalize_key(s): s for s in TEMPLATE_ALL}

print("=== Subjects in TEMPLATE but NOT in SOURCE (normalized) ===")
missing_in_source = []
for tnorm, tname in template_norm.items():
    if tnorm not in source_norm:
        missing_in_source.append(tname)
        print(f"  {tname[:60]}")

print(f"\nTotal missing: {len(missing_in_source)}\n")

print("=== Subjects in SOURCE but NOT in TEMPLATE (normalized) ===")
extra_in_source = []
for snorm, sname in source_norm.items():
    if snorm not in template_norm:
        extra_in_source.append(sname)
        print(f"  {sname[:60]}")

print(f"\nTotal extra: {len(extra_in_source)}")
