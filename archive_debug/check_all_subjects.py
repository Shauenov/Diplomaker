# -*- coding: utf-8 -*-
"""
check_all_subjects.py
Count all subjects from both source and template
"""

import pandas as pd
from generate_diploma_it_kz import PAGE1_SUBJECTS, PAGE2_SUBJECTS, PAGE3_SUBJECTS, PAGE4_SUBJECTS

TEMPLATE_ALL = PAGE1_SUBJECTS + PAGE2_SUBJECTS + PAGE3_SUBJECTS + PAGE4_SUBJECTS

print(f"PAGE1_SUBJECTS: {len(PAGE1_SUBJECTS)}")
for i, s in enumerate(PAGE1_SUBJECTS[:3], 1):
    print(f"  {i}. {s[:60]}")
print()

print(f"PAGE2_SUBJECTS: {len(PAGE2_SUBJECTS)}")
for i, s in enumerate(PAGE2_SUBJECTS[:5], 1):
    print(f"  {i}. {s[:60]}")
print()

print(f"PAGE3_SUBJECTS: {len(PAGE3_SUBJECTS)}")
for i, s in enumerate(PAGE3_SUBJECTS[:3], 1):
    print(f"  {i}. {s[:60]}")
print()

print(f"PAGE4_SUBJECTS: {len(PAGE4_SUBJECTS)}")
for s in PAGE4_SUBJECTS:
    print(f"  - {s[:60]}")
print()

print(f"TOTAL TEMPLATE: {len(TEMPLATE_ALL)}")
print()

# Count source subjects
SOURCE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
ROW_SUBJECT_NAMES = 1
COL_START_SUBJECTS = 2

df = pd.read_excel(SOURCE, sheet_name="3Ғ-1", header=None)
source_count = 0
for col_idx in range(COL_START_SUBJECTS, df.shape[1], 4):
    raw_r2 = df.iloc[ROW_SUBJECT_NAMES, col_idx]
    raw_r3 = df.iloc[ROW_SUBJECT_NAMES + 1, col_idx] if ROW_SUBJECT_NAMES + 1 < len(df) else None
    
    r2_str = str(raw_r2).strip() if raw_r2 is not None and not pd.isna(raw_r2) else ""
    r3_str = str(raw_r3).strip() if raw_r3 is not None and not pd.isna(raw_r3) else ""
    
    if r2_str or r3_str:
        source_count += 1

print(f"TOTAL SOURCE: {source_count}")
print()

if source_count == len(TEMPLATE_ALL):
    print("✅ Counts MATCH!")
else:
    print(f"❌ MISMATCH: template={len(TEMPLATE_ALL)}, source={source_count}")
