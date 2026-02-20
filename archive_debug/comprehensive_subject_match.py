# -*- coding: utf-8 -*-
"""
comprehensive_subject_match.py
Properly check if all source subjects match template
"""

import pandas as pd
import re
from generate_diploma_it_kz import PAGE1_SUBJECTS, PAGE2_SUBJECTS, PAGE3_SUBJECTS, PAGE4_SUBJECTS

TEMPLATE_ALL = PAGE1_SUBJECTS + PAGE2_SUBJECTS + PAGE3_SUBJECTS + PAGE4_SUBJECTS

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

# Read ALL 4 source sheets
SOURCE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
ROW_SUBJECT_NAMES = 1
COL_START_SUBJECTS = 2

all_source_subjects = set()

for sheet_name in ["3Ғ-1", "3Ғ-2", "3Ғ-3", "3Ғ-4"]:
    df = pd.read_excel(SOURCE, sheet_name=sheet_name, header=None)
    
    for col_idx in range(COL_START_SUBJECTS, df.shape[1], 4):
        raw_r2 = df.iloc[ROW_SUBJECT_NAMES, col_idx]
        raw_r3 = df.iloc[ROW_SUBJECT_NAMES + 1, col_idx] if ROW_SUBJECT_NAMES + 1 < len(df) else None
        
        r2_str = str(raw_r2).strip() if raw_r2 is not None and not pd.isna(raw_r2) else ""
        r3_str = str(raw_r3).strip() if raw_r3 is not None and not pd.isna(raw_r3) else ""
        
        if not r2_str and not r3_str:
            continue
        
        primary_raw = r3_str if r3_str else r2_str
        name_kz, name_ru = clean_subject_name(primary_raw)
        all_source_subjects.add(normalize_key(name_kz))

print(f"Source unique subjects (all 4 sheets): {len(all_source_subjects)}")
print(f"Template total subjects: {len(TEMPLATE_ALL)}\n")

# Build template map
template_norm = {normalize_key(s): s for s in TEMPLATE_ALL}

# Find missing
missing = []
for tnorm, tname in sorted(template_norm.items()):
    if tnorm not in all_source_subjects:
        missing.append(tname)

print(f"=== Missing from source ({len(missing)}) ===")
for name in missing:
    if name not in ["Қазақ тілі", "Физика", "Химия"]:  # Skip general subjects
        print(f"  {name[:75]}")

if len(missing) <= 5:
    print(f"\n✅ Only {len(missing)} subjects missing - likely just general education + electives")
else:
    print(f"\n❌ {len(missing)} subjects missing - PROBLEM!")
