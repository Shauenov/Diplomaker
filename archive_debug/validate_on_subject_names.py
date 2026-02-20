# -*- coding: utf-8 -*-
"""
validate_on_subject_names.py
Check if all ОН subjects in template can be found in source
"""

import pandas as pd
import re
from generate_diploma_it_kz import PAGE2_SUBJECTS, PAGE3_SUBJECTS, PAGE4_SUBJECTS

def is_on_subject(name):
    """Check if subject is an ОН sub-item."""
    s = str(name).strip()
    return s.startswith("ОН ") or s.startswith("РО ")

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

# Get all ОН subjects from template
TEMPLATE_ON = []
for subj in PAGE2_SUBJECTS + PAGE3_SUBJECTS + PAGE4_SUBJECTS:
    if is_on_subject(subj):
        TEMPLATE_ON.append(subj)

print(f"Template ОН subjects: {len(TEMPLATE_ON)}\n")

# Read source subjects
SOURCE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
ROW_SUBJECT_NAMES = 1
COL_START_SUBJECTS = 2

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
    source_subjects.append((name_kz, name_ru, primary_raw))

# Filter source ОН
source_on = [(kz, ru, raw) for kz, ru, raw in source_subjects if is_on_subject(kz)]

print(f"Source ОН subjects: {len(source_on)}\n")

# Create lookup maps
template_norm_on = {normalize_key(s): s for s in TEMPLATE_ON}
source_norm_on = {normalize_key(kz): (kz, ru, raw) for kz, ru, raw in source_on}

# Check for matches
print("=== TEMPLATE ОН subjects not found in SOURCE (by normalized matching) ===")
missing_in_source_on = []
for tnorm, tname in sorted(template_norm_on.items()):
    if tnorm not in source_norm_on:
        missing_in_source_on.append(tname)
        print(f"  MISSING: {tname}")

print(f"\nTotal missing ОН: {len(missing_in_source_on)}\n")

if missing_in_source_on:
    print("⚠ These ОН subjects will have NO GRADES because they're not in source file!")
    print("This causes their row to not populate with hours/credits/grades.")
else:
    print("✅ All template ОН subjects found in source!")
