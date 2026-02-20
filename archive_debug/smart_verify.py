"""
Smart subject verification:
- Uses batch_generate_it.py's parse logic for IT groups (3F)
- Reads KZ & RU PAGE_SUBJECTS from the generators
- Directly checks which subjects MATCH and which are MISSING in grades_data
- Does same for accountants (3D) via batch scan
"""
import sys, os, re, json
sys.path.insert(0, r"c:\Users\user\OneDrive\Рабочий стол\template")
os.chdir(r"c:\Users\user\OneDrive\Рабочий стол\template")

import pandas as pd

# ─── Import subject lists from generators ─────────────────────────────────────
from generate_diploma_it_kz import (
    PAGE1_SUBJECTS as IT_KZ_P1, PAGE2_SUBJECTS as IT_KZ_P2,
    PAGE3_SUBJECTS as IT_KZ_P3, PAGE4_SUBJECTS as IT_KZ_P4,
    normalize_key as kz_nkey
)
from generate_diploma_it_ru import (
    PAGE1_SUBJECTS as IT_RU_P1, PAGE2_SUBJECTS as IT_RU_P2,
    PAGE3_SUBJECTS as IT_RU_P3, PAGE4_SUBJECTS as IT_RU_P4,
    normalize_key as ru_nkey
)

IT_KZ_ALL = IT_KZ_P1 + IT_KZ_P2 + IT_KZ_P3 + IT_KZ_P4
IT_RU_ALL = IT_RU_P1 + IT_RU_P2 + IT_RU_P3 + IT_RU_P4

# ─── IT Excel parsing (same as batch_generate_it.py) ────────────────────────
SOURCE_FILE = r"2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
IT_SHEETS   = ["3Ғ-1", "3Ғ-2", "3Ғ-3", "3Ғ-4"]
ACC_SHEETS  = ["3D-1", "3D-2"]

ROW_SUBJECT_NAMES = 1
ROW_HOURS         = 3
ROW_DATA_START    = 5
COL_FULL_NAME     = 1
COL_START_SUBJECTS = 2

def parse_hours_credits(text):
    if not isinstance(text, str) or text.lower() == "nan":
        return "", ""
    m = re.search(r"(\d+)с-(\d+(?:,\d+)?)к", text)
    return (m.group(1), m.group(2)) if m else (text, "")

def clean_subject_name(text):
    if not isinstance(text, str):
        return str(text).strip(), str(text).strip()
    parts = text.split('\n')
    return (parts[0].strip().rstrip(':'), parts[1].strip().rstrip(':')) if len(parts) >= 2 else (text.strip(), text.strip())

def normalize_key(text):
    if not text: return ""
    t = str(text).lower().replace('.','').replace(',','').replace(':','').replace(' ','')
    t = re.sub(r'([a-zа-яё\u0400-\u052f]+)0+([1-9]+)', r'\1\2', t)
    return t.strip()

def get_excel_subjects_for_sheet(sheet_name):
    """Parse Excel sheet and return {norm_kz_key: (kz_name, ru_name, hours, credits)}"""
    df = pd.read_excel(SOURCE_FILE, sheet_name=sheet_name, header=None)
    subj_row   = df.iloc[ROW_SUBJECT_NAMES]
    hours_row  = df.iloc[ROW_HOURS]

    subjects = {}
    col = COL_START_SUBJECTS
    while col < len(subj_row):
        cell_val = subj_row.iloc[col]
        if pd.isna(cell_val) or str(cell_val).strip() in ('', 'nan'):
            col += 4
            continue
        kz_name, ru_name = clean_subject_name(str(cell_val))
        h_cell = hours_row.iloc[col] if col < len(hours_row) else ''
        hours, credits = parse_hours_credits(str(h_cell))
        if kz_name:
            subjects[normalize_key(kz_name)] = {
                'kz': kz_name, 'ru': ru_name,
                'hours': hours, 'credits': credits
            }
        col += 4
    return subjects

# ─── Verify IT groups ─────────────────────────────────────────────────────────
print("=" * 70)
print("IT GROUPS (3F) — Verification")
print("=" * 70)

# Module headers that shouldn't be in Excel (they're summed)
SKIP_SUBJECTS = {'', 'Қорытынды аттестаттау :', 'Итоговая аттестация :'}

all_results = {}

for sheet in IT_SHEETS:
    print(f"\n── Sheet: {sheet} ──────────────────────")
    try:
        excel_subjects = get_excel_subjects_for_sheet(sheet)
    except Exception as e:
        print(f"  ERROR reading sheet: {e}")
        continue

    excel_norm_keys = set(excel_subjects.keys())

    # Check KZ subjects
    kz_missing = []
    kz_matched = []
    for subj in IT_KZ_ALL:
        if subj in SKIP_SUBJECTS or ':' in subj[-2:]:
            continue
        nk = normalize_key(subj)
        # Try prefix for БМ
        if nk in excel_norm_keys:
            kz_matched.append(subj)
        else:
            # BM prefix fallback
            bm_m = re.match(r'(БМ\s*\.?\s*\d+)', subj)
            if bm_m:
                prefix = normalize_key(bm_m.group(1))
                found = any(k.startswith(prefix) for k in excel_norm_keys)
                if found:
                    kz_matched.append(f"{subj} [via prefix]")
                    continue
            # PM / КМ header — check by prefix
            pm_m = re.match(r'(ПМ\s*\d+|КМ\s*\d+)', subj)
            if pm_m:
                prefix = normalize_key(pm_m.group(1))
                found = any(k.startswith(prefix) for k in excel_norm_keys)
                if found:
                    kz_matched.append(f"{subj} [module header, matched]")
                    continue
            kz_missing.append(subj)

    # Check RU subjects
    ru_missing = []
    ru_matched = []
    for subj in IT_RU_ALL:
        if subj in SKIP_SUBJECTS or ':' in subj[-2:]:
            continue
        nk = normalize_key(subj)
        if nk in excel_norm_keys:
            ru_matched.append(subj)
        else:
            bm_m = re.match(r'(БМ\s*\.?\s*\d+)', subj)
            if bm_m:
                prefix = normalize_key(bm_m.group(1))
                found = any(k.startswith(prefix) for k in excel_norm_keys)
                if found:
                    ru_matched.append(f"{subj} [via prefix]")
                    continue
            pm_m = re.match(r'(ПМ\s*\d+|КМ\s*\d+)', subj)
            if pm_m:
                prefix = normalize_key(pm_m.group(1))
                found = any(k.startswith(prefix) for k in excel_norm_keys)
                if found:
                    ru_matched.append(f"{subj} [module header, matched]")
                    continue
            ru_missing.append(subj)

    print(f"  Excel subjects in sheet: {len(excel_subjects)}")
    print(f"  KZ template subjects: {len(IT_KZ_ALL)} → matched={len(kz_matched)}, missing={len(kz_missing)}")
    print(f"  RU template subjects: {len(IT_RU_ALL)} → matched={len(ru_matched)}, missing={len(ru_missing)}")

    if kz_missing:
        print(f"\n  ❌ KZ SUBJECTS NOT FOUND IN EXCEL ({len(kz_missing)}):")
        for s in kz_missing:
            print(f"      '{s}'")
    else:
        print(f"  ✅ KZ: All subjects found in Excel!")

    if ru_missing:
        print(f"\n  ❌ RU SUBJECTS NOT FOUND IN EXCEL ({len(ru_missing)}):")
        for s in ru_missing:
            print(f"      '{s}'")
    else:
        print(f"  ✅ RU: All subjects found in Excel!")

    all_results[sheet] = {
        'kz_missing': kz_missing,
        'ru_missing': ru_missing,
        'excel_count': len(excel_subjects)
    }

# ─── ACCOUNTANTS (3D) ─────────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("ACCOUNTANT GROUPS (3D) — Verification")
print("=" * 70)

# Import accountant subject lists
try:
    import generate_diploma as acc_kz_mod
    import generate_diploma_ru as acc_ru_mod
    # Find all *_SUBJECTS lists
    acc_kz_subjects = []
    acc_ru_subjects = []
    for attr in dir(acc_kz_mod):
        if 'SUBJECT' in attr.upper() or attr.lower() in ['subjects', 'all_subjects']:
            val = getattr(acc_kz_mod, attr)
            if isinstance(val, list):
                acc_kz_subjects.extend(val)
    for attr in dir(acc_ru_mod):
        if 'SUBJECT' in attr.upper() or attr.lower() in ['subjects', 'all_subjects']:
            val = getattr(acc_ru_mod, attr)
            if isinstance(val, list):
                acc_ru_subjects.extend(val)
    print(f"  Accountant KZ subjects: {len(acc_kz_subjects)}")
    print(f"  Accountant RU subjects: {len(acc_ru_subjects)}")
except Exception as e:
    print(f"  Could not import accountant generators: {e}")
    acc_kz_subjects = []
    acc_ru_subjects = []

for sheet in ACC_SHEETS:
    print(f"\n── Sheet: {sheet} ──────────────────────")
    try:
        excel_subjects = get_excel_subjects_for_sheet(sheet)
    except Exception as e:
        print(f"  ERROR: {e}")
        continue
    print(f"  Excel subjects: {len(excel_subjects)}")

    if not acc_kz_subjects:
        print("  (No KZ/RU subject lists imported — skipping match check)")
        continue

    excel_norm_keys = set(excel_subjects.keys())
    kz_match = sum(1 for s in acc_kz_subjects
                   if normalize_key(s) in excel_norm_keys and s.strip() not in SKIP_SUBJECTS)
    kz_miss  = [s for s in acc_kz_subjects
                if normalize_key(s) not in excel_norm_keys and s.strip() not in SKIP_SUBJECTS]
    ru_match = sum(1 for s in acc_ru_subjects
                   if normalize_key(s) in excel_norm_keys and s.strip() not in SKIP_SUBJECTS)
    ru_miss  = [s for s in acc_ru_subjects
                if normalize_key(s) not in excel_norm_keys and s.strip() not in SKIP_SUBJECTS]

    print(f"  KZ matched={kz_match}/{len(acc_kz_subjects)}, missing={len(kz_miss)}")
    print(f"  RU matched={ru_match}/{len(acc_ru_subjects)}, missing={len(ru_miss)}")

    for s in kz_miss: print(f"    ❌ KZ: '{s}'")
    for s in ru_miss: print(f"    ❌ RU: '{s}'")

    if not kz_miss: print("  ✅ KZ: All subjects found in Excel!")
    if not ru_miss: print("  ✅ RU: All subjects found in Excel!")

    all_results[sheet] = {
        'kz_missing': kz_miss,
        'ru_missing': ru_miss,
        'excel_count': len(excel_subjects)
    }

# ─── Global summary ──────────────────────────────────────────────────────────
print("\n" + "=" * 70)
print("GLOBAL SUMMARY")
print("=" * 70)
total_issues = 0
for sheet, res in sorted(all_results.items()):
    kz_ok = '✅' if not res['kz_missing'] else f"❌ {len(res['kz_missing'])} missing"
    ru_ok = '✅' if not res['ru_missing'] else f"❌ {len(res['ru_missing'])} missing"
    total_issues += len(res['kz_missing']) + len(res['ru_missing'])
    print(f"  {sheet:8s}  KZ: {kz_ok:<25}  RU: {ru_ok}")

print(f"\nTotal unmatched subjects: {total_issues}")

with open('smart_verification_report.json', 'w', encoding='utf-8') as f:
    json.dump(all_results, f, ensure_ascii=False, indent=2)
print("Full report -> smart_verification_report.json")
