"""
Smart verification v2:
- Excel contains bilingual cells (KZ\\nRU) — build BOTH kz and ru index
- IT KZ template → check against KZ excel keys
- IT RU template → check against RU excel keys
- Accountant 3D: read structure (row_idx=1 may differ — inspect first)
"""
import sys, os, re, json
sys.path.insert(0, r"c:\Users\user\OneDrive\Рабочий стол\template")
os.chdir(r"c:\Users\user\OneDrive\Рабочий стол\template")

import pandas as pd

from generate_diploma_it_kz import (
    PAGE1_SUBJECTS as IT_KZ_P1, PAGE2_SUBJECTS as IT_KZ_P2,
    PAGE3_SUBJECTS as IT_KZ_P3, PAGE4_SUBJECTS as IT_KZ_P4,
)
from generate_diploma_it_ru import (
    PAGE1_SUBJECTS as IT_RU_P1, PAGE2_SUBJECTS as IT_RU_P2,
    PAGE3_SUBJECTS as IT_RU_P3, PAGE4_SUBJECTS as IT_RU_P4,
)

IT_KZ_ALL = IT_KZ_P1 + IT_KZ_P2 + IT_KZ_P3 + IT_KZ_P4
IT_RU_ALL = IT_RU_P1 + IT_RU_P2 + IT_RU_P3 + IT_RU_P4

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"

# Row config (same for IT and accountants — will verify)
ROW_SUBJECT_NAMES = 1
ROW_HOURS         = 3
COL_START         = 2

def nkey(text):
    if not text: return ""
    t = str(text).lower().replace('.','').replace(',','').replace(':','').replace(' ','')
    t = re.sub(r'([a-zа-яёәіңғүұқөһ]+)0+([1-9]+)', r'\1\2', t)
    return t.strip()


def parse_hours_credits(text):
    if not isinstance(text, str): return "", ""
    m = re.search(r"(\d+)с-(\d+(?:,\d+)?)к", text)
    return (m.group(1), m.group(2)) if m else (text, "")


def get_excel_subject_keys(sheet_name):
    """
    Build two sets: kz_keys and ru_keys from the bilingual Excel column.
    Each subject cell looks like: 'KZ name\nRU name'
    Returns (kz_set, ru_set, subjects_list)
    """
    df = pd.read_excel(SOURCE_FILE, sheet_name=sheet_name, header=None)
    subj_row  = df.iloc[ROW_SUBJECT_NAMES]
    hours_row = df.iloc[ROW_HOURS]

    kz_keys = {}
    ru_keys = {}
    subjects_list = []

    col = COL_START
    while col < len(subj_row):
        cell_val = subj_row.iloc[col]
        if pd.isna(cell_val) or str(cell_val).strip() in ('', 'nan'):
            col += 4
            continue

        raw = str(cell_val).strip()
        parts = raw.split('\n')
        kz_name = parts[0].strip().rstrip(':') if parts else raw
        ru_name = parts[1].strip().rstrip(':') if len(parts) >= 2 else kz_name

        h_raw = hours_row.iloc[col] if col < len(hours_row) else ''
        hours, credits = parse_hours_credits(str(h_raw))

        kz_keys[nkey(kz_name)] = kz_name
        ru_keys[nkey(ru_name)] = ru_name
        subjects_list.append({'kz': kz_name, 'ru': ru_name, 'hours': hours, 'credits': credits})

        col += 4

    return kz_keys, ru_keys, subjects_list


def check_subjects(template_list, excel_keys, label, skip_endings=(':', '')):
    """Check template subjects against a set of excel normalized keys."""
    matched = []
    missing = []

    for subj in template_list:
        s = subj.strip()
        if not s or s.endswith(':'):
            continue

        nk = nkey(s)

        # Direct match
        if nk in excel_keys:
            matched.append(s)
            continue

        # BM prefix fallback: "БМ 1. Xxx" -> look for "бм1..."
        bm_m = re.match(r'(БМ\s*\.?\s*\d+)', s)
        if bm_m:
            prefix = nkey(bm_m.group(1))
            if any(k.startswith(prefix) for k in excel_keys):
                matched.append(f"{s} [BM prefix ✓]")
                continue

        # Module header (КМ/ПМ) — these summarize sub-subjects
        mod_m = re.match(r'^(КМ\s*\d+|ПМ\s*\d+)', s)
        if mod_m:
            prefix = nkey(mod_m.group(1))
            if any(k.startswith(prefix) for k in excel_keys):
                matched.append(f"{s} [module header ✓]")
                continue

        # Practice / Attestation (aggregate rows — OK if not in excel)
        if any(x in s for x in ['практика', 'Практика', 'аттестаттау', 'аттестация', 'Практика', 'Факультатив', 'Факультативтік']):
            matched.append(f"{s} [practice/attestation — aggregate row]")
            continue

        missing.append(s)

    return matched, missing


# ═══════════════════════════════════════════════════════════════════
print("=" * 72)
print("IT GROUPS (3F) vs EXCEL SUBJECT NAMES")
print("=" * 72)

IT_SHEETS = ["3Ғ-1", "3Ғ-2", "3Ғ-3", "3Ғ-4"]
it_summary = {}

for sheet in IT_SHEETS:
    print(f"\n── {sheet} ─────────────────────────────────")
    try:
        kz_keys, ru_keys, subjects = get_excel_subject_keys(sheet)
    except Exception as e:
        print(f"  ERROR: {e}")
        continue

    print(f"  Excel: {len(subjects)} subjects")

    kz_matched, kz_missing = check_subjects(IT_KZ_ALL, kz_keys, "KZ")
    ru_matched, ru_missing = check_subjects(IT_RU_ALL, ru_keys, "RU")

    if kz_missing:
        print(f"  ❌ KZ missing ({len(kz_missing)}):")
        for s in kz_missing: print(f"      '{s}'")
    else:
        print(f"  ✅ KZ: all {len(kz_matched)} subjects matched")

    if ru_missing:
        print(f"  ❌ RU missing ({len(ru_missing)}):")
        for s in ru_missing: print(f"      '{s}'")
    else:
        print(f"  ✅ RU: all {len(ru_matched)} subjects matched")

    it_summary[sheet] = {'kz': kz_missing, 'ru': ru_missing}

# ═══════════════════════════════════════════════════════════════════
print("\n" + "=" * 72)
print("ACCOUNTANT GROUPS (3D) — Reading structure first")
print("=" * 72)

ACC_SHEETS = ["3D-1", "3D-2"]

# Inspect accountant row structure first
for sheet in ACC_SHEETS:
    df = pd.read_excel(SOURCE_FILE, sheet_name=sheet, header=None)
    print(f"\n── {sheet} ─────")
    print(f"  Shape: {df.shape}")
    # Show first 5 rows of col B onwards (first 6 cols)
    for ri in range(min(6, len(df))):
        vals = [str(df.iloc[ri, ci])[:40] if ci < len(df.columns) else '' for ci in range(1, 7)]
        print(f"  Row {ri}: {vals}")

print("\n[Using get_excel_subject_keys for 3D with same structure...]")

acc_summary = {}
for sheet in ACC_SHEETS:
    print(f"\n── {sheet} ─────────────────────────────────")
    try:
        kz_keys, ru_keys, subjects = get_excel_subject_keys(sheet)
        print(f"  Excel: {len(subjects)} subjects found")
        print("  KZ subjects in Excel:")
        for k, v in list(kz_keys.items())[:30]:
            print(f"    • {v}")
    except Exception as e:
        print(f"  ERROR: {e}")

print("\n" + "=" * 72)
print("GLOBAL SUMMARY")
print("=" * 72)
all_ok = True
for sheet in IT_SHEETS:
    r = it_summary.get(sheet, {})
    ok_kz = "✅" if not r.get('kz') else f"❌ {len(r['kz'])} missing"
    ok_ru = "✅" if not r.get('ru') else f"❌ {len(r['ru'])} missing"
    all_ok = all_ok and not r.get('kz') and not r.get('ru')
    print(f"  {sheet}  KZ: {ok_kz:<30} RU: {ok_ru}")

print()
if all_ok:
    print("🎉 ALL IT subjects match Excel data perfectly!")
else:
    print("⚠️  Some subjects don't match — see details above")
