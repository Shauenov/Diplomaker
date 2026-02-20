"""
Final smart verification:
Excel structure:
  - IT (3F): 26 entries = общеобразовательные предметы + БМ + КМ (агрегированные) + Практика + Аттестация
  - 3D: 16 entries = общеобразовательные + Базалық модульдер (агрегировано) + Кәсіптік модульдер (агрегировано)

Verification logic:
  1. Базовые предметы (Казахский язык, математика, etc.) - должны совпадать 1:1
  2. БМ предметы - должны совпадать по prefix
  3. КМ/ПМ заголовки - в Excel они есть как аgg строки (OK - aggregate rows)
  4. ОН-предметы - НЕТ в Excel по отдельности (это детальные sub-items КМ) - нормально
  5. Практика + Аттестация - агрегированные строки - OK

Output: comprehensive report in JSON + text
"""
import sys, os, re, json
sys.path.insert(0, r"c:\Users\user\OneDrive\Рабочий стол\template")
os.chdir(r"c:\Users\user\OneDrive\Рабочий стол\template")

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

# Load the saved excel subjects
with open('excel_subjects_all.json', encoding='utf-8') as f:
    excel_data = json.load(f)

def nkey(text):
    if not text: return ""
    t = str(text).lower().replace('.','').replace(',','').replace(':','').replace(' ','')
    t = re.sub(r'([a-zа-яёәіңғүұқөһ]+)0+([1-9]+)', r'\1\2', t)
    return t.strip()

REPORT = {}

def classify_and_check(template_subjects, excel_entries, lang_key='kz'):
    """
    Classify each template subject and check against excel.
    Returns dict with counts and lists.
    """
    kz_set = {nkey(e[lang_key]): e for e in excel_entries}
    
    results = {
        'matched_basic': [],
        'matched_bm': [],
        'matched_module_header': [],
        'skipped_on_subitem': [],
        'skipped_practice_attestation': [],
        'skipped_elective': [],
        'MISSING': [],
    }
    
    for subj in template_subjects:
        s = subj.strip()
        if not s or s.endswith(':') or s == '':
            continue
        
        nk = nkey(s)
        
        # 1. Direct match (basic subjects, БМ full name, etc.)
        if nk in kz_set:
            entry = kz_set[nk]
            results['matched_basic'].append(f"{s} [h={entry['hours']}, k={entry['credits']}]")
            continue
        
        # 2. БМ prefix match
        bm_match = re.match(r'(БМ\s*\.?\s*\d+)', s)
        if bm_match:
            prefix = nkey(bm_match.group(1))
            found = next(((nk2, e) for nk2, e in kz_set.items() if nk2.startswith(prefix)), None)
            if found:
                nk2, e = found
                results['matched_bm'].append(f"{s} → Excel: '{e[lang_key]}' [h={e['hours']}, k={e['credits']}]")
                continue
        
        # 3. КМ/ПМ module header — check aggregate in excel
        mod_match = re.match(r'^(КМ\s*\d+|ПМ\s*\d+)', s)
        if mod_match:
            prefix = nkey(mod_match.group(1))
            found = next(((nk2, e) for nk2, e in kz_set.items() if nk2.startswith(prefix)), None)
            if found:
                nk2, e = found
                results['matched_module_header'].append(f"{s} → Excel aggregate: '{e[lang_key]}' [h={e['hours']}, k={e['credits']}]")
                continue
            else:
                results['MISSING'].append(f"{s} ← КМ/ПМ header NOT in Excel")
                continue
        
        # 4. ОН sub-items — these are expected to NOT be in Excel (they're detailed sub-subjects)
        if re.match(r'^(ОН|РО)\s*\d+\.\d+', s):
            results['skipped_on_subitem'].append(s)
            continue
        
        # 5. Practice / Attestation / Electives — aggregate or special
        if any(x in s for x in ['практика', 'Практика', 'аттестат', 'Аттестат']):
            results['skipped_practice_attestation'].append(s)
            continue
        if any(x in s for x in ['Факультатив', 'Факультативтік']):
            results['skipped_elective'].append(s)
            continue
        
        # 6. Truly missing
        results['MISSING'].append(s)
    
    return results


print("=" * 72)
print("COMPREHENSIVE SUBJECT VERIFICATION REPORT")
print("=" * 72)

IT_SHEETS = ["3Ғ-1", "3Ғ-2", "3Ғ-3", "3Ғ-4"]
all_missing_kz = {}
all_missing_ru = {}

for sheet in IT_SHEETS:
    entries = excel_data[sheet]
    print(f"\n{'─'*72}")
    print(f"Sheet: {sheet} ({len(entries)} Excel subjects)")
    print(f"{'─'*72}")
    
    # KZ check
    kz_res = classify_and_check(IT_KZ_ALL, entries, 'kz')
    print(f"\n  [KZ Template]")
    print(f"  ✅ Matched basic:         {len(kz_res['matched_basic'])}")
    print(f"  ✅ Matched BM:            {len(kz_res['matched_bm'])}")
    print(f"  ✅ Matched КМ/ПМ header:  {len(kz_res['matched_module_header'])}")
    print(f"  ⏭  ОН sub-items (normal): {len(kz_res['skipped_on_subitem'])}")
    print(f"  ⏭  Practice/Attestation:  {len(kz_res['skipped_practice_attestation'])}")
    print(f"  ⏭  Electives:             {len(kz_res['skipped_elective'])}")
    if kz_res['MISSING']:
        print(f"  ❌ MISSING ({len(kz_res['MISSING'])}):")
        for s in kz_res['MISSING']: print(f"      - {s}")
    else:
        print(f"  ✅ NO MISSING subjects!")
    all_missing_kz[sheet] = kz_res['MISSING']
    
    # RU check (using ru key from Excel)
    ru_res = classify_and_check(IT_RU_ALL, entries, 'ru')
    print(f"\n  [RU Template]")
    print(f"  ✅ Matched basic:         {len(ru_res['matched_basic'])}")
    print(f"  ✅ Matched BM:            {len(ru_res['matched_bm'])}")
    print(f"  ✅ Matched КМ/ПМ header:  {len(ru_res['matched_module_header'])}")
    print(f"  ⏭  ОН sub-items (normal): {len(ru_res['skipped_on_subitem'])}")
    print(f"  ⏭  Practice/Attestation:  {len(ru_res['skipped_practice_attestation'])}")
    print(f"  ⏭  Electives:             {len(ru_res['skipped_elective'])}")
    if ru_res['MISSING']:
        print(f"  ❌ MISSING ({len(ru_res['MISSING'])}):")
        for s in ru_res['MISSING']: print(f"      - {s}")
    else:
        print(f"  ✅ NO MISSING subjects!")
    all_missing_ru[sheet] = ru_res['MISSING']

# ─── Accountants ─────────────────────────────────────────────────────
print(f"\n{'='*72}")
print("ACCOUNTANT GROUPS (3D)")
print(f"{'='*72}")

ACC_SHEETS = ["3D-1", "3D-2"]
for sheet in ACC_SHEETS:
    entries = excel_data[sheet]
    print(f"\n  {sheet}: {len(entries)} subjects in Excel")
    print("  Excel subjects (KZ):")
    for e in entries:
        h = e['hours']; c = e['credits']
        print(f"    {'✅' if h else '⚠ '} {e['kz']:60s}  h={h}, k={c}")
    print("  Excel subjects (RU):")
    for e in entries:
        h = e['hours']; c = e['credits']
        print(f"    {'✅' if h else '⚠ '} {e['ru']:60s}  h={h}, k={c}")

# ─── Final Summary ───────────────────────────────────────────────────
print(f"\n{'='*72}")
print("FINAL SUMMARY")  
print(f"{'='*72}")

total_issues = sum(len(v) for v in all_missing_kz.values()) + sum(len(v) for v in all_missing_ru.values())

for sheet in IT_SHEETS:
    kz_ok = "✅ OK" if not all_missing_kz.get(sheet) else f"❌ {len(all_missing_kz[sheet])} missing"
    ru_ok = "✅ OK" if not all_missing_ru.get(sheet) else f"❌ {len(all_missing_ru[sheet])} missing"
    print(f"  {sheet}   KZ: {kz_ok:<20}  RU: {ru_ok}")

for sheet in ACC_SHEETS:
    print(f"  {sheet}   (3D structure uses aggregate module rows — no individual subject check)")

print(f"\nTotal truly missing subjects: {total_issues}")

# Save full details
full_report = {
    'it_missing_kz': all_missing_kz,
    'it_missing_ru': all_missing_ru,
    'summary': 'ОН sub-items are NOT stored individually in Excel (expected). Only basic+BM+KM headers verified.'
}
with open('final_verification_report.json', 'w', encoding='utf-8') as f:
    json.dump(full_report, f, ensure_ascii=False, indent=2)
print("Saved → final_verification_report.json")
