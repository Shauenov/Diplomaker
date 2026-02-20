"""Save full Excel subjects list to JSON for analysis."""
import sys, os, re, json
sys.path.insert(0, r"c:\Users\user\OneDrive\Рабочий стол\template")
os.chdir(r"c:\Users\user\OneDrive\Рабочий стол\template")
import pandas as pd

SOURCE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEETS = ["3Ғ-1", "3Ғ-2", "3Ғ-3", "3Ғ-4", "3D-1", "3D-2"]

def nkey(text):
    if not text: return ""
    t = str(text).lower().replace('.','').replace(',','').replace(':','').replace(' ','')
    t = re.sub(r'([a-zа-яёәіңғүұқөһ]+)0+([1-9]+)', r'\1\2', t)
    return t.strip()

def parse_hc(text):
    if not isinstance(text, str): return "", ""
    m = re.search(r"(\d+)с-(\d+(?:,\d+)?)к", text)
    return (m.group(1), m.group(2)) if m else ("","")

result = {}
for sheet in SHEETS:
    df = pd.read_excel(SOURCE, sheet_name=sheet, header=None)
    sr = df.iloc[1]
    hr = df.iloc[3]
    subjects = []
    col = 2
    while col < len(sr):
        cv = sr.iloc[col]
        if pd.isna(cv) or str(cv).strip() in ('','nan'):
            col += 4; continue
        raw = str(cv).strip()
        parts = raw.split('\n')
        kz = parts[0].strip().rstrip(':')
        ru = parts[1].strip().rstrip(':') if len(parts) >= 2 else kz
        hraw = str(hr.iloc[col]) if col < len(hr) else ''
        h, c = parse_hc(hraw)
        subjects.append({'kz': kz, 'ru': ru, 'nkz': nkey(kz), 'nru': nkey(ru), 'hours': h, 'credits': c})
        col += 4
    result[sheet] = subjects

with open('excel_subjects_all.json', 'w', encoding='utf-8') as f:
    json.dump(result, f, ensure_ascii=False, indent=2)

print("Saved. Counts:")
for s, v in result.items():
    print(f"  {s}: {len(v)}")
