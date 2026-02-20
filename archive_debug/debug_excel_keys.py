"""
Debug: show what Excel actually has vs what template subjects look like
"""
import sys, os, re
sys.path.insert(0, r"c:\Users\user\OneDrive\Рабочий стол\template")
os.chdir(r"c:\Users\user\OneDrive\Рабочий стол\template")
import pandas as pd

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"

def nkey(text):
    if not text: return ""
    t = str(text).lower().replace('.','').replace(',','').replace(':','').replace(' ','')
    t = re.sub(r'([a-zа-яёәіңғүұқөһ]+)0+([1-9]+)', r'\1\2', t)
    return t.strip()

def parse_hours_credits(text):
    if not isinstance(text, str): return "", ""
    m = re.search(r"(\d+)с-(\d+(?:,\d+)?)к", text)
    return (m.group(1), m.group(2)) if m else ("","")

# Read IT sheet
sheet = "3Ғ-1"
df = pd.read_excel(SOURCE_FILE, sheet_name=sheet, header=None)

print(f"Sheet {sheet}: {df.shape}")
subj_row  = df.iloc[1]   # row index 1 = Excel row 2
hours_row = df.iloc[3]   # row index 3 = Excel row 4

print(f"\n=== ALL EXCEL SUBJECTS in {sheet} ===")
col = 2
idx = 1
while col < len(subj_row):
    cell_val = subj_row.iloc[col]
    if pd.isna(cell_val) or str(cell_val).strip() in ('', 'nan'):
        col += 4
        continue
    raw = str(cell_val).strip()
    parts = raw.split('\n')
    kz = parts[0].strip().rstrip(':') if parts else raw
    ru = parts[1].strip().rstrip(':') if len(parts) >= 2 else ''
    h_raw = str(hours_row.iloc[col]) if col < len(hours_row) else ''
    hours, credits = parse_hours_credits(h_raw)
    nkz = nkey(kz)
    nru = nkey(ru)
    print(f"  {idx:3}. [{col}] H={hours},K={credits}")
    print(f"       KZ: '{kz}'")
    print(f"           key: {nkz}")
    if ru and ru != kz:
        print(f"       RU: '{ru}'")
        print(f"           key: {nru}")
    idx += 1
    col += 4

print(f"\nTotal: {idx-1} subjects in Excel")
