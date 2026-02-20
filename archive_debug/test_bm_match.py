"""Test BM subject matching via normalize_key()."""
import json, re

d = json.load(open('debug_grades_3F-1.json', encoding='utf-8'))

def normalize_key(text):
    if not text: return ''
    t = text.lower()
    t = t.replace('.', '').replace(',', '').replace(':', '')
    t = t.replace(' ', '')
    t = re.sub(r'([a-z\u0430-\u044f]+)0+([1-9]+)', r'\1\2', t)
    return t.strip()

grades_map = {normalize_key(k): v for k, v in d.items()}

bm_keys = [(k, k) for k in grades_map.keys() if '\u0431\u043c' in k]
print('Normalized keys with bm:')
for k, _ in bm_keys[:10]:
    print(f'  {k}')

bm_subjects_ru = [
    '\u0411\u041c 01 \u0420\u0430\u0437\u0432\u0438\u0442\u0438\u0435 \u0438 \u0441\u043e\u0432\u0435\u0440\u0448\u0435\u043d\u0441\u0442\u0432. \u0444\u0438\u0437\u0438\u0447\u0435\u0441\u043a\u0438\u0445 \u043a\u0430\u0447\u0435\u0441\u0442\u0432',
    '\u0411\u041c 02 \u041f\u0440\u0438\u043c\u0435\u043d\u0435\u043d\u0438\u0435 \u0438\u043d\u0444\u043e\u0440\u043c\u0430\u0446\u0438\u043e\u043d\u043d\u043e-\u043a\u043e\u043c\u043c\u0443\u043d\u0438\u043a\u0430\u0446\u0438\u043e\u043d\u043d\u044b\u0445 \u0438 \u0446\u0438\u0444\u0440\u043e\u0432\u044b\u0445 \u0442\u0435\u0445\u043d\u043e\u043b\u043e\u0433\u0438\u0439',
    '\u0411\u041c 03 \u041f\u0440\u0438\u043c\u0435\u043d\u0435\u043d\u0438\u0435 \u0431\u0430\u0437\u043e\u0432\u044b\u0445 \u0437\u043d\u0430\u043d\u0438\u0439 \u044d\u043a\u043e\u043d\u043e\u043c\u0438\u043a\u0438 \u0438 \u043e\u0441\u043d\u043e\u0432', 
    '\u0411\u041c 04 \u041f\u0440\u0438\u043c\u0435\u043d\u0435\u043d\u0438\u0435 \u043e\u0441\u043d\u043e\u0432 \u0441\u043e\u0446\u0438\u0430\u043b\u044c\u043d\u044b\u0445 \u043d\u0430\u0443\u043a \u0434\u043b\u044f \u0441\u043e\u0446\u0438\u0430\u043b\u0438\u0437\u0430\u0446\u0438\u0438',
]

print('\nMatching test:')
for subj in bm_subjects_ru:
    nk = normalize_key(subj)
    match = grades_map.get(nk)
    hours = match.get('hours') if match else 'N/A'
    print(f'  Subject: {subj[:55]}')
    print(f'  Norm:    {nk[:40]}')
    print(f'  Match:   {bool(match)} hours={hours}')
    print()
