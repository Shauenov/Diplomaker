"""Show ALL БМ keys from debug data."""
import json
d = json.load(open('debug_grades_3F-1.json', encoding='utf-8'))
print("All BM keys:")
for k, v in d.items():
    if k.startswith('\u0411\u041c'):
        print(f"  {repr(k)} -> hours={v.get('hours')}, credits={v.get('credits')}")
