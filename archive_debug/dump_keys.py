"""Dump all DEBUG grades data keys to find correct BM names."""
import json
d = json.load(open('debug_grades_3F-1.json', encoding='utf-8'))
print(f"Total keys: {len(d)}")
print()
for k, v in d.items():
    line = f"  {repr(k)[:80]}: h={v.get('hours','?')}, k={v.get('credits','?')}"
    print(line)
