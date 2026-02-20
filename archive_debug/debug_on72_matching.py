# -*- coding: utf-8 -*-
"""
debug_on72_matching.py
Debug why ОН 7.2 from template doesn't match source
"""

import re

def normalize_key(text):
    if not text:
        return ""
    t = str(text).lower()
    t = t.replace(".", "").replace(",", "").replace(":", "")
    t = t.replace(" ", "")
    t = re.sub(r'([a-zа-я]+)0+([1-9]+)', r'\1\2', t)
    return t.strip()

# Template version
template = "ОН 7.2 Деректер базасымен жұмыс істеу технологиясын пайдалану"

# Source version
source = "ОН7.2 Деректер базасымен жұмыс істеу технологиясын пайдалану."

template_norm = normalize_key(template)
source_norm = normalize_key(source)

print("Template:")
print(f"  Raw: {template}")
print(f"  Normalized: {template_norm}")
print()

print("Source:")
print(f"  Raw: {source}")
print(f"  Normalized: {source_norm}")
print()

if template_norm == source_norm:
    print("✅ They MATCH after normalization!")
else:
    print("❌ They DON'T match!")
    print(f"Diff: template={repr(template_norm)} vs source={repr(source_norm)}")
