import re
import math
import pandas as pd

def normalize_key(title: str) -> str:
    """Очищает строку для поиска (убирает пробелы, точки, запятые, кейс)."""
    if not title or pd.isna(title):
        return ""
    
    t = str(title).strip().lower()
    t = t.replace(".", "").replace(",", "").replace(":", "").replace("-", "")
    t = t.replace(" ", "").replace("\n", "")
    t = t.replace("і", "i").replace("ң", "н").replace("қ", "к").replace("ғ", "г")
    t = t.replace("ү", "у").replace("ұ", "у").replace("ө", "о").replace("ә", "а").replace("ё", "е")
    
    # "км01" -> "км1", "пм05" -> "пм5"
    t = re.sub(r'([a-zа-я])0+([1-9])', r'\1\2', t)
    return t

def parse_hours_credits(val: str) -> tuple[str, str]:
    """Извлекает часы и кредиты из строки вида '72с-3к' или '90с-2,5к'."""
    if not isinstance(val, str):
        return "", ""
    # Ищем 'число'+'с'+'-'+'число(с запятой/точкой)'+'к'
    m = re.search(r"(\d+)\s*с\s*-\s*(\d+(?:[.,]\d+)?)\s*к", val, re.IGNORECASE)
    if m:
        h = m.group(1).replace(',', '.')
        c = m.group(2).replace(',', '.')
        return h, c
    
    # Альтернативный парсинг для '72с' 
    m2 = re.search(r"(\d+)\s*с", val, re.IGNORECASE)
    if m2:
        return m2.group(1), ""
        
    return "", ""

def safe_str(val) -> str:
    """Безопасное преобразование значения ячейки в строку."""
    if pd.isna(val):
        return ""
    v = str(val).strip()
    if v.lower() == 'nan':
        return ""
    
    # Удаляем \.0 для целых чисел, например "5.0" -> "5"
    if v.endswith('.0') and v.replace('.0', '').isdigit():
        return v[:-2]
        
    return v

def clean_name(full_name: str) -> str:
    """Очищает ФИО от лишних пробелов."""
    if not full_name: return ""
    return re.sub(r'\s+', ' ', str(full_name).strip())

# ─────────────────────────────────────────────────────────────
# Grade conversion functions moved to core/converters.py
# which uses centralized GRADE_THRESHOLDS from config/settings.py
# ─────────────────────────────────────────────────────────────
