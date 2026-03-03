# -*- coding: utf-8 -*-
"""
Utility Functions
=================
Shared utility functions for text processing, normalization, and parsing.

Functions:
----------
- is_module_header(): Detect module/section headers (КМ, БМ, etc.)
- normalize_key(): Normalize text for fuzzy matching
- clean_subject_name(): Parse bilingual subject names
- parse_hours_credits(): Parse "72с-3к" format into hours and credits
- robust_clean(): Clean cell values (handle NaN, errors, empty values)
- format_float_value(): Format float values for display
"""

import re
import pandas as pd
from typing import Tuple, Optional


def is_module_header(subject_name: str) -> bool:
    """
    Check if a subject name is a module header.
    
    Module headers include:
    - КМ (Кәсіптік модуль / Professional Module)
    - БМ (Базалық модуль / Base Module)
    - ПМ (Профессионалдық модуль / Professional Module)
    - СМ (Специализированный модуль / Specialized Module)
    - Кәсіптік практика / Профессиональная практика (Professional Practice)
    - Аттестация / Аттестаттау (Attestation)
    - Зачет (Pass/Fail)
    
    Args:
        subject_name: Subject name to check
        
    Returns:
        True if subject is a module header, False otherwise
        
    Examples:
        >>> is_module_header("КМ 01 Web технологиялар")
        True
        >>> is_module_header("БМ 02 Математика негіздері")
        True
        >>> is_module_header("Қазақ тілі")
        False
    """
    if not subject_name:
        return False
    
    s = subject_name.strip().lower()
    
    # Check for module prefixes
    if (s.startswith("км ") or s.startswith("км0") or s.startswith("км_") or
        s.startswith("бм ") or s.startswith("бм0") or s.startswith("бм_") or
        s.startswith("пм ") or s.startswith("пм0") or s.startswith("пм_") or
        s.startswith("см ") or s.startswith("см0") or s.startswith("см_")):
        return True
    
    # Check for practice/attestation keywords
    if ("практика" in s or 
        "аттеста" in s or 
        "зачет" in s or
        "зачёт" in s):
        return True
    
    return False


def normalize_key(text: str) -> str:
    """
    Normalize a subject name to a compact lowercase key for fuzzy matching.
    
    Normalization steps:
    1. Convert to lowercase
    2. Remove punctuation (., ,, :)
    3. Remove all spaces
    4. Remove leading zeros from numbers (e.g., "км01" → "км1")
    
    Args:
        text: Text to normalize
        
    Returns:
        Normalized key string
        
    Examples:
        >>> normalize_key("Қазақ тілі: деңгейлік курс")
        "қазақтілідеңгейліккурс"
        >>> normalize_key("КМ 01 Web технологиялар")
        "км1webтехнологиялар"
        >>> normalize_key("Front-end Web ресурстарды құру")
        "front-endwebресурстардықұру"
    """
    if not text:
        return ""
    
    t = str(text).lower()
    
    # Remove punctuation
    t = t.replace(".", "").replace(",", "").replace(":", "")
    
    # Remove spaces
    t = t.replace(" ", "")
    
    # Remove leading zeros from numbers after letters
    # e.g., "км01" → "км1", "он03" → "он3"
    t = re.sub(r'([a-zа-яәіңғүұқөһё]+)0+([1-9]+)', r'\1\2', t)
    
    return t.strip()


def clean_subject_name(text: str) -> Tuple[str, str]:
    """
    Parse bilingual subject names into (Kazakh, Russian) tuple.
    
    Subject names in source Excel are typically formatted as:
    "Kazakh Name\nRussian Name"
    
    This function:
    1. Splits on newline
    2. Strips whitespace
    3. Removes trailing colons
    
    Args:
        text: Bilingual subject name (newline-separated)
        
    Returns:
        Tuple of (kazakh_name, russian_name)
        
    Examples:
        >>> clean_subject_name("Қазақ тілі\nКазахский язык")
        ("Қазақ тілі", "Казахский язык")
        >>> clean_subject_name("Математика:\nМатематика:")
        ("Математика", "Математика")
        >>> clean_subject_name("Single line")
        ("Single line", "Single line")
    """
    if not isinstance(text, str):
        cleaned = str(text).strip()
        return cleaned, cleaned
    
    # Split on newline
    parts = text.split('\n')
    
    if len(parts) >= 2:
        # Bilingual format
        kz = parts[0].strip().rstrip(':').strip()
        ru = parts[1].strip().rstrip(':').strip()
        return kz, ru
    else:
        # Single language or no newline
        cleaned = text.strip().rstrip(':').strip()
        return cleaned, cleaned


def parse_hours_credits(text: str) -> Tuple[str, str]:
    """
    Parse hours and credits from format "72с-3к" or "108с-4.5к".
    
    Expected format: "{hours}с-{credits}к"
    - с (сағат / часы) = hours
    - к (кредит / кредиты) = credits
    
    Args:
        text: Hours/credits string
        
    Returns:
        Tuple of (hours, credits) as strings. Returns ("", "") if invalid.
        
    Examples:
        >>> parse_hours_credits("72с-3к")
        ("72", "3")
        >>> parse_hours_credits("108с-4.5к")
        ("108", "4.5")
        >>> parse_hours_credits("invalid")
        ("invalid", "")
        >>> parse_hours_credits("NaN")
        ("", "")
    """
    if not isinstance(text, str) or text.lower() == "nan" or not text.strip():
        return "", ""
    
    # Try to match pattern: {digits}с-{digits/decimal}к
    match = re.search(r"(\d+)с-(\d+(?:[.,]\d+)?)к", text)
    if match:
        hours = match.group(1)
        credits = match.group(2).replace(',', '.')  # Normalize comma to dot
        return hours, credits
    
    # If no match, return text as hours, empty credits
    return text, ""


def robust_clean(val) -> str:
    """
    Convert a cell value to a clean string, discarding invalid values.
    
    Invalid values that return empty string:
    - NaN / None
    - "nan" (string)
    - "0" (zero)
    - "#REF!" (Excel error)
    - Empty strings
    
    Args:
        val: Cell value (any type)
        
    Returns:
        Cleaned string value or empty string
        
    Examples:
        >>> robust_clean(85)
        "85"
        >>> robust_clean("A+")
        "A+"
        >>> robust_clean(float('nan'))
        ""
        >>> robust_clean("#REF!")
        ""
        >>> robust_clean(0)
        ""
    """
    # Check for pandas NaN
    if pd.isna(val):
        return ""
    
    # Convert to string for comparison
    str_val = str(val).strip()
    
    # Check for invalid values
    if (str_val.lower() == "nan" or 
        str_val == "#REF!" or
        str_val == ""):
        return ""
    
    return str_val


def format_float_value(value: Optional[float], decimal_places: int = 1) -> str:
    """
    Format a float value with specified decimal places.
    
    Args:
        value: Float value to format (or None)
        decimal_places: Number of decimal places
        
    Returns:
        Formatted string or empty string if None
        
    Examples:
        >>> format_float_value(4.0, 1)
        "4.0"
        >>> format_float_value(3.67, 2)
        "3.67"
        >>> format_float_value(None, 1)
        ""
    """
    if value is None:
        return ""
    
    return f"{value:.{decimal_places}f}"
