# -*- coding: utf-8 -*-
"""
Bridge Module
=============
Connects parser output (SUBJECT_COLUMNS / «руки») with document structure
(PROGRAM_PAGES / «мозг») for diploma generation.

Architecture:
    Excel ──(SUBJECT_COLUMNS)──► Parser ──► raw grades dict
                                                │
    PROGRAM_PAGES (Subject objects) ──► Bridge ◄─┘
                                          │
                                          ▼
                                structured_pages ──► Generator

SUBJECT_COLUMNS содержит индексы колонок (2, 6, 10...) — агент знает,
ОТКУДА парсить данные в исходной таблице ведомости.

PROGRAM_PAGES содержит Subject(name_kz=...) — объектная модель знает,
на какой странице и в каком порядке предмет стоит в дипломе.

META_COLUMNS позволяет агенту автоматически находить номер диплома и годы.
"""

import re
from typing import Dict, List, Any

from config.programs import PROGRAM_IT_PAGES, PROGRAM_ACCOUNTING_PAGES
from src.utils import normalize_key


# ─────────────────────────────────────────────────────────────
# Module header detection (replicates generator.is_module_header)
# ─────────────────────────────────────────────────────────────

_HEADER_PREFIXES = (
    "КМ", "ПМ",
    "Кәсіптік модуль",
    "Базовые модули", "Профессиональные модули",
    "Базалық модул", "Кәсіби модул",
)

# Practice indicators (excludes «практикалық» — lab/tutorial sessions)
_PRACTICE_KEYWORDS = [
    "оқу практика", "учебная практика",
    "өндірістік практика", "производственная практика",
    "преддипломная практика", "тәжірибе",
]


# ─────────────────────────────────────────────────────────────
# Internal helpers
# ─────────────────────────────────────────────────────────────

def _get_pages(program_code: str) -> dict:
    """Get PROGRAM_PAGES for given program code."""
    if program_code in ("3F", "IT"):
        return PROGRAM_IT_PAGES
    elif program_code in ("3D", "ACCOUNTING"):
        return PROGRAM_ACCOUNTING_PAGES
    raise ValueError(f"Unknown program_code: {program_code}")


def _is_header_by_name(name: str) -> bool:
    """Check if subject name looks like a module/section header."""
    if not name:
        return False
    s = name.strip()
    return any(s.startswith(p) for p in _HEADER_PREFIXES)


def _is_practice(name: str) -> bool:
    """Check if subject is a practice (not «практикалық»)."""
    if not name:
        return False
    s = name.lower()
    return any(p in s for p in _PRACTICE_KEYWORDS) and "практикалық" not in s


def _find_grade(subject, grades: dict) -> dict:
    """
    Find best-matching grade entry for a Subject in the parsed grades dict.

    Search order:
    1. Direct match by normalized KZ/RU name
    2. Prefix match (ОН 1.1, РО 2.3, БМ 1, etc.)
    3. Cross-language prefix (ОН↔РО, КМ↔ПМ)
    4. Practice keyword fallback
    """
    nkz = normalize_key(subject.name_kz)
    nru = normalize_key(subject.name_ru)

    # 1. Direct match
    g = grades.get(nkz) or grades.get(nru)
    if g:
        return g

    # 2–3. Prefix + cross-language
    for name in (subject.name_kz, subject.name_ru):
        m = re.match(
            r"((?:БМ|КМ|ПМ|ОН|РО)\s*\.?\s*\d+(?:\.\d+)?)",
            name, re.IGNORECASE,
        )
        if not m:
            continue
        prefix = normalize_key(m.group(1))
        for gk, gv in grades.items():
            if gk.startswith(prefix):
                return gv
        # Cross-language fallback
        alt = prefix
        if prefix.startswith("ро"):
            alt = "он" + prefix[2:]
        elif prefix.startswith("он"):
            alt = "ро" + prefix[2:]
        elif prefix.startswith("пм"):
            alt = "км" + prefix[2:]
        elif prefix.startswith("км"):
            alt = "пм" + prefix[2:]
        if alt != prefix:
            for gk, gv in grades.items():
                if gk.startswith(alt):
                    return gv

    # 4. Practice keyword fallback
    markers = ["кәсіптік практика", "профессиональная практика"]
    low_kz = subject.name_kz.lower()
    low_ru = subject.name_ru.lower()
    if any(pm in low_kz or pm in low_ru for pm in markers):
        for gv in grades.values():
            gv_kz = str(gv.get("subject_kz", "")).lower()
            gv_ru = str(gv.get("subject_ru", "")).lower()
            if any(pm in gv_kz or pm in gv_ru for pm in markers):
                return gv

    return {}


# ─────────────────────────────────────────────────────────────
# Public API
# ─────────────────────────────────────────────────────────────

def build_diploma_pages(
    grades: Dict[str, Any],
    program_code: str,
) -> Dict[int, List[Dict[str, Any]]]:
    """
    Bridge: map parsed grades onto PROGRAM_PAGES document structure.

    Combines two data sources:
      • SUBJECT_COLUMNS → parser → grades dict  («руки» — откуда читать)
      • PROGRAM_PAGES   → Subject objects        («мозг» — куда класть)

    Algorithm:
      1. Загружает PROGRAM_PAGES[program_code]
      2. Для каждого Subject ищет совпадение в grades dict
      3. Разрешает hours/credits (grade → PROGRAM_PAGES fallback)
      4. Проставляет флаги is_header / is_practice
      5. Возвращает готовую структуру для генератора

    Args:
        grades: normalized_key → grade_info dict (output of parser)
        program_code: '3F' (IT) or '3D' (Accounting)

    Returns:
        {page_num: [entry, ...]}
        Each entry dict:
            subject        — Subject object (name_kz/ru, hours, credits, flags)
            hours          — resolved hours  (grade → PROGRAM_PAGES fallback)
            credits        — resolved credits
            points         — score or ''
            letter         — letter grade or ''
            gpa            — GPA string or ''
            traditional_kz — казахская традиционная оценка
            traditional_ru — русская традиционная оценка
            is_header      — True → module/section header, no individual grade
            is_practice    — True → practice subject
    """
    pages = _get_pages(program_code)
    result: Dict[int, List[Dict[str, Any]]] = {}

    for page_num, subjects in pages.items():
        entries: List[Dict[str, Any]] = []

        for subject in subjects:
            grade = _find_grade(subject, grades)

            # Hours/credits: prefer parsed data, fallback to PROGRAM_PAGES
            hours = str(grade.get("hours", "")) if grade else ""
            if not hours or hours in ("0", "", "None"):
                hours = str(subject.hours) if subject.hours else ""

            credits_val = str(grade.get("credits", "")) if grade else ""
            if not credits_val or credits_val in ("0", "", "None"):
                credits_val = str(subject.credits) if subject.credits else ""

            # Header flag: Subject.is_module_header OR name-based detection
            is_header = (
                subject.is_module_header
                or _is_header_by_name(subject.name_kz)
                or _is_header_by_name(subject.name_ru)
            )

            entries.append({
                "subject": subject,
                "hours": hours,
                "credits": credits_val,
                "points": str(grade.get("points", "")) if grade else "",
                "letter": str(grade.get("letter", "")) if grade else "",
                "gpa": str(grade.get("gpa", "")) if grade else "",
                "traditional_kz": str(grade.get("traditional_kz", "")) if grade else "",
                "traditional_ru": str(grade.get("traditional_ru", "")) if grade else "",
                "is_header": is_header,
                "is_practice": (
                    _is_practice(subject.name_kz)
                    or _is_practice(subject.name_ru)
                ),
            })

        result[page_num] = entries

    return result
