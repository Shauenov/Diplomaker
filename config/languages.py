# -*- coding: utf-8 -*-
"""
Language Configuration
======================
Language-specific text, grade names, and translations.

Supported languages:
- KZ (Kazakh / Қазақша)
- RU (Russian / Русский)

TODO: Extensible for Turkish, English, etc.
"""

# ─────────────────────────────────────────────────────────────
# LANGUAGE DEFINITIONS
# ─────────────────────────────────────────────────────────────

LANGUAGES = {
    "KZ": {
        "name": "Казахский",
        "native_name": "Қазақша",
        "code": "kk",
    },
    "RU": {
        "name": "Русский",
        "native_name": "Русский",
        "code": "ru",
    },
}

LANGUAGE_NAMES = {
    "KZ": "Қазақша",
    "RU": "Русский",
}

# ─────────────────────────────────────────────────────────────
# TRADITIONAL GRADE TRANSLATIONS
# ─────────────────────────────────────────────────────────────

TRADITIONAL_GRADES = {
    "KZ": {
        5: "5 (өте жақсы)",           # Excellent / Very Good
        4: "4 (жақсы)",               # Good
        3: "3 (қанағат)",             # Satisfactory
        2: "2 (қанағаттанарлықсыз)",  # Unsatisfactory
    },
    "RU": {
        5: "5 (отлично)",    # Excellent
        4: "4 (хорошо)",     # Good
        3: "3 (удовл)",      # Satisfactory
        2: "2 (неуд)",       # Unsatisfactory
    },
}

# ─────────────────────────────────────────────────────────────
# ELECTIVE PASSING GRADES
# ─────────────────────────────────────────────────────────────
# Electives show pass/fail, not numeric grades

ELECTIVE_GRADES = {
    "KZ": "сынақ",      # Test/Exam (pass)
    "RU": "зачтено",    # Credited (pass)
}

# ─────────────────────────────────────────────────────────────
# DIPLOMA TEXT FIELDS (for future UI/PDF generation)
# ─────────────────────────────────────────────────────────────

DIPLOMA_LABELS = {
    "KZ": {
        "diploma_id": "Диплом №",
        "full_name": "Студенттің аты-жөні",
        "program": "Мамандық",
        "specialization": "Маманданамасы",
        "qualification": "Квалификация",
        "year_range": "Оку жылы",
        "subjects": "Пәндер",
        "hours": "Сағат",
        "credits": "Кредит",
        "points": "Балл",
        "letter_grade": "Әріп баға",
        "gpa": "GPA",
        "traditional_grade": "Дәстүрлі баға",
    },
    "RU": {
        "diploma_id": "Диплом №",
        "full_name": "ФИ студента",
        "program": "Специальность",
        "specialization": "Специализация",
        "qualification": "Квалификация",
        "year_range": "Учебный год",
        "subjects": "Предметы",
        "hours": "Часы",
        "credits": "Кредит",
        "points": "Баллы",
        "letter_grade": "Оценка",
        "gpa": "GPA",
        "traditional_grade": "Традиционная оценка",
    },
}
