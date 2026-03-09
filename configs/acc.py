# -*- coding: utf-8 -*-
"""Accounting program config — derived from centralized config/programs.py"""

from config.programs import PROGRAM_ACCOUNTING_PAGES

# Автоматически строим списки предметов по страницам из единого источника
def _build_page_subjects(pages_dict, lang_attr):
    """Извлекает списки имен предметов из Subject-объектов programs.py."""
    result = {}
    for page_num, subjects in pages_dict.items():
        result[f'p{page_num}'] = [getattr(s, lang_attr) for s in subjects]
    return result

# Термины для таблицы (зачет/отлично)
TERMS = {
    'kz': {
        'traditional_elective': 'сынақ',
        'traditional_practice': 'сынақ',
        'traditional_attestation': 'өте жақсы',
    },
    'ru': {
        'traditional_elective': 'зачтено',
        'traditional_practice': 'зачтено',
        'traditional_attestation': 'отлично',
    }
}

ACC_CONFIG = {
    'kz': _build_page_subjects(PROGRAM_ACCOUNTING_PAGES, 'name_kz'),
    'ru': _build_page_subjects(PROGRAM_ACCOUNTING_PAGES, 'name_ru'),
}
