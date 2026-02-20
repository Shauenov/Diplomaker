# -*- coding: utf-8 -*-
"""
Config Package
==============
Centralized configuration management for diploma automation system.

Modules:
--------
- settings: Global constants, paths, grade thresholds
- languages: Multilingual text and grade mappings (KZ, RU)
- programs: Program-specific definitions (IT 3F, Accounting 3D)
"""

from .settings import *
from .languages import *
from .programs import *

__all__ = [
    # From settings
    "SOURCE_FILE",
    "OUTPUT_DIR",
    "ROW_SUBJECT_NAMES",
    "ROW_HOURS",
    "ROW_DATA_START",
    "COL_NO",
    "COL_FULL_NAME",
    "COL_START_SUBJECTS",
    "GRADE_THRESHOLDS",
    "INSTITUTION_NAME_KZ",
    "INSTITUTION_FULL_KZ",
    "INSTITUTION_NAME_RU",
    "INSTITUTION_FULL_RU",
    "PROGRAM_IT_NAME_KZ",
    "PROGRAM_IT_NAME_RU",
    "PROGRAM_ACCOUNTING_NAME_KZ",
    "PROGRAM_ACCOUNTING_NAME_RU",
    
    # From languages
    "LANGUAGES",
    "LANGUAGE_NAMES",
    "TRADITIONAL_GRADES",
    
    # From programs
    "PROGRAMS",
    "get_program_config",
    "get_sheets_for_program",
]
