# -*- coding: utf-8 -*-
"""
Global Settings & Constants
============================
Centralized configuration for diploma automation system.

This module contains all hardcoded values that were previously scattered
across 25+ scripts. Update these values to configure the system for different
academic years, institutions, or grading policies.
"""

from pathlib import Path

# ─────────────────────────────────────────────────────────────
# FILE PATHS
# ─────────────────────────────────────────────────────────────

# Base workspace path
WORKSPACE_ROOT = Path(__file__).parent.parent

# Source Excel file with student grades
# Format: 2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx
SOURCE_FILE = (
    r"c:\Users\user\OneDrive\Рабочий стол\template\2025-2026 диплом бағалары "
    r"(ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
)

# Output directory for generated diplomas
OUTPUT_DIR = WORKSPACE_ROOT / "output_diplomas"

# Test data file
TEST_GRADES_FILE = WORKSPACE_ROOT / "test_grades_filled.xlsx"

# ─────────────────────────────────────────────────────────────
# EXCEL STRUCTURE (Row Indices - 0-based)
# ─────────────────────────────────────────────────────────────

# Row 2 in Excel = index 1: Module names (КМ, БМ, etc.)
ROW_SUBJECT_NAMES = 1

# Row 3 in Excel = index 2: Sub-subject names (ОН 1.1, ОН 1.2, etc.)
ROW_SUBJECT_NAMES_SUB = 2

# Row 4 in Excel = index 3: Hours/Credits (e.g., "72с-3к")
ROW_HOURS = 3

# Row 5 in Excel = index 4: Column labels (п, б, цэ, трад)
ROW_COLUMN_LABELS = 4

# Row 6 in Excel = index 5: First student data row
ROW_DATA_START = 5

# ─────────────────────────────────────────────────────────────
# EXCEL STRUCTURE (Column Indices - 0-based)
# ─────────────────────────────────────────────────────────────

# Column A = index 0: Row number
COL_NO = 0

# Column B = index 1: Student full name
COL_FULL_NAME = 1

# Column C = index 2: First subject's points column
# Each subject occupies 4 columns: п (points), б (letter), цэ (GPA), трад (traditional)
COL_START_SUBJECTS = 2

# Subjects stride: 4 columns per subject (п, б, цэ, трад)
SUBJECT_COLUMNS_STRIDE = 4

# ─────────────────────────────────────────────────────────────
# GRADE CONVERSION THRESHOLDS
# ─────────────────────────────────────────────────────────────
# Maps percentage score ranges to letter grades, GPA, and traditional grades

GRADE_THRESHOLDS = {
    95: {"letter": "A",  "gpa": 4.0,  "traditional_kz": "5 (өте жақсы)",           "traditional_ru": "5 (отлично)"},
    90: {"letter": "A-", "gpa": 3.67, "traditional_kz": "5 (өте жақсы)",           "traditional_ru": "5 (отлично)"},
    85: {"letter": "B+", "gpa": 3.33, "traditional_kz": "4 (жақсы)",               "traditional_ru": "4 (хорошо)"},
    80: {"letter": "B",  "gpa": 3.0,  "traditional_kz": "4 (жақсы)",               "traditional_ru": "4 (хорошо)"},
    75: {"letter": "B-", "gpa": 2.67, "traditional_kz": "4 (жақсы)",               "traditional_ru": "4 (хорошо)"},
    70: {"letter": "C+", "gpa": 2.33, "traditional_kz": "4 (жақсы)",               "traditional_ru": "4 (хорошо)"},
    65: {"letter": "C",  "gpa": 2.0,  "traditional_kz": "3 (қанағат)",             "traditional_ru": "3 (удовл)"},
    60: {"letter": "C-", "gpa": 1.67, "traditional_kz": "3 (қанағат)",             "traditional_ru": "3 (удовл)"},
    55: {"letter": "D+", "gpa": 1.33, "traditional_kz": "3 (қанағат)",             "traditional_ru": "3 (удовл)"},
    50: {"letter": "D",  "gpa": 1.0,  "traditional_kz": "3 (қанағат)",             "traditional_ru": "3 (удовл)"},
    0:  {"letter": "F",  "gpa": 0.0,  "traditional_kz": "2 (қанағаттанарлықсыз)", "traditional_ru": "2 (неуд)"},
}

# ─────────────────────────────────────────────────────────────
# ATTESTATION & ELECTIVES DEFAULTS
# ─────────────────────────────────────────────────────────────
# Used when source Excel is missing hours/credits data

ATTESTATION_HOURS = "108"
ATTESTATION_CREDITS = "4.5"

ELECTIVE_HOURS = "36"
ELECTIVE_CREDITS = "1.5"

# ─────────────────────────────────────────────────────────────
# INSTITUTION NAMES
# ─────────────────────────────────────────────────────────────
# Institution name and full legal names in Kazakh and Russian

INSTITUTION_NAME_KZ = "ҚАРТУ"
INSTITUTION_FULL_KZ = "ҚАЗАҚ ҰЛТТЫҚ ТЕХНИКАЛЫҚ УНИВЕРСИТЕТІ"

INSTITUTION_NAME_RU = "КНТУ"
INSTITUTION_FULL_RU = "КАЗАХСКИЙ НАЦИОНАЛЬНЫЙ ТЕХНИЧЕСКИЙ УНИВЕРСИТЕТ"

# ─────────────────────────────────────────────────────────────
# PROGRAM NAMES
# ─────────────────────────────────────────────────────────────
# Program display names in Kazakh and Russian

PROGRAM_IT_NAME_KZ = "Ақпараттық технологиялар (3F)"
PROGRAM_IT_NAME_RU = "Информационные технологии (3F)"

PROGRAM_ACCOUNTING_NAME_KZ = "Құрылымдық және қаржылық есептілік (3D)"
PROGRAM_ACCOUNTING_NAME_RU = "Бухгалтерский учёт и финансовая отчётность (3D)"

# ─────────────────────────────────────────────────────────────
# SYSTEM BEHAVIOR
# ─────────────────────────────────────────────────────────────

# Whether to log detailed parsing information
DEBUG_MODE = False

# Log file path
LOG_FILE = WORKSPACE_ROOT / "diploma_automation.log"

# Maximum number of students to process (0 = unlimited)
MAX_STUDENTS = 0

# Whether to validate Excel schema before processing
VALIDATE_SCHEMA = True

# Whether to generate both KZ and RU versions for each student
BILINGUAL_OUTPUT = True
