"""
Pytest Configuration and Fixtures
===================================
Shared test fixtures and configuration for the diploma automation test suite.

Fixtures:
- sample_student: Basic Student object for testing
- sample_grade: Basic Grade object
- sample_subject: Basic Subject object
- sample_subjects: List of test subjects
- test_excel_file: Path to test Excel file
"""

import pytest
from pathlib import Path
import pandas as pd
from tempfile import TemporaryDirectory

from core.models import (
    Student,
    Grade,
    Subject,
    Diploma,
    Language,
    Program,
)


# ─────────────────────────────────────────────────────────────
# BASIC MODEL FIXTURES
# ─────────────────────────────────────────────────────────────


@pytest.fixture
def sample_grade() -> Grade:
    """Sample grade object (85 = B+)."""
    return Grade(
        points="85",
        letter="B+",
        gpa=3.33,
        traditional_kz="4 (жақсы)",
        traditional_ru="4 (хорошо)",
    )


@pytest.fixture
def empty_grade() -> Grade:
    """Empty grade object (no score)."""
    return Grade(
        points="",
        letter="",
        gpa=0.0,
        traditional_kz="",
        traditional_ru="",
    )


@pytest.fixture
def sample_subject() -> Subject:
    """Sample subject object."""
    return Subject(
        name_kz="Қазақ тілі",
        name_ru="Казахский язык",
        hours="72",
        credits="3",
        col_idx=2,
        is_module_header=False,
        is_elective=False,
    )


@pytest.fixture
def sample_module_header() -> Subject:
    """Sample module header subject (КМ)."""
    return Subject(
        name_kz="КМ 01 Компьютерлік негіздер",
        name_ru="КМ 01 Основы компьютеров",
        hours="72",
        credits="3",
        col_idx=2,
        is_module_header=True,
        is_elective=False,
    )


@pytest.fixture
def sample_subjects(sample_subject) -> list:
    """List of sample subjects for testing."""
    return [
        Subject(
            name_kz="Қазақ тілі",
            name_ru="Казахский язык",
            hours="72",
            credits="3",
            col_idx=2,
            is_module_header=False,
            is_elective=False,
        ),
        Subject(
            name_kz="Ағылшын тілі",
            name_ru="Английский язык",
            hours="72",
            credits="3",
            col_idx=6,
            is_module_header=False,
            is_elective=False,
        ),
        Subject(
            name_kz="Информатика",
            name_ru="Информатика",
            hours="72",
            credits="3",
            col_idx=10,
            is_module_header=False,
            is_elective=False,
        ),
    ]


@pytest.fixture
def sample_student(sample_grade) -> Student:
    """Sample student object with grades."""
    student = Student(
        full_name="Иванов Иван Ивановичович",
        diploma_number="2026001",
        grades={
            "Қазақ тілі": sample_grade,
            "Ағылшын тілі": Grade(
                points="92",
                letter="A-",
                gpa=3.67,
                traditional_kz="5 (өте жақсы)",
                traditional_ru="5 (отлично)",
            ),
            "Информатика": Grade(
                points="78",
                letter="C+",
                gpa=2.33,
                traditional_kz="4 (жақсы)",
                traditional_ru="4 (хорошо)",
            ),
        },
        sheet_name="3F-1",
        row_index=5,
    )
    return student


@pytest.fixture
def sample_diploma(sample_student) -> Diploma:
    """Sample diploma object."""
    return Diploma(
        student=sample_student,
        program=Program.IT,
        language=Language.KZ,
        academic_year="2025-2026",
        pages=[],
        institution_name_kz="ҚАРТУ",
        institution_name_ru="КНТУ",
        qualification_name_kz="Ақпараттық технологиялар",
        qualification_name_ru="Информационные технологии",
    )


# ─────────────────────────────────────────────────────────────
# EXCEL FILE FIXTURES
# ─────────────────────────────────────────────────────────────
#
# Layout matches config.settings:
#   ROW_SUBJECT_NAMES = 1 (0-indexed → row index 0)
#   ROW_HOURS         = 3 (0-indexed → row index 2)
#   ROW_DATA_START    = 5 (0-indexed → row index 5)
#   COL_NO            = 0
#   COL_FULL_NAME     = 1
#   COL_START_SUBJECTS= 2
#   SUBJECT_COLUMNS_STRIDE = 4 (п, б, цэ, трад)
#
# Columns: 0=№, 1=ФИО, 2=Subj1(п), 3=Subj1(б), 4=Subj1(цэ), 5=Subj1(трад),
#           6=Subj2(п), 7=Subj2(б), 8=Subj2(цэ), 9=Subj2(трад)


def _build_test_dataframe(subject_names, students_data, num_rows_total=8):
    """
    Build a properly structured test DataFrame for ExcelParser.

    Args:
        subject_names: list of "KZ name\\nRU name" strings
        students_data: list of dicts with keys: no, name, grades (list of ints)
        num_rows_total: total rows in the DataFrame
    """
    num_cols = 2 + len(subject_names) * 4  # №, ФИО, then 4 cols per subject
    rows = []

    # Row 0 (ROW_SUBJECT_NAMES - 1 = 0): Subject names at stride positions
    row0 = [None] * num_cols
    for i, sname in enumerate(subject_names):
        row0[2 + i * 4] = sname
    rows.append(row0)

    # Row 1: empty (sub-subjects)
    rows.append([None] * num_cols)

    # Row 2 (ROW_HOURS - 1 = 2): Hours/credits at stride positions
    row2 = [None] * num_cols
    for i in range(len(subject_names)):
        row2[2 + i * 4] = "72с-3к"
    rows.append(row2)

    # Row 3: empty (spacing)
    rows.append([None] * num_cols)

    # Row 4: empty (п/б/цэ/трад labels)
    rows.append([None] * num_cols)

    # Rows 5+ (ROW_DATA_START = 5): student data
    for sd in students_data:
        row = [None] * num_cols
        row[0] = sd["no"]
        row[1] = sd["name"]
        for j, grade_val in enumerate(sd["grades"]):
            row[2 + j * 4] = grade_val
        rows.append(row)

    # Pad to num_rows_total if needed
    while len(rows) < num_rows_total:
        rows.append([None] * num_cols)

    return pd.DataFrame(rows)


@pytest.fixture
def test_excel_file(tmp_path):
    """Create a test Excel file with sample data matching parser expectations."""
    excel_path = tmp_path / "test_grades.xlsx"

    df = _build_test_dataframe(
        subject_names=[
            "Қазақ тілі\nКазахский язык",
            "Ағылшын тілі\nАнглийский язык",
        ],
        students_data=[
            {"no": 1, "name": "Иванов Иван", "grades": [85, 92]},
            {"no": 2, "name": "Петров Петр", "grades": [90, 88]},
            {"no": 3, "name": "Сидоров Сидор", "grades": [78, 95]},
        ],
    )

    df.to_excel(excel_path, sheet_name="3F-1", header=False, index=False)
    return excel_path


@pytest.fixture
def test_excel_multi_sheet(tmp_path):
    """Create a multi-sheet test Excel file with Cyrillic Ғ sheet names."""
    excel_path = tmp_path / "test_multi_sheet.xlsx"

    df1 = _build_test_dataframe(
        subject_names=["Қазақ тілі\nКазахский язык"],
        students_data=[
            {"no": 1, "name": "Student 1", "grades": [85]},
            {"no": 2, "name": "Student 2", "grades": [90]},
        ],
    )

    df2 = _build_test_dataframe(
        subject_names=["Ағылшын тілі\nАнглийский язык"],
        students_data=[
            {"no": 1, "name": "Student 3", "grades": [88]},
            {"no": 2, "name": "Student 4", "grades": [92]},
        ],
    )

    with pd.ExcelWriter(excel_path) as writer:
        df1.to_excel(writer, sheet_name="3Ғ-1", header=False, index=False)
        df2.to_excel(writer, sheet_name="3Ғ-2", header=False, index=False)

    return excel_path


# ─────────────────────────────────────────────────────────────
# PYTEST CONFIGURATION
# ─────────────────────────────────────────────────────────────


def pytest_configure(config):
    """Configure pytest markers."""
    config.addinivalue_line(
        "markers", "unit: mark test as a unit test (fast, isolated)"
    )
    config.addinivalue_line(
        "markers", "integration: mark test as an integration test (slower, dependencies)"
    )
    config.addinivalue_line(
        "markers", "slow: mark test as slow (may take 5+ seconds)"
    )


# ─────────────────────────────────────────────────────────────
# ASSERTION HELPERS
# ─────────────────────────────────────────────────────────────


def assert_grade_valid(grade: Grade, expected_letter: str, expected_gpa: float):
    """Assert grade has expected values."""
    assert grade.letter == expected_letter, f"Expected letter {expected_letter}, got {grade.letter}"
    assert abs(grade.gpa - expected_gpa) < 0.01, f"Expected GPA {expected_gpa}, got {grade.gpa}"
    assert grade.traditional_kz, "Missing Kazakh traditional grade"
    assert grade.traditional_ru, "Missing Russian traditional grade"
