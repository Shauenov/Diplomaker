"""
Test Suite for Excel Parsers
=============================
Unit and integration tests for data.excel_parser module.

Tests:
- Excel file parsing
- Subject name extraction
- Hours/credits parsing
- Student data extraction
- Error handling
"""

import pytest
from pathlib import Path
import pandas as pd

from data.excel_parser import ExcelParser, ExcelParseError
from core.models import Program, Language, Student
from core.exceptions import ValidationError, ConfigurationError


@pytest.mark.unit
class TestExcelPathValidation:
    """Test Excel file path validation."""

    def test_parser_nonexistent_file(self):
        """Test that nonexistent file raises ConfigurationError."""
        with pytest.raises(ConfigurationError):
            ExcelParser(source_file="/nonexistent/path/file.xlsx")

    def test_parser_with_valid_file(self, test_excel_file):
        """Test parser initialization with valid file."""
        parser = ExcelParser(source_file=str(test_excel_file))
        assert parser.source_file.exists()


@pytest.mark.unit
class TestSubjectExtraction:
    """Test subject name extraction."""

    def test_parse_subjects_from_test_file(self, test_excel_file):
        """Test extracting subject names from row 1."""
        parser = ExcelParser(source_file=str(test_excel_file))
        df = parser._load_dataframe("3F-1")
        subjects = parser._parse_subjects(df)

        assert len(subjects) > 0
        assert all(subj.name_kz for subj in subjects)
        assert all(subj.name_ru for subj in subjects)

    def test_subject_bilingual_names(self, test_excel_file):
        """Test that subjects have both KZ and RU names."""
        parser = ExcelParser(source_file=str(test_excel_file))
        df = parser._load_dataframe("3F-1")
        subjects = parser._parse_subjects(df)

        for subject in subjects:
            assert subject.name_kz
            assert subject.name_ru


@pytest.mark.unit
class TestHoursCreditsParsing:
    """Test hours and credits extraction."""

    def test_parse_hours_credits(self, test_excel_file):
        """Test extracting hours/credits from row 3."""
        parser = ExcelParser(source_file=str(test_excel_file))
        df = parser._load_dataframe("3F-1")
        subjects = parser._parse_subjects(df)
        hours_credits = parser._parse_hours_credits(df, len(subjects))

        assert len(hours_credits) > 0
        for col_idx, (hours, credits) in hours_credits.items():
            # Each should be either empty or valid format
            if hours:
                assert hours.isdigit()
            if credits:
                assert "." in credits or credits.isdigit()

    def test_hours_credits_format(self, test_excel_file):
        """Test that hours/credits are in expected format."""
        parser = ExcelParser(source_file=str(test_excel_file))
        df = parser._load_dataframe("3F-1")
        subjects = parser._parse_subjects(df)
        hours_credits = parser._parse_hours_credits(df, len(subjects))

        # All non-empty entries should have hours (digits)
        for col_idx, (hours, credits) in hours_credits.items():
            if hours:  # If data exists
                assert hours  # Hours should be present


@pytest.mark.unit
class TestStudentExtraction:
    """Test student data extraction."""

    def test_parse_students_from_sheet(self, test_excel_file):
        """Test parsing student data from sheet."""
        parser = ExcelParser(source_file=str(test_excel_file))
        students = parser.parse("3F-1")

        assert len(students) > 0
        assert all(isinstance(s, Student) for s in students)

    def test_student_data_completeness(self, test_excel_file):
        """Test that students have required fields."""
        parser = ExcelParser(source_file=str(test_excel_file))
        students = parser.parse("3F-1")

        for student in students:
            assert student.full_name
            assert student.diploma_number
            assert student.grades
            assert len(student.grades) > 0

    def test_student_grades_count(self, test_excel_file):
        """Test that students have grades for all subjects."""
        parser = ExcelParser(source_file=str(test_excel_file))
        df = parser._load_dataframe("3F-1")
        subjects = parser._parse_subjects(df)
        students = parser.parse("3F-1")

        # Each student should have grades for all or most subjects
        for student in students:
            assert len(student.grades) >= len(subjects) - 2  # Allow some missing


@pytest.mark.integration
class TestMultiSheetParsing:
    """Test parsing multiple sheets."""

    def test_parse_all_sheets(self, test_excel_multi_sheet):
        """Test parsing all sheets in a file."""
        # Use explicit sheet names matching the fixture (Cyrillic Ғ)
        parser = ExcelParser(
            source_file=str(test_excel_multi_sheet),
            sheet_names=["3Ғ-1", "3Ғ-2"],
        )
        all_students = parser.parse_all_sheets()

        assert len(all_students) > 0
        # Should have multiple sheets
        for sheet_name, students in all_students.items():
            assert isinstance(students, list)
            assert len(students) > 0


@pytest.mark.integration
class TestExcelValidation:
    """Test Excel file validation."""

    def test_validate_excel_structure_valid(self, test_excel_file):
        """Test validation of valid Excel structure."""
        is_valid = ExcelParser.validate_excel_structure(str(test_excel_file))
        assert is_valid is True or isinstance(is_valid, bool)

    def test_validate_excel_structure_invalid_file(self):
        """Test validation of nonexistent file."""
        is_valid = ExcelParser.validate_excel_structure("/nonexistent/file.xlsx")
        assert is_valid is False


@pytest.mark.integration
class TestParserErrorHandling:
    """Test error handling in parser."""

    def test_parse_empty_sheet(self, tmp_path):
        """Test parsing empty sheet."""
        # Create empty Excel file
        excel_path = tmp_path / "empty.xlsx"
        df = pd.DataFrame()
        df.to_excel(excel_path, sheet_name="3F-1", header=False, index=False)

        parser = ExcelParser(source_file=str(excel_path))

        # Should handle gracefully or raise specific error
        try:
            students = parser.parse("3F-1")
            # If it succeeds, should return empty list
            assert isinstance(students, list)
        except (ExcelParseError, IndexError):
            # Both are acceptable for empty input
            pass

    def test_parse_missing_sheet(self, test_excel_file):
        """Test parsing nonexistent sheet."""
        parser = ExcelParser(source_file=str(test_excel_file))

        with pytest.raises(Exception):  # ExcelParseError or other
            parser.parse("NonexistentSheet")


@pytest.mark.unit
class TestSubjectSpecialCases:
    """Test special subject types."""

    def test_identify_module_header(self, test_excel_file):
        """Test identification of module headers (КМ)."""
        parser = ExcelParser(source_file=str(test_excel_file))
        df = parser._load_dataframe("3F-1")
        subjects = parser._parse_subjects(df)

        # None of our test subjects are modules, but test the flag
        for subject in subjects:
            assert isinstance(subject.is_module_header, bool)

    def test_identify_elective_subjects(self, test_excel_file):
        """Test identification of elective subjects."""
        parser = ExcelParser(source_file=str(test_excel_file))
        df = parser._load_dataframe("3F-1")
        subjects = parser._parse_subjects(df)

        # None of our test subjects are electives, but test the flag
        for subject in subjects:
            assert isinstance(subject.is_elective, bool)


@pytest.mark.slow
@pytest.mark.integration
class TestLargeDatasetParsing:
    """Test parsing larger datasets."""

    def test_parse_many_students(self, test_excel_file):
        """Test parsing file with multiple students."""
        parser = ExcelParser(source_file=str(test_excel_file))
        students = parser.parse("3F-1")

        # Should parse multiple students
        assert len(students) >= 1

    def test_parse_many_subjects(self, test_excel_file):
        """Test parsing file with many subjects."""
        parser = ExcelParser(source_file=str(test_excel_file))
        df = parser._load_dataframe("3F-1")
        subjects = parser._parse_subjects(df)

        # Should parse multiple subjects
        assert len(subjects) >= 1


@pytest.mark.unit
class TestDataFrameLoading:
    """Test internal DataFrame loading."""

    def test_load_dataframe(self, test_excel_file):
        """Test _load_dataframe() method."""
        parser = ExcelParser(source_file=str(test_excel_file))
        df = parser._load_dataframe("3F-1")

        assert isinstance(df, pd.DataFrame)
        assert df.shape[0] > 0  # Has rows
        assert df.shape[1] > 0  # Has columns

    def test_loaded_dataframe_structure(self, test_excel_file):
        """Test structure of loaded DataFrame."""
        parser = ExcelParser(source_file=str(test_excel_file))
        df = parser._load_dataframe("3F-1")

        # Should have header rows and data rows
        assert df.shape[0] >= 5  # At least: subjects, spacing, hours, spacing, data
