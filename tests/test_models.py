"""
Test Suite for Data Models
===========================
Unit tests for core.models dataclasses and enums.

Tests:
- Model initialization
- Model methods
- Property getters
- Enum values
"""

import pytest
from core.models import (
    Grade,
    Subject,
    Student,
    Diploma,
    Language,
    Program,
    ProcessingResult,
)


@pytest.mark.unit
class TestGradeModel:
    """Test Grade dataclass."""

    def test_grade_creation(self, sample_grade):
        """Test Grade object creation."""
        assert sample_grade.points == "85"
        assert sample_grade.letter == "B+"
        assert sample_grade.gpa == 3.33

    def test_grade_empty_check(self, empty_grade):
        """Test is_empty() method."""
        assert empty_grade.is_empty()

    def test_grade_not_empty(self, sample_grade):
        """Test is_empty() returns False."""
        assert not sample_grade.is_empty()

    def test_grade_get_traditional_language(self, sample_grade):
        """Test get_traditional() for different languages."""
        kz_grade = sample_grade.get_traditional(Language.KZ)
        ru_grade = sample_grade.get_traditional(Language.RU)

        assert kz_grade != ru_grade
        assert "кз" not in kz_grade.lower() or "4" in kz_grade


@pytest.mark.unit
class TestSubjectModel:
    """Test Subject dataclass."""

    def test_subject_creation(self, sample_subject):
        """Test Subject object creation."""
        assert sample_subject.name_kz == "Қазақ тілі"
        assert sample_subject.name_ru == "Казахский язык"
        assert sample_subject.hours == "72"
        assert sample_subject.credits == "3"

    def test_subject_module_header(self, sample_module_header):
        """Test module header identification."""
        assert sample_module_header.is_module_header
        assert "КМ" in sample_module_header.name_kz

    def test_subject_not_module_header(self, sample_subject):
        """Test non-module subject."""
        assert not sample_subject.is_module_header

    def test_subject_get_name_kz(self, sample_subject):
        """Test get_name() for Kazakh."""
        name = sample_subject.get_name(Language.KZ)
        assert name == sample_subject.name_kz

    def test_subject_get_name_ru(self, sample_subject):
        """Test get_name() for Russian."""
        name = sample_subject.get_name(Language.RU)
        assert name == sample_subject.name_ru

    def test_subject_incomplete_missing_hours(self):
        """Test is_incomplete() for missing hours."""
        subject = Subject(
            name_kz="Test",
            name_ru="Test",
            hours="",  # Missing
            credits="3",
            col_idx=2,
            is_module_header=False,
            is_elective=False,
        )
        assert subject.is_incomplete()

    def test_subject_incomplete_missing_credits(self):
        """Test is_incomplete() for missing credits."""
        subject = Subject(
            name_kz="Test",
            name_ru="Test",
            hours="72",
            credits="",  # Missing
            col_idx=2,
            is_module_header=False,
            is_elective=False,
        )
        assert subject.is_incomplete()

    def test_subject_complete(self, sample_subject):
        """Test complete subject."""
        assert not sample_subject.is_incomplete()


@pytest.mark.unit
class TestStudentModel:
    """Test Student dataclass."""

    def test_student_creation(self, sample_student):
        """Test Student object creation."""
        assert sample_student.full_name == "Иванов Иван Ивановичович"
        assert sample_student.diploma_number == "2026001"
        assert len(sample_student.grades) == 3

    def test_student_has_grade(self, sample_student):
        """Test has_grade_for() method."""
        assert sample_student.has_grade_for("Қазақ тілі")
        assert not sample_student.has_grade_for("Отсутствующий предмет")

    def test_student_get_grade(self, sample_student):
        """Test get_grade() method."""
        grade = sample_student.get_grade("Қазақ тілі")
        assert grade is not None
        assert grade.letter == "B+"

    def test_student_get_grade_missing(self, sample_student):
        """Test get_grade() returns None for missing subject."""
        grade = sample_student.get_grade("Несуществующий предмет")
        assert grade is None

    def test_student_add_grade(self, sample_student, sample_grade):
        """Test add_grade() method."""
        sample_student.add_grade("Математика", sample_grade)
        assert sample_student.has_grade_for("Математика")
        assert sample_student.get_grade("Математика") == sample_grade

    def test_student_grades_count(self, sample_student):
        """Test total number of grades."""
        assert len(sample_student.grades) == 3


@pytest.mark.unit
class TestDiplomaModel:
    """Test Diploma dataclass."""

    def test_diploma_creation(self, sample_diploma):
        """Test Diploma object creation."""
        assert sample_diploma.student is not None
        assert sample_diploma.program == Program.IT
        assert sample_diploma.language == Language.KZ
        assert sample_diploma.academic_year == "2025-2026"

    def test_diploma_get_institution_name_kz(self, sample_diploma):
        """Test get_institution_name() for Kazakh."""
        # Diploma already has language set to KZ
        name = sample_diploma.get_institution_name()
        assert name == "ҚАРТУ"

    def test_diploma_get_institution_name_ru(self):
        """Test get_institution_name() for Russian."""
        # Create diploma with Russian language
        diploma_ru = Diploma(
            student=Student("Test", "001", {}, "", "", 0),
            program=Program.IT,
            language=Language.RU,
            academic_year="2025-2026",
            institution_name_kz="ҚАРТУ",
            institution_name_ru="КНТУ",
        )
        name = diploma_ru.get_institution_name()
        assert name == "КНТУ"

    def test_diploma_get_qualification_name_kz(self, sample_diploma):
        """Test get_qualification_name() for Kazakh."""
        # Diploma already has language set to KZ
        name = sample_diploma.get_qualification_name()
        assert "Ақпараттық" in name

    def test_diploma_get_qualification_name_ru(self):
        """Test get_qualification_name() for Russian."""
        # Create diploma with Russian language
        diploma_ru = Diploma(
            student=Student("Test", "001", {}, "", "", 0),
            program=Program.IT,
            language=Language.RU,
            academic_year="2025-2026",
            qualification_name_kz="Ақпараттық технологиялар",
            qualification_name_ru="Информационные технологии",
        )
        name = diploma_ru.get_qualification_name()
        assert "Информационные" in name


@pytest.mark.unit
class TestLanguageEnum:
    """Test Language enum."""

    def test_language_kz(self):
        """Test Kazakh language enum."""
        assert Language.KZ.value == "KZ"

    def test_language_ru(self):
        """Test Russian language enum."""
        assert Language.RU.value == "RU"

    def test_language_enum_values(self):
        """Test all language values are strings."""
        for lang in Language:
            assert isinstance(lang.value, str)


@pytest.mark.unit
class TestProgramEnum:
    """Test Program enum."""

    def test_program_it(self):
        """Test IT program enum."""
        assert Program.IT.value == "IT"

    def test_program_accounting(self):
        """Test Accounting program enum."""
        assert Program.ACCOUNTING.value == "ACCOUNTING"

    def test_program_enum_values(self):
        """Test all program values are strings."""
        for prog in Program:
            assert isinstance(prog.value, str)


@pytest.mark.unit
class TestProcessingResult:
    """Test ProcessingResult dataclass."""

    def test_processing_result_creation(self):
        """Test ProcessingResult object creation."""
        result = ProcessingResult(
            total_students=100,
            successful=95,
            failed=5,
            errors=["Error 1", "Error 2"],
            warnings=["Warning 1"],
            statistics={"rate": 0.95},
        )

        assert result.total_students == 100
        assert result.successful == 95
        assert result.failed == 5
        assert len(result.errors) == 2
        assert len(result.warnings) == 1

    def test_processing_result_success_rate(self):
        """Test success rate calculation."""
        result = ProcessingResult(
            total_students=100,
            successful=80,
            failed=20,
            errors=[],
            warnings=[],
            statistics={},
        )

        success_rate = result.successful / result.total_students
        assert success_rate == 0.8


@pytest.mark.integration
class TestModelIntegration:
    """Test interactions between models."""

    def test_student_with_grades(self, sample_student, sample_grade):
        """Test Student with Grade objects."""
        for subject, grade in sample_student.grades.items():
            assert isinstance(grade, Grade)
            assert grade.letter  # Should have a grade letter

    def test_diploma_with_student(self, sample_diploma):
        """Test Diploma references Student correctly."""
        student = sample_diploma.student
        assert student.full_name
        assert len(student.grades) > 0

    def test_subject_language_consistency(self, sample_subject):
        """Test Subject has both languages."""
        assert sample_subject.name_kz
        assert sample_subject.name_ru
        assert sample_subject.name_kz != sample_subject.name_ru
