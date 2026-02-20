"""
Test Suite for Grade Converters
================================
Unit tests for core.converters module.

Tests:
- convert_score_to_grade() with various scores
- Grade boundary conditions (thresholds)
- Error handling for invalid inputs
- Shortcut functions (get_gpa_value, etc.)
"""

import pytest
from core.converters import (
    convert_score_to_grade,
    get_gpa_value,
    get_letter_grade,
    get_traditional_grade,
)
from core.models import Grade, Language
from core.exceptions import ValidationError


@pytest.mark.unit
class TestGradeConversion:
    """Test grade conversion functionality."""

    def test_convert_score_perfect(self):
        """Test conversion of perfect score (100)."""
        grade = convert_score_to_grade("100")
        assert grade.letter == "A"
        assert grade.gpa == 4.0
        assert grade.points == "100"

    def test_convert_score_high(self):
        """Test conversion of high score (95+)."""
        grade = convert_score_to_grade("95")
        assert grade.letter == "A"
        assert grade.gpa == 4.0

    def test_convert_score_a_minus(self):
        """Test conversion of A- range (90-94)."""
        grade = convert_score_to_grade("92")
        assert grade.letter == "A-"
        assert grade.gpa == 3.67

    def test_convert_score_b_plus(self):
        """Test conversion of B+ range (85-89)."""
        grade = convert_score_to_grade("85")
        assert grade.letter == "B+"
        assert grade.gpa == 3.33

    def test_convert_score_b(self):
        """Test conversion of B range (80-84)."""
        grade = convert_score_to_grade("80")
        assert grade.letter == "B"
        assert grade.gpa == 3.0

    def test_convert_score_c_plus(self):
        """Test conversion of C+ range (70-74)."""
        grade = convert_score_to_grade("70")
        assert grade.letter == "C+"
        assert grade.gpa == 2.33

    def test_convert_score_d(self):
        """Test conversion of D range (50-64)."""
        grade = convert_score_to_grade("50")
        assert grade.letter == "D"
        assert grade.gpa == 1.0

    def test_convert_score_f(self):
        """Test conversion of F (0)."""
        grade = convert_score_to_grade("0")
        assert grade.letter == "F"
        assert grade.gpa == 0.0

    def test_convert_score_empty(self):
        """Test conversion of empty score (no grade)."""
        grade = convert_score_to_grade("")
        assert grade.letter == ""
        assert grade.points == ""

    def test_convert_score_none(self):
        """Test conversion of None (no grade)."""
        grade = convert_score_to_grade(None)
        assert grade.letter == ""

    def test_convert_score_out_of_range_high(self):
        """Test that score > 100 raises ValidationError."""
        with pytest.raises(ValidationError):
            convert_score_to_grade("105")

    def test_convert_score_out_of_range_negative(self):
        """Test that negative score raises ValidationError."""
        with pytest.raises(ValidationError):
            convert_score_to_grade("-5")

    def test_convert_score_non_numeric(self):
        """Test that non-numeric score raises ValidationError."""
        with pytest.raises(ValidationError):
            convert_score_to_grade("abc")

    def test_traditional_kazakh_grade(self, sample_grade):
        """Test Kazakh traditional grade conversion."""
        # sample_grade is 85 = B+ = 4 (жақсы)
        assert "4" in sample_grade.traditional_kz
        assert "жақсы" in sample_grade.traditional_kz

    def test_traditional_russian_grade(self, sample_grade):
        """Test Russian traditional grade conversion."""
        # sample_grade is 85 = B+ = 4 (хорошо)
        assert "4" in sample_grade.traditional_ru
        assert "хорошо" in sample_grade.traditional_ru

    @pytest.mark.parametrize(
        "score,expected_letter",
        [
            ("100", "A"),
            ("95", "A"),
            ("92", "A-"),
            ("90", "A-"),
            ("87", "B+"),
            ("85", "B+"),
            ("82", "B"),
            ("80", "B"),
            ("77", "B-"),
            ("75", "B-"),
            ("72", "C+"),
            ("70", "C+"),
            ("67", "C"),
            ("65", "C"),
            ("60", "C-"),
            ("55", "D+"),
            ("52", "D"),
            ("50", "D"),
        ],
    )
    def test_grade_boundaries_parametrized(self, score, expected_letter):
        """Test all grade boundaries with parametrize."""
        grade = convert_score_to_grade(score)
        assert grade.letter == expected_letter


@pytest.mark.unit
class TestGradeShortcuts:
    """Test shortcut functions for grade retrieval."""

    def test_get_gpa_value_from_string(self):
        """Test extracting GPA from score string."""
        gpa = get_gpa_value("85")
        assert gpa == 3.33

    def test_get_gpa_value_high_score(self):
        """Test GPA for high score."""
        gpa = get_gpa_value("95")
        assert gpa == 4.0

    def test_get_letter_grade_from_string(self):
        """Test extracting letter grade from score string."""
        letter = get_letter_grade("85")
        assert letter == "B+"

    def test_get_traditional_grade_kz(self):
        """Test extracting Kazakh traditional grade."""
        traditional = get_traditional_grade("85", Language.KZ)
        assert "4" in traditional
        assert "жақсы" in traditional

    def test_get_traditional_grade_ru(self):
        """Test extracting Russian traditional grade."""
        traditional = get_traditional_grade("85", Language.RU)
        assert "4" in traditional
        assert "хорошо" in traditional


@pytest.mark.unit
class TestGradeObject:
    """Test Grade model behavior."""

    def test_grade_is_empty(self, empty_grade):
        """Test is_empty() method."""
        assert empty_grade.is_empty()

    def test_grade_is_not_empty(self, sample_grade):
        """Test is_empty() returns False for valid grade."""
        assert not sample_grade.is_empty()

    def test_grade_get_traditional_kz(self, sample_grade):
        """Test get_traditional() method for Kazakh."""
        traditional = sample_grade.get_traditional(Language.KZ)
        assert "4" in traditional

    def test_grade_get_traditional_ru(self, sample_grade):
        """Test get_traditional() method for Russian."""
        traditional = sample_grade.get_traditional(Language.RU)
        assert "4" in traditional

    def test_grade_initialization(self):
        """Test Grade object initialization."""
        grade = Grade(
            points="85",
            letter="B+",
            gpa=3.33,
            traditional_kz="4",
            traditional_ru="4",
        )
        assert grade.points == "85"
        assert grade.letter == "B+"
        assert grade.gpa == 3.33


@pytest.mark.integration
class TestGradeConversionEdgeCases:
    """Test edge cases and special scenarios."""

    def test_whitespace_handling(self):
        """Test that whitespace is handled correctly."""
        grade = convert_score_to_grade("  85  ")
        assert grade.letter == "B+"

    def test_float_score_truncated(self):
        """Test that float scores are handled."""
        # Assuming floats are converted to int
        grade = convert_score_to_grade("85.7")
        assert grade.letter == "B+"

    def test_score_exactly_on_threshold(self):
        """Test score exactly on threshold boundary."""
        grade_85 = convert_score_to_grade("85")
        grade_84 = convert_score_to_grade("84")

        assert grade_85.letter == "B+"
        assert grade_84.letter == "B"

    def test_consistency_multiple_calls(self):
        """Test that same score always produces same grade."""
        grade1 = convert_score_to_grade("85")
        grade2 = convert_score_to_grade("85")

        assert grade1.letter == grade2.letter
        assert grade1.gpa == grade2.gpa
