# -*- coding: utf-8 -*-
"""
Grade Conversion Engine
=======================
Convert percentage scores to letter grades, GPA, and traditional grades.

This module consolidates grade conversion logic that was previously scattered
across multiple scripts. It uses thresholds from config.settings.
"""

from typing import Optional, Union
from config.settings import GRADE_THRESHOLDS
from core.models import Grade, Language
from core.exceptions import ValidationError


def convert_score_to_grade(points: Optional[str]) -> Grade:
    """
    Convert test percentage score to a complete Grade object.
    
    This is the main entry point for grade conversion. It takes a raw score
    (percentage) and produces a Grade with:
    - Letter grade (A, B+, C, D, F)
    - GPA value (4.0 - 0.0)
    - Traditional Kazakh grade (5-2 with description)
    - Traditional Russian grade (5-2 with description)
    
    Args:
        points: Percentage score as string ("85", "92.5", etc.) or None
        
    Returns:
        Grade object with all converted values
        
    Raises:
        ValidationError: If score is invalid (not 0-100 range)
        
    Examples:
        >>> convert_score_to_grade("85")
        Grade(points="85", letter="B+", gpa=3.33, ...)
        
        >>> convert_score_to_grade("92")
        Grade(points="92", letter="A-", gpa=3.67, ...)
        
        >>> convert_score_to_grade(None)
        Grade(points=None, letter="", gpa=None, ...)
        
        >>> convert_score_to_grade("105")
        ValidationError: Score out of valid range
    """
    # Handle empty/missing scores
    if not points or str(points).strip() in ("", "nan", "NaN"):
        return Grade(
            points="",
            letter="",
            gpa=None,
            traditional_kz="",
            traditional_ru=""
        )
    
    # Parse and validate score
    try:
        score = float(str(points))
    except (ValueError, TypeError):
        raise ValidationError(f"Invalid score format: '{points}'")
    
    # Validate range
    if score < 0 or score > 100:
        raise ValidationError(f"Score out of valid range (0-100): {score}")
    
    # Convert to grade
    grade_data = _lookup_grade_threshold(score)
    
    return Grade(
        points=str(points),
        letter=grade_data.get("letter", ""),
        gpa=grade_data.get("gpa"),
        traditional_kz=grade_data.get("traditional_kz", ""),
        traditional_ru=grade_data.get("traditional_ru", "")
    )


def _lookup_grade_threshold(score: float) -> dict:
    """
    Find the appropriate grade threshold for a score.
    
    Uses GRADE_THRESHOLDS from config.settings which maps score ranges
    to grades. The thresholds are stored in descending order.
    
    Args:
        score: Numeric percentage score (0-100)
        
    Returns:
        Dictionary with keys: letter, gpa, traditional_kz, traditional_ru
        
    Examples:
        _lookup_grade_threshold(95) → {"letter": "A", "gpa": 4.0, ...}
        _lookup_grade_threshold(85) → {"letter": "B+", "gpa": 3.33, ...}
        _lookup_grade_threshold(30) → {"letter": "F", "gpa": 0.0, ...}
    """
    # Sort thresholds in descending order
    sorted_thresholds = sorted(GRADE_THRESHOLDS.keys(), reverse=True)
    
    for threshold in sorted_thresholds:
        if score >= threshold:
            return GRADE_THRESHOLDS[threshold]
    
    # Shouldn't reach here if thresholds configured correctly
    return GRADE_THRESHOLDS.get(0, {"letter": "", "gpa": 0.0, "traditional_kz": "", "traditional_ru": ""})


def get_gpa_value(points: Optional[str]) -> Optional[float]:
    """
    Get just the GPA value for a score (shortcut).
    
    Args:
        points: Percentage score
        
    Returns:
        GPA value (0.0 - 4.0) or None if no score
    """
    grade = convert_score_to_grade(points)
    return grade.gpa


def get_letter_grade(points: Optional[str]) -> str:
    """
    Get just the letter grade for a score (shortcut).
    
    Args:
        points: Percentage score
        
    Returns:
        Letter grade (A, B+, C, D, F) or empty string if no score
    """
    grade = convert_score_to_grade(points)
    return grade.letter


def get_traditional_grade(points: Optional[str], language: Union[str, Language] = "KZ") -> str:
    """
    Get traditional grade in specified language (shortcut).
    
    Args:
        points: Percentage score
        language: "KZ" for Kazakh, "RU" for Russian, or Language enum
        
    Returns:
        Traditional grade description or empty string if no score
    """
    # Handle Language enum
    if isinstance(language, Language):
        language = language.value
    
    grade = convert_score_to_grade(points)
    if language.upper() == "KZ":
        return grade.traditional_kz
    elif language.upper() == "RU":
        return grade.traditional_ru
    else:
        raise ValueError(f"Unknown language: {language}")


# ─────────────────────────────────────────────────────────────
# Testing utilities (for verification)
# ─────────────────────────────────────────────────────────────

def verify_grade_conversion():
    """
    Verify grade conversion is working correctly.
    
    Used for testing/validation. Checks key conversion points.
    """
    test_cases = [
        (95, "A"),
        (90, "A-"),
        (85, "B+"),
        (80, "B"),
        (75, "B-"),
        (70, "C+"),
        (65, "C"),
        (50, "D"),
        (0, "F"),
    ]
    
    results = []
    for score, expected_letter in test_cases:
        grade = convert_score_to_grade(str(score))
        actual_letter = grade.letter
        status = "✓" if actual_letter == expected_letter else "✗"
        results.append(f"{status} {score}: expected {expected_letter}, got {actual_letter}")
    
    return results


if __name__ == "__main__":
    # Quick test
    results = verify_grade_conversion()
    for result in results:
        print(result)
