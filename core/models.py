# -*- coding: utf-8 -*-
"""
Data Models
===========
Core dataclasses for diploma automation system.

Models:
-------
- Grade: Single subject grade (points, letter, GPA, traditional)
- Subject: Subject definition (name, hours, credits)
- Student: Student record with all grades
- Diploma: Final diploma output specification
"""

from dataclasses import dataclass, field
from typing import Dict, Optional, List
from enum import Enum


class Language(Enum):
    """Supported languages."""
    KZ = "KZ"  # Kazakh
    RU = "RU"  # Russian


class Program(Enum):
    """Supported educational programs."""
    IT = "IT"
    ACCOUNTING = "ACCOUNTING"


@dataclass
class Grade:
    """
    Grade information for a single subject.
    
    Attributes:
        points: Percentage score (0-100 or empty string)
        letter: Letter grade (A, B+, C, etc.)
        gpa: GPA value (0.0 - 4.0)
        traditional_kz: Traditional Kazakh grade (5-2 / өте жақсы, жақсы, etc.)
        traditional_ru: Traditional Russian grade (5-2 / отлично, хорошо, etc.)
    """
    points: Optional[str] = None
    letter: str = ""
    gpa: Optional[float] = None
    traditional_kz: str = ""
    traditional_ru: str = ""

    def get_traditional(self, language: Language) -> str:
        """Get traditional grade for specified language."""
        if language == Language.KZ:
            return self.traditional_kz
        elif language == Language.RU:
            return self.traditional_ru
        else:
            raise ValueError(f"Unknown language: {language}")

    def is_empty(self) -> bool:
        """Check if grade is completely empty."""
        return all([
            not self.points,
            not self.letter,
            self.gpa is None or self.gpa == 0,
            not self.traditional_kz,
            not self.traditional_ru,
        ])


@dataclass
class Subject:
    """
    Subject definition with bilingual support.
    
    Attributes:
        name_kz: Subject name in Kazakh
        name_ru: Subject name in Russian
        hours: Contact hours (e.g., "72" or "108")
        credits: Credit value (e.g., "3" or "4.5")
        col_idx: Column index in source Excel (0-based)
        is_module_header: True if this is a module header (КМ, БМ)
        is_elective: True if this is an optional/elective course
    """
    name_kz: str
    name_ru: str
    hours: Optional[str] = None
    credits: Optional[str] = None
    col_idx: Optional[int] = None
    is_module_header: bool = False
    is_elective: bool = False

    def get_name(self, language: Language) -> str:
        """Get subject name for specified language."""
        if language == Language.KZ:
            return self.name_kz
        elif language == Language.RU:
            return self.name_ru
        else:
            raise ValueError(f"Unknown language: {language}")

    def is_incomplete(self) -> bool:
        """Check if subject is missing critical data (hours/credits)."""
        return not self.hours or not self.credits


@dataclass
class Student:
    """
    Student record with grades for all subjects.
    
    Attributes:
        full_name: Student full name
        diploma_number: Diploma ID number
        diploma_number_clean: Cleaned diploma number (without JB/KZ prefix)
        grades: Dictionary mapping subject name (KZ or RU) to Grade object
        sheet_name: Source Excel sheet name (3F-1, 3F-2, etc.)
        row_index: Row number in source Excel (0-based)
    """
    full_name: str
    diploma_number: str
    grades: Dict[str, Grade] = field(default_factory=dict)
    diploma_number_clean: str = ""
    sheet_name: str = ""
    row_index: int = 0

    def has_grade_for(self, subject_name: str) -> bool:
        """Check if student has a grade for given subject."""
        return subject_name in self.grades

    def get_grade(self, subject_name: str) -> Optional[Grade]:
        """Get grade for subject, or None if not found."""
        return self.grades.get(subject_name)

    def add_grade(self, subject_name: str, grade: Grade) -> None:
        """Add or update grade for subject."""
        self.grades[subject_name] = grade


@dataclass
class DiplomaPage:
    """
    Single page of a diploma supplement.
    
    Attributes:
        page_number: Page number (1-4)
        subjects: List of subjects to display on this page
        layout_data: Optional formatting/layout metadata
    """
    page_number: int
    subjects: List[str] = field(default_factory=list)
    layout_data: Dict = field(default_factory=dict)


@dataclass
class Diploma:
    """
    Complete diploma specification for generation.
    
    Attributes:
        student: Student object with grades
        program: Educational program (IT, ACCOUNTING)
        language: Output language (KZ, RU)
        academic_year: Academic year (e.g., "2025-2026")
        pages: List of diploma pages
        institution_name_kz: Institution name in Kazakh
        institution_name_ru: Institution name in Russian
        qualification_name_kz: Qualification name in Kazakh
        qualification_name_ru: Qualification name in Russian
    """
    student: Student
    program: Program
    language: Language
    academic_year: str = "2025-2026"
    pages: List[DiplomaPage] = field(default_factory=list)
    institution_name_kz: str = ""
    institution_name_ru: str = ""
    qualification_name_kz: str = ""
    qualification_name_ru: str = ""

    def get_institution_name(self) -> str:
        """Get institution name in current language."""
        if self.language == Language.KZ:
            return self.institution_name_kz
        elif self.language == Language.RU:
            return self.institution_name_ru
        else:
            return self.institution_name_kz or self.institution_name_ru

    def get_qualification_name(self) -> str:
        """Get qualification name in current language."""
        if self.language == Language.KZ:
            return self.qualification_name_kz
        elif self.language == Language.RU:
            return self.qualification_name_ru
        else:
            return self.qualification_name_kz or self.qualification_name_ru


@dataclass
class ProcessingResult:
    """
    Result of processing a batch of students.
    
    Attributes:
        total_students: Total students processed
        successful: Number of diplomas successfully generated
        failed: Number that failed
        errors: List of error messages
        warnings: List of warning messages
        statistics: Additional metrics (e.g., processing time)
    """
    total_students: int = 0
    successful: int = 0
    failed: int = 0
    errors: List[str] = field(default_factory=list)
    warnings: List[str] = field(default_factory=list)
    statistics: Dict = field(default_factory=dict)

    def success_rate(self) -> float:
        """Calculate success rate as percentage."""
        if self.total_students == 0:
            return 0.0
        return (self.successful / self.total_students) * 100

    def add_error(self, message: str) -> None:
        """Add error message."""
        self.errors.append(message)
        self.failed += 1

    def add_warning(self, message: str) -> None:
        """Add warning message."""
        self.warnings.append(message)
