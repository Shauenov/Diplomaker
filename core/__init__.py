# -*- coding: utf-8 -*-
"""
Core Package
============
Core data models and utilities for diploma automation.

Modules:
--------
- models: Data classes (Grade, Subject, Student, Diploma)
- utils: Shared utility functions (normalization, parsing)
- exceptions: Custom exception types
"""

from .models import (
    Grade, Subject, Student, Diploma, DiplomaPage,
    Language, Program, ProcessingResult
)
from .utils import (
    normalize_key, clean_subject_name, parse_hours_credits,
    robust_clean, is_module_header, format_float_value
)
from .exceptions import (
    DiplomaAutomationError,
    ConfigurationError,
    ParseError,
    ValidationError,
    GenerationError,
)

__all__ = [
    # Models
    "Grade",
    "Subject",
    "Student",
    "Diploma",
    "DiplomaPage",
    "Language",
    "Program",
    "ProcessingResult",
    
    # Utils
    "normalize_key",
    "clean_subject_name",
    "parse_hours_credits",
    "robust_clean",
    "is_module_header",
    "format_float_value",
    
    # Exceptions
    "DiplomaAutomationError",
    "ConfigurationError",
    "ParseError",
    "ValidationError",
    "GenerationError",
]
