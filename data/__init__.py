"""
Data Layer - Excel I/O and File Operations

Provides interfaces for:
- Reading source Excel files with student grades
- Writing output diploma Excel files
- Data validation and schema checking
"""

from .excel_parser import ExcelParser, ExcelParseError
from .excel_generator import DiplomaGenerator, DiplomaGenerationError

__all__ = [
    "ExcelParser",
    "ExcelParseError",
    "DiplomaGenerator",
    "DiplomaGenerationError",
]
