# -*- coding: utf-8 -*-
"""
Excel Parser
============
Parse source Excel files containing student grades.

Classes:
--------
- ExcelParser: Main parser for reading Excel grade files
- ExcelParseError: Custom exception for parsing errors

Usage:
------
    from data.excel_parser import ExcelParser
    from core.models import Program
    
    parser = ExcelParser(source_file="grades.xlsx")
    students = parser.parse("3F-1")
    
    for student in students:
        print(f"{student.full_name}: {len(student.grades)} subjects")
"""

import pandas as pd
from pathlib import Path
from typing import List, Dict, Optional, Tuple
import logging

from core.models import Student, Subject, Grade, Program, Language
from core.utils import (
    clean_subject_name, parse_hours_credits, 
    robust_clean, is_module_header, normalize_key
)
from core.exceptions import ConfigurationError, ParseError


logger = logging.getLogger(__name__)


class ExcelParseError(ParseError):
    """Custom exception for Excel parsing errors."""
    pass


class ExcelParser:
    """
    Parser for extracting student data from Excel grade sheets.
    
    Expected Excel Structure:
    -------------------------
    Row 0: (Optional header)
    Row 1: Subject names (bilingual: "Kazakh\\nRussian")
    Row 2: (Empty spacing row)
    Row 3: Hours and credits ("72с-3к" format)
    Row 4+: Student data rows
    
    Columns:
    --------
    Col 0: Row number (№)
    Col 1: Student full name
    Col 2+: Subject columns (each subject has 4 columns: Points, Letter, GPA, Traditional)
    
    Attributes:
        source_file: Path to source Excel file
        sheet_names: List of sheet names to process (None = all sheets)
        row_subject_names: Row index for subject names (default: 1)
        row_hours_credits: Row index for hours/credits (default: 3)
        row_data_start: First data row index (default: 4)
        col_student_name: Column index for student names (default: 1)
        col_start_subjects: First subject column index (default: 2)
        cols_per_subject: Number of columns per subject (default: 4)
    """
    
    def __init__(
        self,
        source_file: str,
        sheet_names: Optional[List[str]] = None,
        row_subject_names: int = 1,
        row_hours_credits: int = 3,
        row_data_start: int = 4,
        col_student_name: int = 1,
        col_start_subjects: int = 2,
        cols_per_subject: int = 4,
    ):
        """
        Initialize Excel parser.
        
        Args:
            source_file: Path to source Excel file
            sheet_names: Sheet names to parse (None = detect automatically)
            row_subject_names: Row index for subject names (0-based)
            row_hours_credits: Row index for hours/credits (0-based)
            row_data_start: First data row index (0-based)
            col_student_name: Column index for student names (0-based)
            col_start_subjects: First subject column index (0-based)
            cols_per_subject: Columns per subject (Points, Letter, GPA, Traditional)
            
        Raises:
            ConfigurationError: If source file doesn't exist
        """
        self.source_file = Path(source_file)
        self.sheet_names = sheet_names
        
        # Row indices
        self.row_subject_names = row_subject_names
        self.row_hours_credits = row_hours_credits
        self.row_data_start = row_data_start
        
        # Column indices
        self.col_student_name = col_student_name
        self.col_start_subjects = col_start_subjects
        self.cols_per_subject = cols_per_subject
        
        # Validate file exists
        if not self.source_file.exists():
            raise ConfigurationError(
                f"Source Excel file not found: {self.source_file}"
            )
    
    def parse(self, sheet_name: str) -> List[Student]:
        """
        Parse single sheet and return list of students with grades.
        
        Args:
            sheet_name: Name of Excel sheet to parse
            
        Returns:
            List of Student objects with populated grades
            
        Raises:
            ExcelParseError: If parsing fails
        """
        logger.info(f"Parsing sheet: {sheet_name}")
        
        try:
            # Load DataFrame
            df = self._load_dataframe(sheet_name)
            
            # Parse subjects and hours/credits
            subjects = self._parse_subjects(df)
            hours_credits = self._parse_hours_credits(df, len(subjects))
            
            # Merge hours/credits into subjects
            for idx, subject in enumerate(subjects):
                col_idx = self.col_start_subjects + (idx * self.cols_per_subject)
                if col_idx in hours_credits:
                    hours, credits = hours_credits[col_idx]
                    subject.hours = hours
                    subject.credits = credits
                    subject.col_idx = col_idx
            
            # Parse students
            students = self._parse_students(df, subjects, sheet_name)
            
            logger.info(f"  Parsed {len(students)} students, {len(subjects)} subjects")
            
            return students
            
        except Exception as e:
            raise ExcelParseError(
                f"Failed to parse sheet '{sheet_name}': {str(e)}"
            ) from e
    
    def get_subjects(self, sheet_name: str) -> List[Subject]:
        """
        Load subjects with hours/credits fully populated.
        
        Parameters:
            sheet_name (str): Sheet name to load
        
        Returns:
            List[Subject]: Subject objects with hours and credits filled in
        """
        df = self._load_dataframe(sheet_name)
        subjects = self._parse_subjects(df)
        hours_credits = self._parse_hours_credits(df, len(subjects))
        self._apply_hours_credits(subjects, hours_credits)
        return subjects
    
    @staticmethod
    def _apply_hours_credits(
        subjects: List[Subject],
        hours_credits: Dict[int, Tuple[str, str]],
    ) -> None:
        """
        Fill Subject.hours and Subject.credits from parsed dict.
        
        Mutates subjects in-place.
        """
        for subj in subjects:
            if subj.col_idx in hours_credits:
                subj.hours, subj.credits = hours_credits[subj.col_idx]
    
    def parse_all_sheets(self) -> Dict[str, List[Student]]:
        """
        Parse all sheets in the Excel file.
        
        Returns:
            Dictionary mapping sheet names to lists of students
            
        Raises:
            ExcelParseError: If parsing fails
        """
        if self.sheet_names:
            # Use specified sheet names
            sheets = self.sheet_names
        else:
            # Auto-detect sheet names
            try:
                xl_file = pd.ExcelFile(self.source_file)
                sheets = xl_file.sheet_names
                logger.info(f"Auto-detected {len(sheets)} sheets: {sheets}")
            except Exception as e:
                raise ExcelParseError(
                    f"Failed to read Excel file: {str(e)}"
                ) from e
        
        all_students = {}
        
        for sheet_name in sheets:
            try:
                students = self.parse(sheet_name)
                all_students[sheet_name] = students
            except Exception as e:
                logger.warning(f"  Failed to parse sheet '{sheet_name}': {e}")
                all_students[sheet_name] = []
        
        return all_students
    
    def _load_dataframe(self, sheet_name: str) -> pd.DataFrame:
        """
        Load Excel sheet as DataFrame.
        
        Args:
            sheet_name: Sheet name to load
            
        Returns:
            DataFrame with no header (all rows as data)
        """
        try:
            df = pd.read_excel(
                self.source_file,
                sheet_name=sheet_name,
                header=None  # Don't treat any row as header
            )
            return df
        except Exception as e:
            raise ExcelParseError(
                f"Failed to load sheet '{sheet_name}': {str(e)}"
            ) from e
    
    def _parse_subjects(self, df: pd.DataFrame) -> List[Subject]:
        """
        Extract subject names from row 1.
        
        Args:
            df: Source DataFrame
            
        Returns:
            List of Subject objects with bilingual names
        """
        subjects = []
        
        col_idx = self.col_start_subjects
        
        while col_idx < df.shape[1]:
            # Read subject name cell
            raw_subject = df.iloc[self.row_subject_names, col_idx]
            
            # Skip empty columns
            if pd.isna(raw_subject) or not str(raw_subject).strip():
                col_idx += self.cols_per_subject
                continue
            
            # Parse bilingual name
            name_kz, name_ru = clean_subject_name(str(raw_subject))
            
            # Create subject object
            subject = Subject(
                name_kz=name_kz,
                name_ru=name_ru,
                col_idx=col_idx,
                is_module_header=is_module_header(name_kz),
            )
            
            subjects.append(subject)
            
            # Move to next subject (skip 4 columns: Points, Letter, GPA, Traditional)
            col_idx += self.cols_per_subject
        
        return subjects
    
    def _parse_hours_credits(
        self, 
        df: pd.DataFrame, 
        num_subjects: int
    ) -> Dict[int, Tuple[str, str]]:
        """
        Extract hours and credits from row 3.
        
        Args:
            df: Source DataFrame
            num_subjects: Number of subjects to read
            
        Returns:
            Dictionary mapping column index to (hours, credits) tuple
        """
        hours_credits = {}
        
        col_idx = self.col_start_subjects
        
        for _ in range(num_subjects):
            # Read hours/credits cell
            raw_hc = df.iloc[self.row_hours_credits, col_idx]
            
            # Parse format "72с-3к"
            hours, credits = parse_hours_credits(str(raw_hc))
            
            hours_credits[col_idx] = (hours, credits)
            
            # Move to next subject
            col_idx += self.cols_per_subject
        
        return hours_credits
    
    def _parse_students(
        self, 
        df: pd.DataFrame, 
        subjects: List[Subject],
        sheet_name: str
    ) -> List[Student]:
        """
        Extract student data rows.
        
        Args:
            df: Source DataFrame
            subjects: List of Subject objects
            sheet_name: Current sheet name
            
        Returns:
            List of Student objects
        """
        students = []
        
        for row_idx in range(self.row_data_start, df.shape[0]):
            # Get student name
            raw_name = df.iloc[row_idx, self.col_student_name]
            
            # Skip empty rows
            if pd.isna(raw_name) or not str(raw_name).strip():
                continue
            
            student_name = robust_clean(raw_name)
            if not student_name:
                continue
            
            # Create student object
            student = Student(
                full_name=student_name,
                diploma_number="",  # Will be filled later
                sheet_name=sheet_name,
                row_index=row_idx,
            )
            
            # Parse grades for each subject
            for subject in subjects:
                col_idx = subject.col_idx
                
                # Read grade columns: Points, Letter, GPA, Traditional
                raw_points = df.iloc[row_idx, col_idx] if col_idx < df.shape[1] else None
                raw_letter = df.iloc[row_idx, col_idx + 1] if col_idx + 1 < df.shape[1] else None
                raw_gpa = df.iloc[row_idx, col_idx + 2] if col_idx + 2 < df.shape[1] else None
                raw_traditional = df.iloc[row_idx, col_idx + 3] if col_idx + 3 < df.shape[1] else None
                
                # Clean values
                points = robust_clean(raw_points)
                letter = robust_clean(raw_letter)
                gpa_str = robust_clean(raw_gpa)
                traditional = robust_clean(raw_traditional)
                
                # Parse GPA as float
                gpa = None
                if gpa_str:
                    try:
                        gpa = float(gpa_str.replace(',', '.'))
                    except ValueError:
                        gpa = None
                
                # Create Grade object
                grade = Grade(
                    points=points if points else None,
                    letter=letter,
                    gpa=gpa,
                    traditional_kz=traditional,
                    traditional_ru=traditional,
                )
                
                # Add grade to student (use KZ name as key)
                student.add_grade(subject.name_kz, grade)
            
            students.append(student)
        
        return students
    
    @staticmethod
    def validate_excel_structure(file_path: str) -> bool:
        """
        Validate that Excel file has expected structure.
        
        Args:
            file_path: Path to Excel file
            
        Returns:
            True if structure is valid, False otherwise
        """
        try:
            # Check file exists
            path = Path(file_path)
            if not path.exists():
                return False
            
            # Try to load file
            df = pd.read_excel(file_path, header=None)
            
            # Basic checks
            if df.shape[0] < 5:  # Need at least 5 rows
                logger.warning("Excel has too few rows")
                return False
            
            if df.shape[1] < 6:  # Need at least 6 columns
                logger.warning("Excel has too few columns")
                return False
            
            return True
            
        except Exception as e:
            logger.error(f"Validation failed: {e}")
            return False
