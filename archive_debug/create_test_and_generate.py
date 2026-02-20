#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Create Test Data and Generate Sample Diplomas
==============================================
Development utility for creating test Excel files and generating sample diplomas.

This script:
1. Creates a sample Excel file with test student data
2. Parses the test data using ExcelParser
3. Generates diplomas using DiplomaGenerator
4. Outputs to 'test_output' directory

Usage:
    python create_test_and_generate.py
"""

import sys
from pathlib import Path
import pandas as pd
from datetime import datetime
import logging

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent))

from core.models import Student, Subject, Grade, Program, Language
from core.utils import parse_hours_credits, clean_subject_name
from data.excel_parser import ExcelParser
from data.excel_generator import DiplomaGenerator


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


def create_test_excel(output_path: str = "test_grades.xlsx"):
    """
    Create a sample Excel file with test student data.
    
    Args:
        output_path: Path for output Excel file
        
    Returns:
        Path to created file
    """
    logger.info("Creating test Excel file...")
    
    # Sample bilingual subjects
    subjects = [
        "Қазақ тілі\nКазахский язык",
        "Ағылшын тілі\nАнглийский язык",
        "Математика\nМатематика",
        "Информатика\nИнформатика",
        "Физика\nФизика",
    ]
    
    # Hours/credits for each subject
    hours_credits = [
        "72с-3к",
        "108с-4к",
        "72с-3к",
        "108с-4.5к",
        "72с-3к",
    ]
    
    # Create DataFrame structure matching parser expectations:
    # Row 0: Optional header row
    # Row 1: Subject names (ROW_SUBJECT_NAMES = 1)
    # Row 2: Empty spacing row
    # Row 3: Hours/Credits (ROW_HOURS = 3)
    # Row 4: Empty spacing row  
    # Row 5+: Student data (ROW_DATA_START = 5)
    
    data = {}
    
    # Column 0: Row number
    data[0] = [
        "№",           # Row 0: Header  
        "",            # Row 1: Empty
        "",            # Row 2: Empty
        "",            # Row 3: Empty
        "",            # Row 4: Empty
    ] + list(range(1, 6))  # Rows 5-9: Student numbers
    
    # Column 1: Student names
    data[1] = [
        "Аты-жөні / ФИО",  # Row 0: Header
        "",                 # Row 1: Empty
        "",                 # Row 2: Empty
        "",                 # Row 3: Empty
        "",                 # Row 4: Empty
    ] + [
        "Иванов Иван Иванович",
        "Петров Петр Петрович",
        "Сидоров Сидор Сидорович",
        "Казиева Айгерім Нұрланқызы",
        "Смирнова Анна Андреевна",
    ]
    
    # Add subject columns (4 columns per subject: Points, Letter, GPA, Traditional)
    col_idx = 2
    for subject_idx, subject_name in enumerate(subjects):
        # Column for Points (first column of 4 for this subject)
        data[col_idx] = [
            "",                             # Row 0: Empty
            subject_name,                   # Row 1: Subject name (ROW_SUBJECT_NAMES)
            "",                             # Row 2: Empty spacing
            hours_credits[subject_idx],     # Row 3: Hours/Credits (ROW_HOURS)
            "Баллдар\nБаллы",              # Row 4: Points header
        ]
        
        # Columns for Letter, GPA, Traditional
        # Columns for Letter, GPA, Traditional
        data[col_idx + 1] = ["", "", "", "", "Әріп\nБуква"]
        data[col_idx + 2] = ["", "", "", "", "GPA"]
        data[col_idx + 3] = ["", "", "", "", "Дәстүрлі\nТрадиционная"]
        
        # Student grades (5 students, starting from row 5)
        sample_grades = [
            (85, "B+", "3.33", "жақсы"),
            (90, "A-", "3.67", "өте жақсы"),
            (78, "C+", "2.33", "қанағаттанарлық"),
            (95, "A", "4.0", "өте жақсы"),
            (88, "B", "3.0", "жақсы"),
        ]
        
        for student_idx, (points, letter, gpa, traditional) in enumerate(sample_grades):
            data[col_idx].append(points)           # Points column
            data[col_idx + 1].append(letter)       # Letter column
            data[col_idx + 2].append(gpa)          # GPA column
            data[col_idx + 3].append(traditional)  # Traditional column
        
        col_idx += 4
    
    # Create DataFrame
    df = pd.DataFrame(data)
    
    # Write to Excel
    output_file = Path(output_path)
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='3F-1', index=False, header=False)
    
    logger.info(f"  Created: {output_file}")
    logger.info(f"  Subjects: {len(subjects)}")
    logger.info(f"  Students: 5")
    
    return output_file


def parse_test_data(excel_file: Path) -> dict:
    """
    Parse test Excel file using ExcelParser.
    
    Args:
        excel_file: Path to test Excel file
        
    Returns:
        Dictionary with parsed data
    """
    logger.info("\nParsing test Excel file...")
    
    parser = ExcelParser(
        source_file=str(excel_file),
        row_subject_names=1,  # Row 1 (0-based) for subject names
        row_hours_credits=3,  # Row 3 (0-based) for hours/credits
        row_data_start=5,     # Row 5 (0-based) for first student
    )
    
    students = parser.parse("3F-1")
    
    logger.info(f"  Parsed {len(students)} students")
    
    for idx, student in enumerate(students, 1):
        logger.info(f"    {idx}. {student.full_name} ({len(student.grades)} grades)")
    
    return {
        'students': students,
        'parser': parser,
    }


def generate_sample_diplomas(students: list, parser: ExcelParser, output_dir: Path):
    """
    Generate sample diplomas for test students.
    
    Args:
        students: List of Student objects
        parser: ExcelParser instance with subject data
        output_dir: Output directory path
    """
    logger.info(f"\nGenerating diplomas to: {output_dir}")
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Get original subjects with bilingual names
    subjects = parser.get_subjects("3F-1")
    
    # Create generators for both languages
    gen_kz = DiplomaGenerator(Program.IT, Language.KZ, "2025-2026")
    gen_ru = DiplomaGenerator(Program.IT, Language.RU, "2025-2026")
    
    # Generate diplomas for all students
    for idx, student in enumerate(students, 1):
        logger.info(f"\n  Student {idx}: {student.full_name}")
        
        # Generate KZ diploma
        try:
            filename_kz = f"{student.full_name.replace(' ', '_')}_KZ_test.xlsx"
            filepath_kz = output_dir / filename_kz
            
            excel_kz = gen_kz.generate(student, subjects, "letter")
            
            with open(filepath_kz, "wb") as f:
                f.write(excel_kz)
            
            logger.info(f"    ✓ KZ: {filename_kz}")
        except Exception as e:
            logger.error(f"    ✗ KZ failed: {e}")
        
        # Generate RU diploma
        try:
            filename_ru = f"{student.full_name.replace(' ', '_')}_RU_test.xlsx"
            filepath_ru = output_dir / filename_ru
            
            excel_ru = gen_ru.generate(student, subjects, "letter")
            
            with open(filepath_ru, "wb") as f:
                f.write(excel_ru)
            
            logger.info(f"    ✓ RU: {filename_ru}")
        except Exception as e:
            logger.error(f"    ✗ RU failed: {e}")


def main():
    """Main execution function."""
    logger.info("=" * 70)
    logger.info("Test Data Creation and Diploma Generation")
    logger.info("=" * 70)
    
    # Step 1: Create test Excel
    test_excel = create_test_excel("test_grades.xlsx")
    
    # Step 2: Parse test data
    parsed = parse_test_data(test_excel)
    students = parsed['students']
    parser = parsed['parser']
    
    # Step 3: Generate sample diplomas
    output_dir = Path("test_output") / datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    generate_sample_diplomas(students, parser, output_dir)
    
    logger.info("\n" + "=" * 70)
    logger.info("✓ Complete!")
    logger.info(f"  Test Excel: {test_excel}")
    logger.info(f"  Output Dir: {output_dir}")
    logger.info("=" * 70)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        logger.error(f"Fatal error: {e}", exc_info=True)
        sys.exit(1)
