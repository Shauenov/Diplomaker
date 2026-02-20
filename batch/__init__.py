"""
Batch Processing Module
=======================
Orchestrates diploma generation workflow for multiple students and sheets.

Handles:
- Reading source Excel files
- Processing student batches  
- Generating bilingual diplomas
- File output and logging
- Error reporting
"""

import logging
import sys
from pathlib import Path
from typing import List, Dict, Optional
from datetime import datetime

from config.settings import OUTPUT_DIR, DEBUG_MODE, LOG_FILE
from core.models import Program, Language
from core.models import Subject, Student, ProcessingResult
from core.exceptions import DiplomaAutomationError
from data.excel_parser import ExcelParser
from data.excel_generator import DiplomaGenerator


# Configure logging
logging.basicConfig(
    level=logging.DEBUG if DEBUG_MODE else logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE),
        logging.StreamHandler(sys.stdout),
    ],
)
logger = logging.getLogger(__name__)


class BatchProcessor:
    """
    Process multiple students and generate diplomas.

    Parameters:
        program (Program): Program code (IT, ACCOUNTING)
        batch_name (str, optional): Name for this batch (for logging)
        output_dir (Path, optional): Override output directory

    Example:
        >>> processor = BatchProcessor(Program.IT)
        >>> result = processor.process_all_sheets()
        >>> print(f"Generated {result.successful}/{result.total_students} diplomas")
    """

    def __init__(
        self,
        program: Program = Program.IT,
        batch_name: Optional[str] = None,
        output_dir: Optional[Path] = None,
    ):
        """Initialize batch processor."""
        self.program = program
        self.output_dir = Path(output_dir) if output_dir else OUTPUT_DIR
        self.output_dir.mkdir(parents=True, exist_ok=True)

        # Create batch subdirectory
        self.batch_name = batch_name or datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        self.batch_dir = self.output_dir / self.batch_name
        self.batch_dir.mkdir(parents=True, exist_ok=True)

        logger.info(f"Initialized BatchProcessor for {program.value}")
        logger.info(f"Output directory: {self.batch_dir}")

    def process_all_sheets(self) -> ProcessingResult:
        """
        Process all sheets for the program.

        Returns:
            ProcessingResult with statistics
        """
        logger.info(f"Starting batch processing for {self.program.value}")
        result = ProcessingResult(
            total_students=0,
            successful=0,
            failed=0,
            errors=[],
            warnings=[],
            statistics={},
        )

        try:
            parser = ExcelParser(program=self.program)
            all_sheets = parser.parse_all_sheets()

            total_sheets = len(all_sheets)
            logger.info(f"Parsed {total_sheets} sheets")

            # Process each sheet
            for sheet_idx, (sheet_name, students) in enumerate(all_sheets.items(), 1):
                logger.info(f"[{sheet_idx}/{total_sheets}] Processing sheet '{sheet_name}'")

                # Process students in this sheet
                sheet_result = self._process_sheet(sheet_name, students)

                # Accumulate results
                result.total_students += sheet_result["total"]
                result.successful += sheet_result["successful"]
                result.failed += sheet_result["failed"]
                result.errors.extend(sheet_result["errors"])

                logger.info(
                    f"Sheet '{sheet_name}': "
                    f"{sheet_result['successful']}/{sheet_result['total']} successful"
                )

        except Exception as e:
            error_msg = f"Fatal error: {str(e)}"
            logger.error(error_msg)
            result.errors.append(error_msg)

        # Log final summary
        result.statistics = {
            "batch_name": self.batch_name,
            "program": self.program.value,
            "timestamp": datetime.now().isoformat(),
            "output_dir": str(self.batch_dir),
        }

        logger.info(_format_summary(result))
        return result

    def _process_sheet(self, sheet_name: str, students: List[Student]) -> Dict:
        """
        Process all students in a sheet.

        Returns:
            Dict with sheet statistics
        """
        successful = 0
        failed = 0
        errors = []

        # Get subject definitions from parser
        parser = ExcelParser(program=self.program)
        df_subjects = parser._parse_subjects(
            parser._load_dataframe(sheet_name)
        ) if hasattr(parser, '_load_dataframe') else []

        # Generators for both languages
        gen_kz = DiplomaGenerator(self.program, Language.KZ)
        gen_ru = DiplomaGenerator(self.program, Language.RU)

        # Process each student
        for student_idx, student in enumerate(students, 1):
            try:
                # Generate KZ diploma
                filename_kz = (
                    f"{student.full_name.replace(' ', '_')}_KZ_" 
                    f"{datetime.now().year}-{datetime.now().year + 1}.xlsx"
                )
                filepath_kz = self.batch_dir / filename_kz

                excel_kz = gen_kz.generate(student, df_subjects)
                with open(filepath_kz, "wb") as f:
                    f.write(excel_kz)

                # Generate RU diploma
                filename_ru = (
                    f"{student.full_name.replace(' ', '_')}_RU_" 
                    f"{datetime.now().year}-{datetime.now().year + 1}.xlsx"
                )
                filepath_ru = self.batch_dir / filename_ru

                excel_ru = gen_ru.generate(student, df_subjects)
                with open(filepath_ru, "wb") as f:
                    f.write(excel_ru)

                successful += 1
                logger.debug(
                    f"  [{student_idx}] ✓ {student.full_name}: "
                    f"2 diplomas generated"
                )

            except Exception as e:
                failed += 1
                error_msg = f"{student.full_name}: {str(e)}"
                errors.append(error_msg)
                logger.warning(f"  [{student_idx}] ✗ {error_msg}")

        return {
            "total": len(students),
            "successful": successful,
            "failed": failed,
            "errors": errors,
        }


def _format_summary(result: ProcessingResult) -> str:
    """Format processing result for logging."""
    return (
        f"\n{'=' * 60}\n"
        f"BATCH PROCESSING SUMMARY\n"
        f"{'=' * 60}\n"
        f"Total students: {result.total_students}\n"
        f"Successful: {result.successful}\n"
        f"Failed: {result.failed}\n"
        f"Success rate: {100 * result.successful / max(1, result.total_students):.1f}%\n"
        f"Output directory: {result.statistics.get('output_dir', 'N/A')}\n"
        f"Timestamp: {result.statistics.get('timestamp', 'N/A')}\n"
        f"{'=' * 60}"
    )


def process_it_program(
    batch_name: Optional[str] = None,
    output_dir: Optional[Path] = None,
) -> ProcessingResult:
    """
    Process entire IT program (3F-1 through 3F-4).

    Parameters:
        batch_name (str, optional): Custom batch name
        output_dir (Path, optional): Override output directory

    Returns:
        ProcessingResult with statistics
    """
    processor = BatchProcessor(Program.IT, batch_name, output_dir)
    return processor.process_all_sheets()


def process_accounting_program(
    batch_name: Optional[str] = None,
    output_dir: Optional[Path] = None,
) -> ProcessingResult:
    """
    Process entire Accounting program (3D-1 through 3D-3).

    Parameters:
        batch_name (str, optional): Custom batch name
        output_dir (Path, optional): Override output directory

    Returns:
        ProcessingResult with statistics
    """
    processor = BatchProcessor(Program.ACCOUNTING, batch_name, output_dir)
    return processor.process_all_sheets()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(
        description="Generate diplomas for all students in a program."
    )
    parser.add_argument(
        "--program",
        choices=["IT", "ACCOUNTING"],
        default="IT",
        help="Program to process (default: IT)",
    )
    parser.add_argument(
        "--batch-name",
        help="Custom batch name for output directory",
    )
    parser.add_argument(
        "--output-dir",
        help="Override output directory",
    )

    args = parser.parse_args()

    program = Program.IT if args.program == "IT" else Program.ACCOUNTING
    result = process_it_program(args.batch_name, Path(args.output_dir) if args.output_dir else None)

    sys.exit(0 if result.failed == 0 else 1)
