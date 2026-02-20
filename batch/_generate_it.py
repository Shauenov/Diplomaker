#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
IT Program Batch Diploma Generator
===================================
Generate all diplomas for IT 3F groups (3F-1, 3F-2, 3F-3, 3F-4).

Phase 2 unified batch processor using new architecture:
- Parse source Excel using data.excel_parser
- Extract subject names and hours/credits from source
- Generate diplomas using data.excel_generator
- Output to organized directory
"""

import sys
from pathlib import Path
from datetime import datetime
import logging

from config.settings import OUTPUT_DIR
from core.models import Program, Language
from data.excel_parser import ExcelParser
from data.excel_generator import DiplomaGenerator


# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)
logger = logging.getLogger(__name__)


def main():
    """Main batch processing function."""
    logger.info("=" * 70)
    logger.info("IT Program Diploma Generator (Phase 2 - Unified Architecture)")
    logger.info("=" * 70)

    # Create output directory
    batch_name = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    batch_dir = OUTPUT_DIR / batch_name
    batch_dir.mkdir(parents=True, exist_ok=True)

    logger.info(f"Output directory: {batch_dir}\n")

    # Initialize parser and generators
    parser = ExcelParser(program=Program.IT)
    gen_kz = DiplomaGenerator(Program.IT, Language.KZ, "2025-2026")
    gen_ru = DiplomaGenerator(Program.IT, Language.RU, "2025-2026")

    total_students = 0
    successful = 0
    failed_students = []

    # Process all sheets
    try:
        all_students = parser.parse_all_sheets()

        for sheet_name, students in all_students.items():
            logger.info(f"\nProcessing {len(students)} students from '{sheet_name}'...")

            # Parse subjects once per sheet
            df = parser._load_dataframe(sheet_name)
            subjects = parser._parse_subjects(df)
            logger.info(f"  Found {len(subjects)} subjects")

            for idx, student in enumerate(students, 1):
                try:
                    # Generate KZ diploma
                    filename_kz = (
                        f"{student.full_name.replace(' ', '_')}_"
                        f"{sheet_name}_KZ_2025-2026.xlsx"
                    )
                    filepath_kz = batch_dir / filename_kz

                    excel_kz = gen_kz.generate(student, subjects, "letter")
                    with open(filepath_kz, "wb") as f:
                        f.write(excel_kz)

                    # Generate RU diploma
                    filename_ru = filename_kz.replace("_KZ_", "_RU_")
                    filepath_ru = batch_dir / filename_ru

                    excel_ru = gen_ru.generate(student, subjects, "letter")
                    with open(filepath_ru, "wb") as f:
                        f.write(excel_ru)

                    successful += 1
                    total_students += 1

                    if idx % 5 == 0 or idx == len(students):
                        logger.info(
                            f"  [{idx}/{len(students)}] ✓ {student.full_name} "
                            f"(2 diplomas)"
                        )

                except Exception as e:
                    failed_students.append((student.full_name, str(e)))
                    logger.warning(f"  ✗ {student.full_name}: {str(e)}")
                    total_students += 1

        # Print summary
        logger.info("\n" + "=" * 70)
        logger.info("BATCH PROCESSING SUMMARY")
        logger.info("=" * 70)
        logger.info(f"Total students: {total_students}")
        logger.info(f"Successful: {successful}")
        logger.info(f"Failed: {len(failed_students)}")
        logger.info(
            f"Success rate: {100 * successful / max(1, total_students):.1f}%"
        )
        logger.info(f"Output directory: {batch_dir}")

        if failed_students:
            logger.info("\nFailed students:")
            for name, error in failed_students:
                logger.info(f"  - {name}: {error}")

        logger.info("=" * 70 + "\n")

        return 0 if len(failed_students) == 0 else 1

    except Exception as e:
        logger.error(f"Fatal error during batch processing: {str(e)}")
        return 2


if __name__ == "__main__":
    sys.exit(main())
