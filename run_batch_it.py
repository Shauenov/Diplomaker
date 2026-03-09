#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
Batch IT Diploma Generator (Legacy Stack)
==========================================
Generates KZ+RU diplomas for all IT groups (3Ғ-1..3Ғ-4)
using the legacy openpyxl-based template filling approach.

Usage:
    python run_batch_it.py
    python run_batch_it.py --source "path/to/grades.xlsx"
    python run_batch_it.py --lang KZ
    python run_batch_it.py --lang RU
"""

import os
import sys
import argparse
from datetime import datetime
import pandas as pd

from configs import get_config
from src.parser import parse_excel_sheet
from src.generator import DiplomaGenerator
from core.bridge import build_diploma_pages
from core.consistency import assert_program_subject_mapping, validate_program_subject_mapping
from config.settings import SOURCE_FILE


def main():
    parser = argparse.ArgumentParser(description="Batch IT Diploma Generator")
    parser.add_argument("--source", type=str, default=str(SOURCE_FILE),
                        help="Путь к исходному Excel-файлу с оценками")
    parser.add_argument("--lang", type=str, default="ALL", choices=["KZ", "RU", "ALL"],
                        help="Язык диплома: KZ, RU или ALL (оба)")
    parser.add_argument("--output", type=str, default=None,
                        help="Папка для вывода (по умолчанию: output_diplomas/<timestamp>)")
    args = parser.parse_args()

    # Preflight: ensure parsing map and program layout are consistent
    assert_program_subject_mapping("3F")
    report = validate_program_subject_mapping("3F")
    if report["extra_count"] > 0:
        print(
            f"[WARN] Non-blocking mapping extras for 3F: "
            f"{report['extra_count']} (examples: {report['extra_examples'][:3]})"
        )

    source_file = args.source
    if not os.path.exists(source_file):
        print(f"[ERROR] Исходный файл не найден: {source_file}")
        sys.exit(1)

    # Создаём папку для вывода
    if args.output:
        output_dir = args.output
    else:
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        output_dir = os.path.join("output_diplomas", timestamp)
    os.makedirs(output_dir, exist_ok=True)

    print("=" * 70)
    print("  IT Program Diploma Generator (Batch)")
    print("=" * 70)
    print(f"  Источник: {source_file}")
    print(f"  Вывод:    {output_dir}")
    print()

    # Загружаем Excel
    xl = pd.ExcelFile(source_file)

    # Ищем листы 3Ғ-*
    target_sheets = [s for s in xl.sheet_names if s.startswith("3Ғ")]
    if not target_sheets:
        # Попробуем латинскую F
        target_sheets = [s for s in xl.sheet_names if s.startswith("3F")]

    if not target_sheets:
        print(f"[ERROR] Не найдены листы 3Ғ-* или 3F-* в файле.")
        print(f"  Доступные листы: {xl.sheet_names}")
        sys.exit(1)

    print(f"  Найдены листы: {target_sheets}")

    langs_to_run = ["kz", "ru"] if args.lang == "ALL" else [args.lang.lower()]

    total_students = 0
    total_generated = 0
    total_errors = 0

    for sheet_name in target_sheets:
        print(f"\n{'─' * 50}")
        print(f"  Лист: {sheet_name}")
        print(f"{'─' * 50}")

        df = xl.parse(sheet_name=sheet_name, header=None)
        students = parse_excel_sheet(df, sheet_name, start_row=5)
        print(f"  Найдено студентов: {len(students)}")

        if not students:
            print("  [WARN] Нет студентов на этом листе, пропуск.")
            continue

        total_students += len(students)

        for lang in langs_to_run:
            config, terms, template_name = get_config("3F", lang)
            template_path = os.path.join("templates", template_name)

            if not os.path.exists(template_path):
                print(f"  [ERROR] Шаблон не найден: {template_path}")
                continue

            print(f"\n  Генерация {lang.upper()} дипломов (шаблон: {template_name})...")

            for idx, student in enumerate(students, 1):
                safe_name = student['name'].replace('/', ' ').replace('\\', ' ').strip()
                out_name = f"{sheet_name}_{safe_name}_{lang.upper()}.xlsx"
                out_path = os.path.join(output_dir, out_name)

                try:
                    # Bridge: map parsed grades onto PROGRAM_PAGES structure
                    structured = build_diploma_pages(student['grades'], "3F")

                    generator = DiplomaGenerator(template_path, out_path, config, terms)
                    generator.fill_from_pages(student, structured, lang=lang)
                    generator.close()
                    total_generated += 1

                    # Краткий отчёт: имя и количество предметов с оценками
                    grades = student.get('grades', {})
                    graded = sum(1 for g in grades.values() if g.get('points'))
                    print(f"    [{idx:2d}] ✓ {safe_name} ({graded} оценок)")

                except Exception as e:
                    total_errors += 1
                    print(f"    [{idx:2d}] ✗ {safe_name}: {e}")

    # Итоговый отчёт
    print(f"\n{'=' * 70}")
    print(f"  ИТОГО:")
    print(f"    Студентов: {total_students}")
    print(f"    Сгенерировано: {total_generated}")
    print(f"    Ошибок: {total_errors}")
    print(f"    Папка: {output_dir}")
    print(f"{'=' * 70}")


if __name__ == "__main__":
    main()
