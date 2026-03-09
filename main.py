import os
import argparse
import pandas as pd
from typing import List

from configs import get_config
from src.parser import parse_excel_sheet
from src.generator import DiplomaGenerator
from core.bridge import build_diploma_pages
from core.consistency import assert_program_subject_mapping, validate_program_subject_mapping
import sys

# Force utf-8 for Windows console
if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def main():
    parser = argparse.ArgumentParser(description="Diploma Generator (Modular)")
    parser.add_argument("--source", type=str, required=True, help="Путь к исходному Excel-файлу")
    parser.add_argument("--group", type=str, required=True, choices=["3F", "3D"], help="Группа: 3F (IT) или 3D (Бухгалтеры)")
    parser.add_argument("--lang", type=str, default="ALL", choices=["KZ", "RU", "ALL"], help="Язык диплома")
    args = parser.parse_args()

    # Preflight: ensure parsing map and program layout are consistent
    assert_program_subject_mapping(args.group)
    report = validate_program_subject_mapping(args.group)
    if report["extra_count"] > 0:
        print(
            f"[WARN] Non-blocking mapping extras for {args.group}: "
            f"{report['extra_count']} (examples: {report['extra_examples'][:3]})"
        )
    
    source_file = args.source
    if not os.path.exists(source_file):
        print(f"File not found: {source_file}")
        return
        
    os.makedirs("output", exist_ok=True)
    
    # 1. Загружаем Excel используя Calamine (ОЧЕНЬ БЫСТРО)
    print(f"Loading '{source_file}' using calamine...", flush=True)
    try:
        xl = pd.ExcelFile(source_file, engine='calamine')
    except Exception as e:
        print(f"Failed to load with calamine, falling back to default: {e}")
        xl = pd.ExcelFile(source_file)
    
    target_prefix = args.group
    if target_prefix == "3F":
        target_prefix = "3Ғ"  # Автозамена на казахскую 'Ғ'
        
    target_sheets = [s for s in xl.sheet_names if s.startswith(target_prefix)]
    
    if not target_sheets:
        print(f"No sheets found starting with {target_prefix}")
        return
        
    langs_to_run = ["kz", "ru"] if args.lang == "ALL" else [args.lang.lower()]
    
    for sheet_name in target_sheets:
        print(f"\nProcessing sheet: {sheet_name}", flush=True)
        # Parse fully using calamine DataFrame extractor
        df = xl.parse(sheet_name, header=None)
        
        # Limit processing range for speed (simulating openpyxl's max_row=200, max_col=300)
        df = df.iloc[:200, :300]
        
        students = parse_excel_sheet(df, sheet_name, start_row=5)
        print(f"  Found {len(students)} students.")
        
        for lang in langs_to_run:
            print(f"  Generating {lang.upper()} diplomas...")
            config, terms, template_name = get_config(args.group, lang)
            template_path = os.path.join("templates", template_name)
            
            if not os.path.exists(template_path):
                print(f"  [ERROR] Template {template_path} not found. Skipping {lang}.")
                continue
                
            for s in students:
                safe_name = s['name'].replace('/', ' ').replace('\\', ' ')
                out_name = f"{sheet_name}_{safe_name}_{lang.upper()}.xlsx"
                out_path = os.path.join("output", out_name)
                
                try:
                    # Bridge: map parsed grades onto PROGRAM_PAGES structure
                    structured = build_diploma_pages(s['grades'], args.group)

                    generator = DiplomaGenerator(template_path, out_path, config, terms)
                    generator.fill_from_pages(s, structured, lang=lang)
                    generator.close()
                    print(f"    + {out_name}", flush=True)
                except Exception as e:
                    print(f"    - [ERROR] Failed to generate {out_name}: {e}")

if __name__ == "__main__":
    main()
