import os
import shutil
from parse_grades import parse_workbook, SOURCE_FILE
from generate_diploma import (
    generate_diploma, normalize_key, is_module_header,
    PAGE1_SUBJECTS, PAGE2_SUBJECTS, PAGE3_SUBJECTS, PAGE4_SUBJECTS
)
import re

OUTPUT_DIR = "output_diplomas"

def extract_code(text):
    """Extract code like 'ОН 1.2', 'КМ 1', 'БМ 01'."""
    match = re.match(r"^([A-Za-zА-Яа-яЁё]+\s+[\d\.]+(?:\.\d+)?)", text.strip())
    if match:
        return normalize_key(match.group(1))
    return None

def aggregate_module_grades(grades, all_subjects):
    """Calculate totals for KM/Module headers based on their children."""
    kv_map = grades
    active_km = None
    accum_hours = 0
    accum_credits = 0
    
    def parse_num(val):
        try:
            return float(str(val).replace(',', '.'))
        except (ValueError, TypeError):
            return 0

    for subject in all_subjects:
        # Start of a new module?
        if subject.startswith("КМ ") or subject.startswith("Кәсіптік модуль"):
            # Save previous if exists
            if active_km is not None:
                if active_km not in kv_map: kv_map[active_km] = {}
                kv_map[active_km]["hours"] = int(accum_hours) if accum_hours % 1 == 0 else accum_hours
                kv_map[active_km]["credits"] = int(accum_credits) if accum_credits % 1 == 0 else accum_credits
                # Ensure no grades are set for the header
                kv_map[active_km]["points"] = ""
                kv_map[active_km]["letter"] = ""
                kv_map[active_km]["gpa"] = ""
                kv_map[active_km]["traditional"] = ""

            # Reset for new module
            active_km = subject
            accum_hours = 0
            accum_credits = 0
            
        elif active_km:
            # Accumulate if it's a child subject (OH ...)
            if subject.startswith("ОН ") or subject.startswith("БМ "):
                g = kv_map.get(subject, {})
                h = parse_num(g.get("hours", 0))
                c = parse_num(g.get("credits", 0))
                accum_hours += h
                accum_credits += c

    # Save last one
    if active_km is not None:
        if active_km not in kv_map: kv_map[active_km] = {}
        kv_map[active_km]["hours"] = int(accum_hours) if accum_hours % 1 == 0 else accum_hours
        kv_map[active_km]["credits"] = int(accum_credits) if accum_credits % 1 == 0 else accum_credits

import sys

def main():
    sys.stdout.reconfigure(encoding='utf-8')
    
    if not os.path.exists(SOURCE_FILE):
        print(f"Error: Source file '{SOURCE_FILE}' not found.")
        return

    print("Parsing workbook... this may take a moment.")
    students = parse_workbook(SOURCE_FILE)
    print(f"Parsed {len(students)} students.")

    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)

    # Collect all generator keys in PRECISE ORDER for aggregation
    all_gen_keys = (PAGE1_SUBJECTS + PAGE2_SUBJECTS + 
                    PAGE3_SUBJECTS + PAGE4_SUBJECTS)
    
    # 1. Create a "Code" map for robust matching
    gen_code_map = {}
    for k in all_gen_keys:
        code = extract_code(k)
        if code:
            gen_code_map[code] = k
    
    # Also keep standard map
    gen_key_map = {normalize_key(k): k for k in all_gen_keys}

    for i, student in enumerate(students):
        full_name = student['full_name']
        safe_name = full_name.replace(" ", "_").replace(".", "")
        filename = f"Diplom_{safe_name}.xlsx"
        output_path = os.path.join(OUTPUT_DIR, filename)
        
        # KEY RE-MAPPING
        parsed_grades = student['grades']
        aligned_grades = {}
        
        for p_key, p_data in parsed_grades.items():
            norm_p_key = normalize_key(p_key)
            p_code = extract_code(p_key)
            target_key = None
            
            # 1. Check Code
            if p_code and p_code in gen_code_map:
                target_key = gen_code_map[p_code]
            # 2. Check Exact keys
            elif norm_p_key in gen_key_map:
                target_key = gen_key_map[norm_p_key]
            # 3. Fuzzy Fallback
            if not target_key:
                for g_norm, g_key in gen_key_map.items():
                   if g_norm.startswith(norm_p_key) or norm_p_key.startswith(g_norm):
                       target_key = g_key
                       break
            
            if target_key:
                aligned_grades[target_key] = p_data

        # Update student data
        student['grades'] = aligned_grades
        
        # AGGREGATE MODULES
        aggregate_module_grades(student['grades'], all_gen_keys)
        
        # PREPARE RUSSIAN GRADES (Map KZ keys -> RU keys by index)
        # Import RU subjects dynamically to allow this script to run even if ru generator changes
        import generate_diploma_ru
        
        ru_all_keys = (generate_diploma_ru.PAGE1_SUBJECTS + 
                       generate_diploma_ru.PAGE2_SUBJECTS + 
                       generate_diploma_ru.PAGE3_SUBJECTS + 
                       generate_diploma_ru.PAGE4_SUBJECTS)
        
        # Map KZ -> RU
        # Ensure lists are same length
        if len(all_gen_keys) == len(ru_all_keys):
            kz_to_ru_map = dict(zip(all_gen_keys, ru_all_keys))
            grades_ru = {}
            for k, v in student['grades'].items():
                if k in kz_to_ru_map:
                    grades_ru[kz_to_ru_map[k]] = v
            
            # Prepare RU student data
            student_ru = student.copy()
            student_ru['grades'] = grades_ru
            # Add specific RU headers if needed (e.g. College Name)
            # generate_diploma_ru handles defaults, but we can pass explicit if available
            student_ru['college_name_ru'] = "Жамбылском инновационным высшем колледже"
            student_ru['full_name_ru'] = student['full_name'] # Or assume same? Usually names are same in Latin/Cyrillic or need translit? 
            # The source Excel has only one name column. We use it for both.
            
            filename_ru = f"Diplom_RU_{safe_name}.xlsx"
            path_ru = os.path.join(OUTPUT_DIR, filename_ru)
            
            print(f"Generating RU diploma for: {full_name}")
            try:
                generate_diploma_ru.generate_diploma_ru(data=student_ru, output_path=path_ru)
            except Exception as e:
                print(f"Failed to generate RU: {e}")
        else:
            print(f"Warning: KZ and RU subject lists have different lengths ({len(all_gen_keys)} vs {len(ru_all_keys)}). Skipping RU generation.")

        print(f"Generating diploma for: {full_name}")
        try:
            generate_diploma(data=student, output_path=output_path)
        except Exception as e:
            print(f"Failed to generate for {full_name}: {e}")

    print(f"\nProcessing complete. Check {OUTPUT_DIR}/")

if __name__ == "__main__":
    main()
