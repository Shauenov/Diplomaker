import pandas as pd
import os
import re
import copy
import argparse
from generate_diploma_it_kz import generate_diploma as generate_kz
from generate_diploma_it_ru import generate_diploma as generate_ru
from parse_grades import calculate_grade_details

# ─────────────────────────────────────────────────────────────
# HARDCODED CONFIGURATION — do NOT change source file
# ─────────────────────────────────────────────────────────────
SOURCE_FILE = r"c:\Users\user\OneDrive\Рабочий стол\template\2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"

# All F-group (IT) sheets in the source file
IT_SHEETS = ["3Ғ-1", "3Ғ-2", "3Ғ-3", "3Ғ-4"]

OUTPUT_DIR = "Diplomas_Batch"

# Row Indices (0-based) for real source file layout:
#   Row 1 (idx 0): subject numbers
#   Row 2 (idx 1): subject names  (KZ\nRU bilingual)
#   Row 3 (idx 2): student-name label
#   Row 4 (idx 3): hours/credits  ("72с-3к")
#   Row 5 (idx 4): column labels  п / б / цэ / трад
#   Row 6 (idx 5): first student  ← ROW_DATA_START
ROW_SUBJECT_NAMES = 1       # Excel Row 2
ROW_HOURS        = 3        # Excel Row 4
ROW_DATA_START   = 5        # Excel Row 6 — students begin here

# Column Indices (0-based)
COL_NO             = 0      # Column A — row number
COL_FULL_NAME      = 1      # Column B — full name
COL_START_SUBJECTS = 2      # Column C — first subject's 'п' (points) column
# Each subject occupies 4 columns: п (points), б (letter), цэ (GPA), трад (traditional)
# We read ONLY the 'п' column and derive the rest internally.

def parse_hours_credits(text):
    """Parse '72с-3к' into (72, 3)."""
    if not isinstance(text, str) or text.lower() == "nan":
        return "", ""
    match = re.search(r"(\d+)с-(\d+(?:,\d+)?)к", text)
    if match:
        return match.group(1), match.group(2)
    return text, ""

def clean_subject_name(text):
    """Split 'NameKZ\nNameRU' into ('NameKZ', 'NameRU'), stripping colons."""
    if not isinstance(text, str):
        return str(text).strip(), str(text).strip()
    parts = text.split('\n')
    if len(parts) >= 2:
        return parts[0].strip().rstrip(':').strip(), parts[1].strip().rstrip(':').strip()
    return text.strip().rstrip(':').strip(), text.strip().rstrip(':').strip()

def normalize_key(text):
    """Normalize a subject name to a compact lowercase key."""
    if not text:
        return ""
    t = str(text).lower()
    t = t.replace(".", "").replace(",", "").replace(":", "")
    t = t.replace(" ", "")
    t = re.sub(r'([a-zа-я]+)0+([1-9]+)', r'\1\2', t)
    return t.strip()


def robust_clean(val):
    """Convert a cell value to a clean string; discard NaN / zero / REF errors."""
    if pd.isna(val) or str(val).lower() == "nan" or str(val).strip() == "0" or str(val) == "#REF!":
        return ""
    return str(val).strip()


def derive_grade_obj(hours, credits, raw_points):
    """
    Build a grade dict from raw points only.
    Letter, GPA, and Traditional grades are derived INTERNALLY — never read
    from the source Excel file.
    """
    pts_clean = robust_clean(raw_points)

    # Calculate derived fields from points
    derived = calculate_grade_details(pts_clean) if pts_clean else {
        "letter": "", "gpa": "", "traditional_kz": "", "traditional_ru": ""
    }

    return {
        "hours":          robust_clean(hours),
        "credits":        robust_clean(credits),
        "points":         pts_clean,
        "letter":         derived.get("letter", ""),
        "gpa":            str(derived.get("gpa", "")) if derived.get("gpa", "") != "" else "",
        "traditional_kz": derived.get("traditional_kz", ""),
        "traditional_ru": derived.get("traditional_ru", ""),
        # 'traditional' is set per-language by adapt_grades_for_lang()
        "traditional":    derived.get("traditional_kz", ""),
    }


def adapt_grades_for_lang(grades, lang):
    """
    Return a shallow-adapted copy of the grades dict where every entry has
    'traditional' set to the correct language variant (KZ or RU).
    """
    key = "traditional_kz" if lang == "kz" else "traditional_ru"
    adapted = {}
    for subject, g in grades.items():
        entry = dict(g)
        entry["traditional"] = g.get(key, g.get("traditional", ""))
        adapted[subject] = entry
    return adapted


def get_student_row_data(df, row_idx, subject_columns):
    """
    Extract grades for a specific student row.
    ONLY the 'п' (points/percentage) column is read from the source file.
    Letter, GPA, and Traditional grades are computed internally.
    """
    student_name = df.iloc[row_idx, COL_FULL_NAME]

    # Skip empty rows
    if pd.isna(student_name) or str(student_name).strip() == "":
        return None

    # ── Diploma ID ────────────────────────────────────────────
    div_id = None
    val_no = df.iloc[row_idx, COL_NO]

    # Scan far-right columns for a JB/KZ-prefixed diploma number
    for c in range(200, df.shape[1]):
        val = df.iloc[row_idx, c]
        if isinstance(val, str) and (val.startswith("JB") or val.startswith("KZ")):
            div_id = val
            break

    if not div_id:
        if pd.notna(val_no):
            try:
                num = int(float(val_no))
                div_id = str(num).zfill(7)
            except Exception:
                div_id = str(val_no)
        else:
            div_id = str(row_idx).zfill(7)

    div_id_clean = div_id.replace("JB", "").replace("KZ", "").strip()

    # ── Build grades dict ─────────────────────────────────────
    grades = {}

    for subj in subject_columns:
        col_idx = subj["col_idx"]

        try:
            # READ ONLY the points ('п') column — ignore б / цэ / трад from Excel
            raw_points = df.iloc[row_idx, col_idx]
        except IndexError:
            raw_points = ""

        grade_obj = derive_grade_obj(subj["hours"], subj["credits"], raw_points)

        # Register under all key variants so the generators can find the subject
        for key in (
            subj["name_kz"],
            subj["name_ru"],
            subj["name_kz"].strip(),
            subj["name_ru"].strip(),
            normalize_key(subj["name_kz"]),
            normalize_key(subj["name_ru"]),
        ):
            grades[key] = grade_obj

        # Also store under module-level alias names (row 2 when row 3 was primary)
        for alias in subj.get("extra_kz_keys", []):
            if alias:
                grades[alias] = grade_obj
                grades[normalize_key(alias)] = grade_obj
        for alias in subj.get("extra_ru_keys", []):
            if alias:
                grades[alias] = grade_obj
                grades[normalize_key(alias)] = grade_obj

    # ── Hardcoded Electives (no numeric grade — сынақ / зачтено) ─
    ELECTIVES_KZ = [
        "Ф1 Факультативтік ағылшын тілі",
        "Ф2 Факультативтік түрік тілі",
        "Ф3 Факультативтік кәсіпкерлік қызмет негіздері",
    ]
    ELECTIVES_RU = [
        "Факультатив английский язык",
        "Ф1 Факультатив английский язык",
        "Факультатив турецкий язык",
        "Ф2 Факультатив турецкий язык",
        "Факультатив основы предпринимательской деятельности",
        "Ф3 Факультатив основы предпринимательской деятельности",
    ]
    ELECTIVE_BASE = {"hours": "36", "credits": "1.5", "points": "", "letter": "", "gpa": ""}

    for subj in ELECTIVES_KZ:
        grades[subj] = {**ELECTIVE_BASE, "traditional": "сынақ",  "traditional_kz": "сынақ",  "traditional_ru": "зачтено"}
    for subj in ELECTIVES_RU:
        grades[subj] = {**ELECTIVE_BASE, "traditional": "зачтено", "traditional_kz": "сынақ", "traditional_ru": "зачтено"}
    
    # ── Attestation (if no hours in source, use default) ─
    ATTESTATION_KZ = "Қорытынды аттестаттау"
    ATTESTATION_RU = "Итоговая аттестация"
    if ATTESTATION_KZ not in grades or (not grades[ATTESTATION_KZ].get("hours")):
        grades[ATTESTATION_KZ] = {"hours": "108", "credits": "4.5", "points": "", "letter": "", "gpa": "", "traditional": "", "traditional_kz": "", "traditional_ru": ""}
    if ATTESTATION_RU not in grades or (not grades[ATTESTATION_RU].get("hours")):
        grades[ATTESTATION_RU] = {"hours": "108", "credits": "4.5", "points": "", "letter": "", "gpa": "", "traditional": "", "traditional_kz": "", "traditional_ru": ""}

    return {
        "full_name":          student_name,
        "diploma_number":     div_id,
        "diploma_number_clean": div_id_clean,
        "grades":             grades,
    }

def main():
    parser = argparse.ArgumentParser(description="Generate IT diplomas for all F-group students.")
    parser.add_argument("--source", default=None,
                        help="Override source Excel file (default: hardcoded 2025-2026 grades file)")
    parser.add_argument("--out", default=OUTPUT_DIR,
                        help=f"Output directory (default: {OUTPUT_DIR})")
    args = parser.parse_args()

    source_file = args.source if args.source else SOURCE_FILE
    out_dir = args.out

    if not os.path.exists(out_dir):
        os.makedirs(out_dir)

    total_processed = 0

    for sheet_name in IT_SHEETS:
        print(f"\n{'='*60}")
        print(f"Processing sheet: {sheet_name}")
        print(f"{'='*60}")
        print(f"Loading {source_file}...")
        try:
            df = pd.read_excel(source_file, sheet_name=sheet_name, header=None)
        except Exception as e:
            print(f"  [SKIP] Could not load sheet '{sheet_name}': {e}")
            continue

        # 1. Parse Subjects from header rows
        print("  Parsing subjects...")
        subject_columns = []

        row_hours_data = df.iloc[ROW_HOURS]
        ROW_SUB_NAMES = ROW_SUBJECT_NAMES + 1  # Row 3 (0-based idx 2): ОН/РО sub-subject names

        for col_idx in range(COL_START_SUBJECTS, df.shape[1], 4):  # 4 cols per subject
            raw_r2 = df.iloc[ROW_SUBJECT_NAMES, col_idx]  # row 2: module/section name
            raw_r3 = df.iloc[ROW_SUB_NAMES, col_idx] if ROW_SUB_NAMES < len(df) else None  # row 3: sub-subject name

            r2_str = str(raw_r2).strip() if raw_r2 is not None and not pd.isna(raw_r2) and str(raw_r2).strip() else ""
            r3_str = str(raw_r3).strip() if raw_r3 is not None and not pd.isna(raw_r3) and str(raw_r3).strip() else ""

            if not r2_str and not r3_str:
                continue  # completely empty column

            # Row 3 (sub-subject) takes priority as primary name over row 2 (module header)
            primary_raw = r3_str if r3_str else r2_str
            name_kz, name_ru = clean_subject_name(primary_raw)

            raw_hours_val = row_hours_data[col_idx]
            h_raw_str = str(raw_hours_val).strip()
            if pd.isna(raw_hours_val) or h_raw_str.lower() == "nan" or h_raw_str == "":
                hours, credits = "", ""
            else:
                hours, credits = parse_hours_credits(h_raw_str)

            entry = {
                "col_idx":       col_idx,
                "name_kz":       name_kz,
                "name_ru":       name_ru,
                "hours":         hours,
                "credits":       credits,
                "extra_kz_keys": [],
                "extra_ru_keys": [],
            }

            # When both row 2 and row 3 carry different names,
            # also store under the row-2 (module) name as an alias.
            if r2_str and r3_str:
                alias_kz, alias_ru = clean_subject_name(r2_str)
                entry["extra_kz_keys"].append(alias_kz)
                entry["extra_ru_keys"].append(alias_ru)

            subject_columns.append(entry)

        print(f"  Found {len(subject_columns)} subjects.")

        # 2. Process Students
        print("  Processing students...")
        sheet_processed = 0

        for row_idx in range(ROW_DATA_START, len(df)):
            student_data = get_student_row_data(df, row_idx, subject_columns)
            if not student_data:
                continue

            sheet_processed += 1
            total_processed += 1
            name = student_data["full_name"]

            # Debug dump for first student of each sheet
            if sheet_processed == 1:
                import json
                debug_path = f"debug_grades_{sheet_name.replace('Ғ', 'F')}.json"
                with open(debug_path, "w", encoding="utf-8") as f:
                    json.dump(student_data["grades"], f, ensure_ascii=False, indent=2)
                print(f"  [DEBUG] Grade dump: {debug_path}")

            print(f"  Generating: {name} ({student_data['diploma_number']})")
            sanitized_name = "".join([c for c in name if c.isalnum() or c in (' ', '.', '_')]).strip()
            sheet_label = sheet_name.replace("Ғ", "F")

            # Adapt grades: KZ version uses KZ traditional grades, RU version uses RU
            grades_kz = adapt_grades_for_lang(student_data["grades"], "kz")
            grades_ru = adapt_grades_for_lang(student_data["grades"], "ru")

            data_kz = {
                "full_name":      name,
                "grades":         grades_kz,
                "diploma_id":     student_data["diploma_number_clean"],
                "college_name":   "Жамбыл инновациялық жоғары колледжінде",
                "specialization": "06130100 - Бағдарламалық қамтамасыз ету",
                "qualification":  "4S06130105 - Ақпараттық жүйелер технигі",
                "start_year":     "2023",
                "end_year":       "2026",
            }

            data_ru = {
                "full_name_ru":      name,
                "grades":            grades_ru,
                "diploma_id":        student_data["diploma_number_clean"],
                "college_name_ru":   "Жамбылском инновационным высшем колледже",
                "specialization_ru": "06130100 - Прграммное обеспечение",
                "qualification_ru":  "4S06130105 - Техник информационных систем",
                "start_year":        "2023",
                "end_year":          "2026",
            }

            try:
                path_kz = os.path.join(out_dir, f"{sheet_label}_{sanitized_name}_KZ.xlsx")
                generate_kz(data_kz, path_kz)

                path_ru = os.path.join(out_dir, f"{sheet_label}_{sanitized_name}_RU.xlsx")
                generate_ru(data_ru, path_ru)

            except Exception as e:
                print(f"  ERROR generating for {name}: {e}")
                import traceback; traceback.print_exc()
                continue

        print(f"  Sheet done — {sheet_processed} students.")

    print(f"\n{'='*60}")
    print(f"All done. Total diplomas generated: {total_processed} students.")
    print(f"Output directory: {os.path.abspath(out_dir)}")


if __name__ == "__main__":
    main()
