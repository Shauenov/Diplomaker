"""
Diploma Supplement Generator (Kazakh Version)
==============================================
Generates an Excel file (diploma.xlsx) that is a visual replica
of the official diploma supplement document.

Uses xlsxwriter for precise formatting control.

Subject names are HARDCODED (they don't change between students).
Grade cells are LEFT EMPTY — fill them in later via automation or manually.

Layout: Page 1 = Header + Subjects 1–17
        Page 2 = Subjects 18–29

Usage:
    python generate_diploma.py
"""

import xlsxwriter
import os
import re


# ─────────────────────────────────────────────────────────────
# 1. SUBJECT NAMES (Hardcoded — same for all students)
# ─────────────────────────────────────────────────────────────

PAGE1_SUBJECTS = [
    "Қазақ тілі",
    "Қазақ әдебиеті",
    "Орыс тілі және әдебиеті",
    "Ағылшын тілі",
    "Қазақстан тарихы",
    "Математика",
    "Информатика",
    "Алғашқы әскери және технологиялық дайындық",
    "Дене тәрбиесі",
    "Физика",
    "Химия",
    "Биология",
    "Дүниежүзі тарихы",
]

PAGE2_SUBJECTS = [
    "БМ 1 Дене қасиеттерін дамыту және жетілдіру",
    "БМ 2 Ақпараттық-коммуникациялық және цифрлық технологияларды қолдану",
    "БМ 3 Экономиканың базалық білімін және кәсіпкерлік негіздерін қолдану",
    "БМ 4 Қоғам мен еңбек ұжымында әлеуметтену және бейімделу үшін әлеуметтік ғылымдар негіздерін қолдану",
    "КМ 01 Кәсіптік қызмет саласында коммуникация үшін ауызша және жазбаша қарым-қатынас дағдыларын пайдалану",
    "ОН 1.1 Кәсіптік салада қарым-қатынас жасауға қажетті ағылшын тілінің лексика-грамматикалық материалдарын меңгеру",
    "ОН 1.2 Кәсіптік бағыттағы мәтіндерді оқу және аудару",
    "ОН 1.3 Ойды жеткізудегі негізгі сөйлеу формаларын қалыптастыру",
    "ОН 1.4 Іскерлік мақсатта екінші шетел тілін қолдану",
    "КМ 02 Кәсіби қызметте математикалық, статистикалық есептерді қолдану",
    "ОН 2.1 Тізбекті алгебра негізгі түсінігі мен әдістерін қолдану",
    "ОН 2.2 Аналитикалық геометрия негізгі түсінігі мен әдістерін қолдану",
    "ОН 2.3 Математикалық талдаудың негіздерін қалыптастыру",
    "КМ 03 Front-end web ресурстарды құру",
    "ОН 3.1 Контентті басқару жүйесі үшін жеке шаблондар мен плагиндер жасау",
    "ОН 3.2 Веб-сайттың көрінісін өзгерту үшін CSS немесе басқа сыртқы файлдарды пайдалану",
    "ОН 3.3 Пайдаланушылар үшін web-сайттарды құру, жаңарту және іздеу жүйесін құрастыру",
    "КМ 04 Графикалық дизайн жасау",
    "ОН 4.1 Өнеркәсіптік дизайнның жан-жағынан жинақталған бұйымдардың жобасын жасау",
    "ОН 4.2 Дизайн макеттерін дайындау",
    "ОН 4.3 Компьютерде растрлық және векторлық бейнелерді құрастыру",
]

PAGE3_SUBJECTS = [
    "КМ 05 Алгоритмге кіріспе түрлерін пайдаланып, бағдарламалар жасау",
    "ОН 5.1 Клиент-сервер негізінде бағдарламалық шешімдердің кодтарын құрастыру",
    "ОН 5.2 Кодтарды өзгерту үшін соңғы бағдарламалық жасақтаманы әзірлеу орталары мен құралдарын пайдалану",
    "ОН 5.3 Жүйені дамыту әдістемелерін пайдалану",
    "КМ 06 IT Менеджмент стандарттарымен экологиялық және қауіпсіздік шаралармен жобаларды басқару",
    "ОН 6.1 Пайдаланушы сұхбат, сауалнама, құжаттарды іздеу және талдау, бірлескен бағдарламаны әзірлеу және бақылау талаптарын өңдеу",
    "ОН 6.2 Шешім қабылдау үшін баламаларды әзірлеу, ең қолайлы баламаны таңдау және қажетті шешімді жасау",
    "ОН 6.3 Бағдарламалық қамтамасыздандырудың мақсат пен міндет қоюды жүзеге асыру және қойылатын талаптарды әзірлеу",
    "ОН 6.4 Өндірістегі бағдарламалық компоненттерге техникалық, экологиялық және қауіпсіздік шараларын ескере отырып IT шешімдерін әзірлеу",
    "КМ 07 Back-end web ресурстарын құру",
    "ОН 7.1 Сервер мен клиенттік жүйелер арасындағы байланысты басқару",
    "ОН 7.2 Деректер базасымен жұмыс істеу технологиясын пайдалану",
    "ОН 7.3 Бағдарламалық жасақтаманың дизайн үлгілерін пайданалу",
    "КМ 08 UX/UI визуалды дизайн жасау",
    "ОН 8.1 Мультимедиялық қосымшаларды, web-элементтерді құрастыру",
    "ОН 8.2 Графикалық интерфейске қосу үшін графикалық материалдарды дайындау",
    "ОН 8.3 Графикалық пайдаланушы интерфейс элементтерінің визуалды дизайнын жасау",
    "КМ 09 Мобильді қосымшаларды әзірлеу",
    "ОН 9.1 Деректерді жасау, сақтау және басқару үшін дерекқорды басқару жүйесін пайдалану",
    "ОН 9.2 Деңгейлі құрылғыларды жасау",
    "ОН 9.3 Клиент-сервер негізіндегі жүйе үшін мобильді интерфейсті құрастыру",
]

PAGE4_SUBJECTS = [
    "КМ 10 Жобалық іс-шараларды орындау",
    "ОН 10.1 Өз бетінше жобалау жұмыстарына практикалық дағдыларды қалыптастыру",
    "ОН 10.2 Тәжірибе негізінде қолданылатын заманауи жобалау және есептеу әдістерін қолдану",
    "ОН 10.3 Дипломдық жоба тақырыбы бойынша техникалық шарттар мен техникалық ұсыныстар әзірлеу",
    "ОН 10.4 Дипломдық жоба тақырыбы бойынша бағдарламалық өнімді әзірлеу",
    "Кәсіптік практика КМ3. ОН3.2, ОН3.3; КМ4. ОН4.3; КМ5. ОН5.2, ОН5.3; КМ7. ОН7.1, ОН7.2, ОН7.3; КМ8. ОН8.1, ОН8.2, ОН8.3; КМ9. ОН9.1, ОН9.2, ОН9.3.",
    "Қорытынды аттестаттау :",
    "Ф1 Факультативтік ағылшын тілі",
    "Ф2 Факультативтік түрік тілі",
    "Ф3 Факультативтік кәсіпкерлік қызмет негіздері",
]


# ─────────────────────────────────────────────────────────────
# 2. STYLES
# ─────────────────────────────────────────────────────────────

def create_formats(workbook):
    """Create all reusable cell formats."""
    styles = {}

    # --- Header formats ---
    styles["title_bold"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 13,
        "bold": True,
        "align": "center",
        "valign": "vcenter",
    })

    styles["header_center"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 11,
        "align": "center",
        "valign": "vcenter",
    })

    styles["year_bold"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 13,
        "bold": True,
        "align": "center",
        "valign": "vcenter",
    })

    # --- Data table formats ---
    styles["index"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 8,  # Reduced from 9
        "align": "center",
        "valign": "top",
    })

    styles["subject"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 8,  # Reduced from 9
        "align": "left",
        "valign": "top",
        "text_wrap": True,
    })

    styles["subject_bold"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 8,  # Reduced from 9
        "bold": True,
        "align": "left",
        "valign": "top",
        "text_wrap": True,
    })

    styles["grade_center"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 8,  # Reduced from 9
        "align": "center",
        "valign": "top",
    })

    styles["grade_left"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 8,  # Reduced from 9
        "align": "left",
        "valign": "top",
    })

    return styles


# ─────────────────────────────────────────────────────────────
# 3. COLUMN LAYOUT CONSTANTS
# ─────────────────────────────────────────────────────────────

COL_INDEX   = 0   # A — №
COL_SUBJECT = 1   # B — Пән атауы
COL_HOURS   = 2   # C — Сағат
COL_CREDITS = 3   # D — Кредит
COL_POINTS  = 4   # E — Балл
COL_LETTER  = 5   # F — Әріп
COL_GPA     = 6   # G — GPA
COL_TRAD    = 7   # H — Дәстүрлі баға

# Approximate character width for subject column to calculate row height
# Increased slightly since font is smaller, but we want tighter packing
SUBJECT_COL_CHAR_WIDTH = 45 


# ─────────────────────────────────────────────────────────────
# 4. HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────

def calc_row_height(text, col_chars=SUBJECT_COL_CHAR_WIDTH, line_height=11):
    """Calculate the minimum row height needed to fit the text.

    Only increases height for multi-line subjects.
    Single-line subjects keep the default Excel row height (no stretching).
    """
    if not text:
        return None  # let Excel use default height
    
    # Heuristic: If it has newlines or is very long
    lines = text.split('\n')
    total_lines = 0
    for line in lines:
        total_lines += max(1, -(-len(line) // col_chars))
        
    if total_lines <= 1:
        return 12 # Force smaller single line height
    return total_lines * line_height


def is_module_header(subject_name):
    """Check if a subject row is a module header.
    БМ subjects are regular subjects — NOT headers.
    """
    s = subject_name.strip()
    return (s.startswith("КМ") or 
            s.startswith("Кәсіптік практика") or 
            s.startswith("Қорытынды аттестаттау"))


def setup_sheet(worksheet):
    """Apply common page/column settings to a worksheet."""
    worksheet.hide_gridlines(2)
    worksheet.set_portrait()
    worksheet.set_paper(9)  # A4
    worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)

    # Column widths
    worksheet.set_column(COL_INDEX,   COL_INDEX,   4)    # A — №
    worksheet.set_column(COL_SUBJECT, COL_SUBJECT, 30)   # B — Subject (Reduced from 35)
    worksheet.set_column(COL_HOURS,   COL_HOURS,   5)    # C — Hours
    worksheet.set_column(COL_CREDITS, COL_CREDITS, 4)    # D — Credits
    worksheet.set_column(COL_POINTS,  COL_POINTS,  5)    # E — Points
    worksheet.set_column(COL_LETTER,  COL_LETTER,  4)    # F — Letter
    worksheet.set_column(COL_GPA,     COL_GPA,     5)    # G — GPA
    worksheet.set_column(COL_TRAD,    COL_TRAD,    13)   # H — Traditional


# ─────────────────────────────────────────────────────────────
# 5. PAGE BUILDERS
# ─────────────────────────────────────────────────────────────

def write_header(worksheet, styles, data):
    """Write the header section on a worksheet.

    Placeholders are left EMPTY for future automation:
        Row 1: {diploma_id}
        Row 3: {full_name}
        Row 5: {start_year}  ...  {end_year}
        Row 6: {college_name}
        Row 8: {specialization}
        Row 11: {qualification}
    """
    row = 1

    # {diploma_id}
    worksheet.merge_range(row, 0, row, 7, data.get("diploma_id", ""), styles["title_bold"])
    row += 2

    # {full_name}
    worksheet.merge_range(row, 0, row, 7, data.get("full_name", ""), styles["title_bold"])
    row += 2

    # {start_year}  ...  {end_year}
    worksheet.merge_range(row, 0, row, 3, data.get("start_year", ""), styles["year_bold"])
    worksheet.merge_range(row, 4, row, 7, data.get("end_year", ""), styles["year_bold"])
    row += 1

    # {college_name}
    worksheet.merge_range(row, 0, row, 7, data.get("college_name", ""), styles["header_center"])
    row += 2

    # {specialization}
    worksheet.merge_range(row, 0, row, 7, data.get("specialization", ""), styles["header_center"])
    row += 3

    # {qualification}
    worksheet.merge_range(row, 0, row, 7, data.get("qualification", ""), styles["header_center"])
    # Return the row number after header (+ 6 empty rows gap before data)
    return row + 1


def _write_row(worksheet, styles, row_idx, item_num, subject_name, grade_info):
    """Helper to write a single subject row."""
    is_header = is_module_header(subject_name)

    # Set row height only if text overflows
    height = calc_row_height(subject_name)
    if height is not None:
        worksheet.set_row(row_idx, height)

    # Write index
    worksheet.write(row_idx, COL_INDEX, item_num, styles["index"])

    # Write subject (bold for КМ headers)
    fmt = styles["subject_bold"] if is_header else styles["subject"]
    worksheet.write(row_idx, COL_SUBJECT, subject_name, fmt)

    # Grade cells (Cols C–H)
    if not is_header:
        worksheet.write(row_idx, COL_HOURS,   grade_info.get("hours", ""),   styles["grade_center"])
        worksheet.write(row_idx, COL_CREDITS, grade_info.get("credits", ""), styles["grade_center"])
        worksheet.write(row_idx, COL_POINTS,  grade_info.get("points", ""),  styles["grade_center"])
        worksheet.write(row_idx, COL_LETTER,  grade_info.get("letter", ""),  styles["grade_center"])
        worksheet.write(row_idx, COL_GPA,     grade_info.get("gpa", ""),     styles["grade_center"])
        worksheet.write(row_idx, COL_TRAD,    grade_info.get("traditional", ""), styles["grade_left"])


def write_subjects(worksheet, styles, start_row, subjects, start_index, grades_data):
    """Write a subject block with empty grade cells."""
    current_row = start_row
    item_num = start_index
    
    # Pre-compute normalized keys for grades
    grades_map = {normalize_key(k): v for k, v in grades_data.items()}

    for i, subject in enumerate(subjects):
        is_header = is_module_header(subject)

        # Set row height only if text overflows
        height = calc_row_height(subject)
        if height is not None:
            worksheet.set_row(current_row, height)

        # Write index
        worksheet.write(current_row, COL_INDEX, item_num, styles["index"])

        # Write subject (bold for KM headers)
        fmt = styles["subject_bold"] if is_header else styles["subject"]
        worksheet.write(current_row, COL_SUBJECT, subject, fmt)

        # Grade cells (Cols C–H)
        if is_header:
            # ── Sum hours/credits from all sub-subjects below this header ──
            total_hours = 0.0
            total_credits = 0.0

            for sub_idx in range(i + 1, len(subjects)):
                sub_name = subjects[sub_idx]
                if is_module_header(sub_name):
                    break
                g = grades_data.get(sub_name) or grades_map.get(normalize_key(sub_name), {})
                if g:
                    try:
                        h = float(str(g.get("hours") or "0").replace(",", "."))
                    except Exception:
                        h = 0
                    try:
                        c = float(str(g.get("credits") or "0").replace(",", "."))
                    except Exception:
                        c = 0
                    total_hours += h
                    total_credits += c

            def _fmt(v):
                return str(int(v)) if v == int(v) else str(v)

            if total_hours > 0:
                # КМ with sub-subjects: write summed values only
                worksheet.write(current_row, COL_HOURS,   _fmt(total_hours),   styles["grade_center"])
                worksheet.write(current_row, COL_CREDITS, _fmt(total_credits), styles["grade_center"])
            else:
                # Terminal module (БМ, practice, attestation): direct lookup for hours/credits
                g = grades_data.get(subject) or grades_map.get(normalize_key(subject), {})
                if not g:
                    if "кәсіптік практика" in subject.lower():
                        for k, v in grades_map.items():
                            if "кәсіптікпрактика" in k:
                                g = v; break
                    elif "аттестаттау" in subject.lower():
                        for k, v in grades_map.items():
                            if "аттестаттау" in k:
                                g = v; break
                if g:
                    worksheet.write(current_row, COL_HOURS,   g.get("hours",   ""), styles["grade_center"])
                    worksheet.write(current_row, COL_CREDITS, g.get("credits", ""), styles["grade_center"])
            # Module headers NEVER show points / letter / GPA / traditional

        else:
            # NORMAL SUBJECT
            grade = grades_data.get(subject)
            if not grade:
                norm_key = normalize_key(subject)
                grade = grades_map.get(norm_key, {})
            
            # Fallback for БМ subjects: match by "БМ N" prefix
            if not grade and subject.startswith('БМ'):
                bm_prefix_match = re.match(r'(БМ\s*\.?\s*\d+)', subject)
                if bm_prefix_match:
                    bm_prefix = normalize_key(bm_prefix_match.group(1))
                    for k, v in grades_map.items():
                        if k.startswith(bm_prefix):
                            grade = v
                            break

            worksheet.write(current_row, COL_HOURS,   grade.get("hours", ""),   styles["grade_center"])
            worksheet.write(current_row, COL_CREDITS, grade.get("credits", ""), styles["grade_center"])
            worksheet.write(current_row, COL_POINTS,  grade.get("points", ""),  styles["grade_center"])
            worksheet.write(current_row, COL_LETTER,  grade.get("letter", ""),  styles["grade_center"])
            worksheet.write(current_row, COL_GPA,     grade.get("gpa", ""),     styles["grade_center"])
            worksheet.write(current_row, COL_TRAD,    grade.get("traditional", ""), styles["grade_left"])

        item_num += 1
        current_row += 1

    return current_row


# ─────────────────────────────────────────────────────────────
# 6. MAIN GENERATION FUNCTION
# ─────────────────────────────────────────────────────────────

def get_empty_data():
    """Return an empty data structure for testing/template generation."""
    return {
        "diploma_id": "",
        "full_name": "",
        "start_year": "",
        "end_year": "",
        "college_name": "",
        "specialization": "",
        "qualification": "",
        "grades": {}  # Key: subject_name, Value: {hours, credits, points, letter, gpa, traditional}
    }


def generate_diploma(data=None, output_path="Diplom_IT_KZ_Template.xlsx"):
    """Generate an Excel diploma supplement file.
    
    Args:
        data: dict with student info and grades. If None, uses empty placeholders.
              Can explicitly provide "pages": [[subj_list_p1], [subj_list_p2], ...] 
              to override default hardcoded subjects.
        output_path: file path to save.
    """
    if data is None:
        data = get_empty_data()
        


    workbook = xlsxwriter.Workbook(output_path)
    styles = create_formats(workbook)
    
    # HARDCODED PAGINATION from Template
    pages_subjects = [
        PAGE1_SUBJECTS,
        PAGE2_SUBJECTS,
        PAGE3_SUBJECTS,
        PAGE4_SUBJECTS
    ]

    # ═══════════════════ PAGE 1 ═══════════════════
    ws1 = workbook.add_worksheet("Бет 1")
    setup_sheet(ws1)

    # Write header with data
    header_end_row = write_header(ws1, styles, data)

    # 6 empty rows gap before subjects
    data_start = header_end_row + 6

    # Write Page 1 subjects
    # Calculate start index for next pages
    idx1 = 1
    cnt1 = write_subjects(ws1, styles, data_start, pages_subjects[0], start_index=idx1, grades_data=data.get("grades", {}))
    
    # ═══════════════════ PAGE 2 ═══════════════════
    ws2 = workbook.add_worksheet("Бет 2")
    setup_sheet(ws2)
    
    idx2 = idx1 + len(pages_subjects[0])
    cnt2 = write_subjects(ws2, styles, start_row=1, subjects=pages_subjects[1], start_index=idx2, grades_data=data.get("grades", {}))

    # ═══════════════════ PAGE 3 ═══════════════════
    ws3 = workbook.add_worksheet("Бет 3")
    setup_sheet(ws3)
    
    idx3 = idx2 + len(pages_subjects[1])
    cnt3 = write_subjects(ws3, styles, start_row=1, subjects=pages_subjects[2], start_index=idx3, grades_data=data.get("grades", {}))

    # ═══════════════════ PAGE 4 ═══════════════════
    ws4 = workbook.add_worksheet("Бет 4")
    setup_sheet(ws4)
    
    idx4 = idx3 + len(pages_subjects[2])
    cnt4 = write_subjects(ws4, styles, start_row=1, subjects=pages_subjects[3], start_index=idx4, grades_data=data.get("grades", {}))
    
    workbook.close()
    print(f"✅ Diploma generated: {os.path.abspath(output_path)}")
    return output_path


def write_header(worksheet, styles, data):
    """Write heder with dynamic data."""
    row = 1
    # {diploma_id}
    worksheet.merge_range(row, 0, row, 7, data.get("diploma_id", ""), styles["title_bold"])
    row += 2
    # {full_name}
    worksheet.merge_range(row, 0, row, 7, data.get("full_name", ""), styles["title_bold"])
    row += 2
    # {start_year} ... {end_year}
    worksheet.merge_range(row, 0, row, 3, data.get("start_year", ""), styles["year_bold"])
    worksheet.merge_range(row, 4, row, 7, data.get("end_year", ""), styles["year_bold"])
    row += 1
    # {college_name}
    worksheet.merge_range(row, 0, row, 7, data.get("college_name", ""), styles["header_center"])
    row += 2
    # {specialization}
    worksheet.merge_range(row, 0, row, 7, data.get("specialization", ""), styles["header_center"])
    row += 3
    # {qualification}
    worksheet.merge_range(row, 0, row, 7, data.get("qualification", ""), styles["header_center"])
    return row + 1


def normalize_key(text):
    """Normalize subject name for looser matching."""
    if not text: return ""
    # Lowercase
    t = text.lower()
    # Remove dots, commas, colons
    t = t.replace(".", "").replace(",", "").replace(":", "")
    # Remove spaces
    t = t.replace(" ", "")
    # Remove leading zeros in numbers (e.g. pm01 -> pm1)
    t = re.sub(r'([a-zа-я]+)0+([1-9]+)', r'\1\2', t)
    return t.strip()


if __name__ == "__main__":
    # Example: Generate with empty data
    try:
        output_file = generate_diploma(output_path="Diplom_IT_KZ_Template.xlsx")
    except Exception as e:
        print(f"Error generating: {e}")

    # try:
    #     os.startfile(output_file)
    # except Exception:
    #     pass
