"""
Diploma Supplement Generator (Russian IT Version)
==============================================
Generates an Excel file (Diplom_IT_RU_Template.xlsx) that is a visual replica
of the official diploma supplement document (Russian) for IT students.

Based on generate_diploma_it_kz.py logic/layout.
"""

import xlsxwriter
import os
import re


# ─────────────────────────────────────────────────────────────
# 1. SUBJECT NAMES (Russian IT)
# ─────────────────────────────────────────────────────────────

PAGE1_SUBJECTS = [
    "Казахский язык",
    "Казахская литература",
    "Русский язык и литература",
    "Английский язык",
    "История Казахстана",
    "Математика",
    "Информатика",
    "Начальная военная и технологическая подготовка",
    "Физическая культура",
    "Физика",
    "Химия",
    "Биология",
    "Всемирная история",
]

PAGE2_SUBJECTS = [
    "БМ 1. Развитие и совершенствование физических качеств",
    "БМ 2. Применение информационно-коммуникационных и цифровых технологий",
    "БМ 3. Применение базовых знаний экономики и основ предпринимательства",
    "БМ 4. Применение основ социальных наук для социализации и адаптации в обществе и трудовом коллективе",
    "ПМ 01 Использование устных и письменных навыков общения в профессиональной сфере",
    "РО 1.1 Освоение лексико-грамматического материала английского языка для профессионального общения",
    "РО 1.2 Чтение и перевод профессионально ориентированных текстов",
    "РО 1.3 Формирование ключевых речевых форм для выражения мысли",
    "РО 1.4 Использование второго иностранного языка в деловом общении",
    "ПМ 02 Применение математических и статистических расчетов в профессиональной деятельности",
    "РО 2.1 Применение основных понятий и методов алгебры последовательностей",
    "РО 2.2 Использование базовых знаний аналитической геометрии",
    "РО 2.3 Освоение основ математического анализа",
    "ПМ 03 Разработка web-ресурсов Front-end",
    "РО 3.1 Создание шаблонов и плагинов для систем управления контентом",
    "РО 3.2 Изменение дизайна веб-сайтов с использованием CSS и других внешних файлов",
    "РО 3.3 Разработка, обновление и настройка веб-сайтов для пользователей",
    "ПМ 04 Графический дизайн",
    "РО 4.1 Проектирование промышленных изделий с учетом комплексного дизайна",
    "РО 4.2 Создание макетов дизайна",
    "РО 4.3 Разработка растровых и векторных изображений на компьютере",
]

PAGE3_SUBJECTS = [
    "ПМ 05 Программирование и разработка алгоритмов",
    "РО 5.1 Разработка клиент-серверных программных решений.",
    "РО 5.2 Использование новейших сред разработки программного обеспечения и инструментов для модификации кодов",
    "РО 5.3 Применение методов развития программных систем",
    "ПМ 06 Стандарты управления ИТ, меры по охране окружающей среды и безопасности, управление проектом",
    "РО 6.1 Интервью с пользователями, опросы, поиск и анализ документов, совместная разработка программ и отслеживание обработки требований.",
    "РО 6.2 Разработка альтернатив для принятия решений, выбор наиболее подходящей альтернативы и принятие необходимого решения",
    "РО 6.3 Реализация целей и задач программного обеспечения с учетом требований производства",
    "РО 6.4 Разрабатывать ИТ-решения для компонентов программного обеспечения в производстве с учетом технических, экологических и мер безопасности.",
    "ПМ 07 Разработка web-ресурсов Back-end",
    "РО 7.1 Управление связью между серверными и клиентскими системами",
    "РО 7.2 Использование технологий работы с базами данных",
    "РО 7.3 Применение шаблонов проектирования программного обеспечения",
    "ПМ 08 Визуальный дизайн UX/UI",
    "РО 8.1 Разработка мультимедийных компонентов и веб-элементов",
    "РО 8.2 Создание графических материалов для пользовательского интерфейса",
    "РО 8.3 Создание визуального дизайна элементов графического пользовательского интерфейса.",
    "ПМ 09 Разработка мобильных приложений",
    "РО 9.1 Использование систем управления базами данных для хранения и обработки данных",
    "РО 9.2 Создание уровневых устройств.",
    "РО 9.3 Создание мобильных приложений на основе клиент-серверных решений",
]

PAGE4_SUBJECTS = [
    "ПМ 10 Реализация проектной деятельности",
    "РО 10.1 Развитие практических навыков проектирования",
    "РО 10.2 Применение современных методов проектирования и расчетов",
    "РО 10.3 Разработка технических условий и предложений по дипломному проекту",
    "РО 10.4 Разработка программного продукта по теме дипломного проекта",
    "Профессиональная практика ПМ3. РО3.2, РО3.3; ПМ4. РО4.3; ПМ5. РО5.2, РО5.3; ПМ7. РО7.1, РО7.2, РО7.3; ПМ8. РО8.1, РО8.2, РО8.3; ПМ9. РО9.1, РО9.2, РО9.3.",
    "Итоговая аттестация:",
    "Факультатив английский язык",
    "Факультатив турецкий язык",
    "Факультатив основы предпринимательской деятельности",
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
        "font_size": 8,  # Same as KZ
        "align": "center",
        "valign": "top",
    })

    styles["subject"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 8,  # Same as KZ
        "align": "left",
        "valign": "top",
        "text_wrap": True,
    })

    styles["subject_bold"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 8,  # Same as KZ
        "bold": True,
        "align": "left",
        "valign": "top",
        "text_wrap": True,
    })

    styles["grade_center"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 8,  # Same as KZ
        "align": "center",
        "valign": "top",
    })

    styles["grade_left"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 8,  # Same as KZ
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
SUBJECT_COL_CHAR_WIDTH = 45 


# ─────────────────────────────────────────────────────────────
# 4. HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────

def calc_row_height(text, col_chars=SUBJECT_COL_CHAR_WIDTH, line_height=11):
    """Calculate the minimum row height needed to fit the text."""
    if not text:
        return None  # let Excel use default height
    
    lines = text.split('\n')
    total_lines = 0
    for line in lines:
        total_lines += max(1, -(-len(line) // col_chars))
        
    if total_lines <= 1:
        return 12 # Force smaller single line height
    return total_lines * line_height


def is_module_header(subject_name):
    """Check if a subject row is a module header (PM, Practice, Final).
    БМ subjects are regular subjects — NOT headers.
    """
    s = subject_name.strip()
    return (s.startswith("ПМ") or 
            s.startswith("Профессиональная практика") or 
            s.startswith("Итоговая аттестация"))


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
    """Write the header section on a worksheet (Russian metadata)."""
    row = 1

    # {diploma_id}
    worksheet.merge_range(row, 0, row, 7, data.get("diploma_id", ""), styles["title_bold"])
    row += 2

    # {full_name_ru}
    worksheet.merge_range(row, 0, row, 7, data.get("full_name_ru", ""), styles["title_bold"])
    row += 2

    # {start_year}  ...  {end_year}
    worksheet.merge_range(row, 0, row, 3, data.get("start_year", ""), styles["year_bold"])
    worksheet.merge_range(row, 4, row, 7, data.get("end_year", ""), styles["year_bold"])
    row += 1

    # {college_name_ru}
    college = data.get("college_name_ru", "")
    worksheet.merge_range(row, 0, row, 7, college, styles["header_center"])
    row += 2

    # {specialization_ru}
    spec = data.get("specialization_ru", "")
    worksheet.merge_range(row, 0, row, 7, spec, styles["header_center"])
    row += 3

    # {qualification_ru}
    qual = data.get("qualification_ru", "")
    worksheet.merge_range(row, 0, row, 7, qual, styles["header_center"])

    # Return the row number after header (+ 6 empty rows gap before data)
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
                # ПМ with sub-subjects: write summed values only
                worksheet.write(current_row, COL_HOURS,   _fmt(total_hours),   styles["grade_center"])
                worksheet.write(current_row, COL_CREDITS, _fmt(total_credits), styles["grade_center"])
            else:
                # Terminal module (БМ, practice, attestation): direct lookup for hours/credits
                g = grades_data.get(subject) or grades_map.get(normalize_key(subject), {})
                if not g:
                    if "профессиональная практика" in subject.lower():
                        for k, v in grades_map.items():
                            if "профессиональнаяпрактика" in k:
                                g = v; break
                    elif "аттестация" in subject.lower():
                        for k, v in grades_map.items():
                            if "аттестация" in k:
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
                import re as _re
                bm_prefix_match = _re.match(r'(БМ\s*\.?\s*\d+)', subject)
                if bm_prefix_match:
                    bm_prefix = normalize_key(bm_prefix_match.group(1))
                    for k, v in grades_map.items():
                        if k.startswith(bm_prefix):
                            grade = v
                            break

            # Special logic for Electives: "зачтено"
            trad_val = grade.get("traditional", "")
            if subject.startswith("Факультатив"):
                trad_val = "зачтено"

            worksheet.write(current_row, COL_HOURS,   grade.get("hours", ""),   styles["grade_center"])
            worksheet.write(current_row, COL_CREDITS, grade.get("credits", ""), styles["grade_center"])
            worksheet.write(current_row, COL_POINTS,  grade.get("points", ""),  styles["grade_center"])
            worksheet.write(current_row, COL_LETTER,  grade.get("letter", ""),  styles["grade_center"])
            worksheet.write(current_row, COL_GPA,     grade.get("gpa", ""),     styles["grade_center"])
            worksheet.write(current_row, COL_TRAD,    trad_val,                 styles["grade_left"])

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
        "full_name_ru": "",
        "start_year": "",
        "end_year": "",
        "college_name_ru": "",
        "specialization_ru": "",
        "qualification_ru": "",
        "grades": {} 
    }


def generate_diploma(data=None, output_path="Diplom_IT_RU_Template.xlsx"):
    """Generate an Excel diploma supplement file."""
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

    # ═══════════════════ LISТ 1 ═══════════════════
    ws1 = workbook.add_worksheet("Лист 1")
    setup_sheet(ws1)

    # Write header with data
    header_end_row = write_header(ws1, styles, data)

    # 6 empty rows gap before subjects
    data_start = header_end_row + 6

    # Write Page 1 subjects
    idx1 = 1
    cnt1 = write_subjects(ws1, styles, data_start, pages_subjects[0], start_index=idx1, grades_data=data.get("grades", {}))

    # ═══════════════════ LISТ 2 ═══════════════════
    ws2 = workbook.add_worksheet("Лист 2")
    setup_sheet(ws2)
    
    idx2 = idx1 + len(pages_subjects[0])
    cnt2 = write_subjects(ws2, styles, start_row=1, subjects=pages_subjects[1], start_index=idx2, grades_data=data.get("grades", {}))

    # ═══════════════════ LISТ 3 ═══════════════════
    ws3 = workbook.add_worksheet("Лист 3")
    setup_sheet(ws3)
    
    idx3 = idx2 + len(pages_subjects[1])
    cnt3 = write_subjects(ws3, styles, start_row=1, subjects=pages_subjects[2], start_index=idx3, grades_data=data.get("grades", {}))

    # ═══════════════════ LISТ 4 ═══════════════════
    ws4 = workbook.add_worksheet("Лист 4")
    setup_sheet(ws4)
    
    idx4 = idx3 + len(pages_subjects[2])
    cnt4 = write_subjects(ws4, styles, start_row=1, subjects=pages_subjects[3], start_index=idx4, grades_data=data.get("grades", {}))
    
    workbook.close()
    print(f"✅ Russian Status Diploma generated: {os.path.abspath(output_path)}")
    return output_path


if __name__ == "__main__":
    generate_diploma()
