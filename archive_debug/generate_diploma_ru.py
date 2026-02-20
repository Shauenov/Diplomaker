"""
Diploma Supplement Generator (Russian Version)
==============================================
Generates an Excel file (diploma_ru.xlsx) that is a visual replica
of the official diploma supplement document (Russian).

Uses xlsxwriter for precise formatting control.
Subject names are HARDCODED (Russian).
"""

import xlsxwriter
import os

# ─────────────────────────────────────────────────────────────
# 1. SUBJECT NAMES (Russian)
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
    "География",
    "Биология",
    "Физика",
    "Графика и проектирование",
    "БМ 01 Развитие и совершенств. физических качеств",
    "БМ 02 Применение информационно-коммуникационных и цифровых технологий",
    "БМ 03 Применение базовых знаний экономики и основ предпринимательства",
    "БМ 04 Применение основ социальных наук для социализации и адаптации в обществе и трудовом коллективе",
]

PAGE2_SUBJECTS = [
    "ПМ 1 Понимание целей и видов бизнеса, а также взаимодействия с основными заинтересованными сторонами",
    "РО 1.1 Понимание целей и видов бизнеса, их взаимодействия с основными заинтересованными сторонами и внешней средой",
    "РО 1.2 Знание показательных и логарифмических функций, систем линейных уравнений и матриц, линейных неравенств и линейного программирования, основ теории вероятностей; применение этих понятий для анализа и интерпретации информации в бизнесе и финансовых расчетах",
    "РО 1.3 Понимание сути и назначения финансовой отчетности, определение качественных характеристик финансовой информации, подготовка финансовой отчетности",
    "РО 1.4 Понимание основных маркетинговых концепций, исследование маркетинговой среды, изучение поведения потребителей и организаций, сегментация рынков, размещение товаров и разработка новых продуктов, знание инструментов и методов, используемых в этих процессах",
    "ПМ 2 Использование языковых навыков в профессиональной сфере",
    "РО 2.1 Свободное владение навыками чтения, говорения и письма на английском языке на академическом уровне",
    "РО 2.2 Свободное владение навыками говорения и письма на английском языке в профессиональной сфере на уровне B2",
    "РО 2.3 Использование казахского языка в деловых целях",
    "РО 2.4 Использование турецкого языка в деловых целях",
    "ПМ 3 Участие в составлении бухгалтерской (финансовой) отчетности",
    "РО 3.1 Понимание характеристик и целей управленческой информации, учет затрат, планирование и контроль эффективности бизнеса",
]

PAGE3_SUBJECTS = [
    "РО 3.2 Понимание законодательства о трудовых отношениях, принципов управления и регулирования деятельности компаний",
    "РО 3.3 Использование математических инструментов, поддерживающих процесс принятия деловых решений, применение аналитических методов в различных бизнес-приложениях",
    "РО 3.4 Знание основных экономических принципов, макроэкономических проблем и показателей, расчет фискальных и кредитно-денежных политик, анализ механизмов их влияния на макроэкономику",
    "ПМ 4 Участие в комплексном анализе хозяйственно-финансовой деятельности организации и ее подразделений",
    "РО 4.1 Сравнение альтернативных методов оценки инвестиций и финансирования, оценка целесообразности различных способов решения проблем в финансовой сфере",
    "РО 4.2 Определение информации и технологических систем, необходимых для управления продуктивностью организаций, учет затрат и применение методов управленческого учета",
    "РО 4.3 Понимание функционирования и структуры налоговой системы, а также принципов ее управления",
    "РО 4.4 Учет операций в соответствии со стандартами МСФО, анализ и интерпретация финансовой отчетности",
    "РО 4.5 Знание основных понятий бизнес-статистики, сбор, обобщение и методы анализа данных",
    "РО 4.6 Информационные системы бухгалтерского учета",
    "РО 4.7 Понимание концепции аудита, его функций, корпоративного управления, включая вопросы этики и профессионального поведения, применение Международных стандартов аудита (МСА)",
    "ПМ 5 Оценка влияния экономической среды на финансовый менеджмент",
]

PAGE4_SUBJECTS = [
    "РО 5.1 Понимание роли и целей функций финансового управления, анализ влияния экономической среды на финансовый менеджмент",
    "РО 5.2 Эффективная оценка инвестиций, определение и анализ альтернативных источников финансирования бизнеса",
    "Профессиональная практика ПМ1 РО 1.3; ПМ3 РО3.1, РО3.2, РО3.3, РО3.4; ПМ4 РО 4.1, РО 4.2, РО 4.3, РО 4.4, РО 4.5, РО 4.6, РО 4.7; ПМ5 РО 5.1, РО 5.2.",
    "Итоговая аттестация:",
    "Факультативный английский язык",
    "Факультативный турецкий язык",
    "Факультативный курс «Кейсы в бизнесе и бухгалтерском учете» (Cases in Business and Accounting)",
    "Факультативный курс «Анализ бизнес-данных» (Business data analysis (Excel, Macros, Google Sheets, SQL, Python, Power BI, Tableau))",
    "Факультативный курс «Основы предпринимательской деятельности» (Entrepreneurship)",
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
        "font_size": 9,
        "align": "center",
        "valign": "top",
        "border": 0 # Explicit no border to be safe, though default is none
    })

    styles["subject"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 9,
        "align": "left",
        "valign": "top",
        "text_wrap": True,
    })

    styles["subject_bold"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 9,
        "bold": True,
        "align": "left",
        "valign": "top",
        "text_wrap": True,
    })

    styles["grade_center"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 9,
        "align": "center",
        "valign": "top",
    })

    styles["grade_left"] = workbook.add_format({
        "font_name": "Times New Roman",
        "font_size": 9,
        "align": "left",
        "valign": "top",
    })

    return styles


# ─────────────────────────────────────────────────────────────
# 3. COLUMN LAYOUT CONSTANTS
# ─────────────────────────────────────────────────────────────

COL_INDEX   = 0   # A — №
COL_SUBJECT = 1   # B — Наименование
COL_HOURS   = 2   # C — Часы
COL_CREDITS = 3   # D — Кредит
COL_POINTS  = 4   # E — Балл
COL_LETTER  = 5   # F — Әріп
COL_GPA     = 6   # G — GPA
COL_TRAD    = 7   # H — Трад

# Approximate character width for subject column to calculate row height
SUBJECT_COL_CHAR_WIDTH = 42


# ─────────────────────────────────────────────────────────────
# 4. HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────

def calc_row_height(text, col_chars=SUBJECT_COL_CHAR_WIDTH, line_height=13):
    """Calculate the minimum row height needed to fit the text."""
    if not text:
        return None
    # Adjust for Russian/Cyrillic slightly? Standard is fine.
    num_lines = max(1, -(-len(text) // col_chars))
    if num_lines <= 1:
        return None
    return num_lines * line_height


def is_module_header(subject_name):
    """Check if a subject row is a module header (ПМ, Practice, Final) — displayed bold."""
    # Russian Headers: "ПМ", "Профессиональная практика", "Итоговая аттестация"
    return (subject_name.startswith("ПМ ") or 
            subject_name.startswith("Профессиональная практика") or 
            subject_name.startswith("Итоговая аттестация"))


def setup_sheet(worksheet):
    """Apply common page/column settings to a worksheet."""
    worksheet.hide_gridlines(2)
    worksheet.set_portrait()
    worksheet.set_paper(9)  # A4
    worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)

    # Column widths
    worksheet.set_column(COL_INDEX,   COL_INDEX,   3)    # A
    worksheet.set_column(COL_SUBJECT, COL_SUBJECT, 32)   # B
    worksheet.set_column(COL_HOURS,   COL_HOURS,   5)    # C
    worksheet.set_column(COL_CREDITS, COL_CREDITS, 4)    # D
    worksheet.set_column(COL_POINTS,  COL_POINTS,  5)    # E
    worksheet.set_column(COL_LETTER,  COL_LETTER,  4)    # F
    worksheet.set_column(COL_GPA,     COL_GPA,     5)    # G
    worksheet.set_column(COL_TRAD,    COL_TRAD,    13)   # H


# ─────────────────────────────────────────────────────────────
# 5. PAGE BUILDERS
# ─────────────────────────────────────────────────────────────

def write_header(worksheet, styles, data):
    """Write heder with dynamic data (Russian)."""
    row = 1
    # {diploma_id}
    worksheet.merge_range(row, 0, row, 7, data.get("diploma_id", ""), styles["title_bold"])
    row += 2
    # {full_name}
    worksheet.merge_range(row, 0, row, 7, data.get("full_name_ru", data.get("full_name", "")), styles["title_bold"])
    row += 2
    # {start_year} ... {end_year}
    worksheet.merge_range(row, 0, row, 3, data.get("start_year", ""), styles["year_bold"])
    worksheet.merge_range(row, 4, row, 7, data.get("end_year", ""), styles["year_bold"])
    row += 1
    # {college_name} - Russian
    # User's image shows "Жамбылском инновационном высшем колледже"
    # Assuming the calling script might pass "college_name_ru".
    college = data.get("college_name_ru", "Жамбылском инновационным высшем колледже")
    worksheet.merge_range(row, 0, row, 7, college, styles["header_center"])
    row += 2
    
    # {specialization}
    # "04110100 – Учет и аудит"
    spec = data.get("specialization_ru", "04110100 – Учет и аудит")
    worksheet.merge_range(row, 0, row, 7, spec, styles["header_center"])
    row += 3
    
    # {qualification}
    # "4S04110102 - Бухгалтер"
    qual = data.get("qualification_ru", "4S04110102 - Бухгалтер")
    worksheet.merge_range(row, 0, row, 7, qual, styles["header_center"])

    return row + 1


def normalize_key(text):
    """Normalize subject name for looser matching (lowercase, no spaces)."""
    return text.lower().replace(" ", "").strip()


def write_subjects(worksheet, styles, start_row, subjects, start_index, grades_data):
    """Write subjects and fill grades if available."""
    current_row = start_row
    item_num = start_index

    # Pre-compute normalized keys for grades
    grades_map = {normalize_key(k): v for k, v in grades_data.items()}

    for subject in subjects:
        is_header = is_module_header(subject)

        # Set row height
        height = calc_row_height(subject)
        if height is not None:
            worksheet.set_row(current_row, height)

        # Write index
        worksheet.write(current_row, COL_INDEX, item_num, styles["index"])

        # Write subject
        fmt = styles["subject_bold"] if is_header else styles["subject"]
        worksheet.write(current_row, COL_SUBJECT, subject, fmt)

        # Retrieve grade data
        grade = grades_data.get(subject)
        if not grade:
            grade = grades_map.get(normalize_key(subject), {})

        # Logic for writing grades
        if is_header:
            # Header: Write ONLY Hours and Credits
            worksheet.write(current_row, COL_HOURS,   grade.get("hours", ""),   styles["grade_center"])
            worksheet.write(current_row, COL_CREDITS, grade.get("credits", ""), styles["grade_center"])
            
        else:
            # Regular Subject
            
            # Special Case: Electives (Факультатив) -> Traditional = "зачтено" (for Russian)
            # User request for KZ was "сынақ". For RU it is "зачтено".
            trad_val = grade.get("traditional_ru", grade.get("traditional", ""))
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
    return {
        "diploma_id": "",
        "full_name_ru": "",
        "start_year": "",
        "end_year": "",
        "college_name_ru": "",
        "grades": {}
    }


def generate_diploma_ru(data=None, output_path="diploma_ru_v1.xlsx"):
    """Generate an Excel diploma supplement file (Russian)."""
    if data is None:
        data = get_empty_data()

    workbook = xlsxwriter.Workbook(output_path)
    styles = create_formats(workbook)

    # ═══════════════════ LISТ 1 ═══════════════════
    ws1 = workbook.add_worksheet("Лист 1")
    setup_sheet(ws1)

    # Write header
    header_end_row = write_header(ws1, styles, data)

    # 6 empty rows gap
    data_start = header_end_row + 6

    # Write Page 1 subjects
    write_subjects(ws1, styles, data_start, PAGE1_SUBJECTS, start_index=1, grades_data=data.get("grades", {}))

    # ═══════════════════ LISТ 2 ═══════════════════
    ws2 = workbook.add_worksheet("Лист 2")
    setup_sheet(ws2)
    write_subjects(ws2, styles, start_row=1, subjects=PAGE2_SUBJECTS, start_index=18, grades_data=data.get("grades", {}))

    # ═══════════════════ LISТ 3 ═══════════════════
    ws3 = workbook.add_worksheet("Лист 3")
    setup_sheet(ws3)
    write_subjects(ws3, styles, start_row=1, subjects=PAGE3_SUBJECTS, start_index=30, grades_data=data.get("grades", {}))

    # ═══════════════════ LISТ 4 ═══════════════════
    ws4 = workbook.add_worksheet("Лист 4")
    setup_sheet(ws4)
    write_subjects(ws4, styles, start_row=1, subjects=PAGE4_SUBJECTS, start_index=42, grades_data=data.get("grades", {}))
    
    workbook.close()
    print(f"✅ Russian Diploma generated: {os.path.abspath(output_path)}")
    return output_path


if __name__ == "__main__":
    generate_diploma_ru(output_path="diploma_ru_template.xlsx")
