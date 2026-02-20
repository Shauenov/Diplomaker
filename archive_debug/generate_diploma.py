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
    "География",
    "Биология",
    "Физика",
    "Графика және жобалау",
    "БМ 01 Дене қасиеттерін дамыту және жетілдіру",
    "БМ 02 Ақпараттық-коммуникациялық және цифрлық технологияларды қолдану",
    "БМ 03 Экономиканың базалық білімін және кәсіпкерлік негіздерін қолдану",
    "БМ 04 Қоғам мен еңбек ұжымында әлеуметтену және бейімделу үшін әлеуметтік ғылымдар негіздерін қолдану",
]

PAGE2_SUBJECTS = [
    "КМ 1 Бизнестің мақсаттары мен түрлерін, негізгі мүдделі тараптармен өзара әрекеттесуін түсіну",
    "ОН 1.1 Бизнестің мақсаттары мен түрлерін, олардың негізгі мүдделі тараптармен және сыртқы ортамен өзара әрекеттесуін түсіну",
    "ОН 1.2 Кәсіпорындық және логарифмдік функциялар, сызықтық теңдеулер мен матрицалар жүйелері, сыныптық, теңсіздіктер және сыныптық, бағдарламалау, математикалық білу, бизнес және қаражылық қолдану мәселелерінде ақпаратты талдау және түсіндіру үшін ұрымдарды қолдана білу",
    "ОН 1.3 Қаржылық есептіліктін мәні мен мақсатын түсіну, қаржылық ақпаратты сапалық, сипаттамаларын анықтау, қаржылық есептіліктің дайындау",
    "ОН 1.4 Маркетингтің негізгі тұжырымдамаларды түсіну, маркетингтік ортаны зерттеу, тұтынушылар мен ұйымның сатып алу тәртібін түсіну, маркетингтік стратегияны және өнімдерді орналастыру, жаңа өнімдерді әзірлеу үшін қолданылатын құралдар мен әдістерді білу",
    "КМ 2 Кәсіби салада тілдік дағдыларды қолдану",
    "ОН 2.1 Академиялық, деңгейде Ағылшын тілінің оқитын, айтылым және жазылым дағдыларын орындай мәтіндерау",
    "ОН 2.2 Кәсіби салада Ағылшын тілінің айтылым және жазылым дағдыларын В2 деңгейінде еркін меңгеру",
    "ОН 2.3 Іскерлік мақсатта қазақ тілін қолдану",
    "ОН 2.4 Іскерлік мақсатта түрік тілін қолдану",
    "КМ 3 Бухгалтерлік (қаражылық) есептілікті жасауға қатысу",
    "ОН 3.1 Басқару ақпаратының сипатын, мақсатын түсіну, шығындарды есепке алу, жоспарлау, бизнестің тиімділігін бақылау",
]


PAGE3_SUBJECTS = [
    "ОН 3.2 Еңбек қатынастарына қатысты заңды түсіну, компаниялардың қалай басқарылатындығын және реттелетінін сипаттау және түсіну",
    "ОН 3.3 Іскерлік шешім қабылдау процесін қолдайтын жалпы математикалық құралдарды қолдану, аналитикалық әдістерді әртүрлі бизнес қолданбаларымда қолдану",
    "ОН 3.4 Негізгі экономикалық принциптерді, макроэкономикалық мәселелерді және көрсеткіштерді есептеуді білу, фискалдық және ақша-несие саясатының макроэкономикаға әсер ету механизмдерді талдау",
    "КМ 4 Ұйымның және оның бөлімшелерінің шаруашылық-қаржылық қызметін кешенді талдауға қатысу",
    "ОН 4.1 Инвестициялар мен қаржыландыруды бағалаудың баламалы тәсілдерін салыстыру, қаржы саласындағы проблемаларды шешудің әртүрлі тәсілдерінің орындылығын бағалау",
    "ОН 4.2 Ұйымдарға өнімділікті басқару және өлшеу үшін қажет ақпаратты, технологиялық жүйелерді анықтау, шығындарды есепке алу және басқару есебі әдістерін қолдану",
    "ОН 4.3 Салық жүйесінің жұмыс істеуі мен көлемін және оны басқаруды түсіну",
    "ОН 4.4 ХҚЕС стандарттарына сәйкес операцияларды есепке алу, Қаржылық есептерді талдау және түсіну",
    "ОН 4.5 Бизнес статистикадағы негізгі түсініктерді, деректер материалдарын жинау, қорытындылау және талдау әдістерін білу",
    "ОН 4.6 Бухгалтерлік есептің ақпараттық жүйелері",
    "ОН 4.7 Аудит ұғымының, функцияларының, корпоративтік басқарудың, оның ішінде этика мен кәсіби мінез-құлықтың анықтамасы, Халықаралық аудит стандарттарын (АХС) қолдану",
    "КМ 5 Қаржы менеджментіне экономикалық ортаның әсерін бағалау",
]

PAGE4_SUBJECTS = [
    "ОН 5.1 Қаржылық басқару функциясының рөлі мен мақсатын түсіну, Қаржы менеджментіне экономикалық ортаның әсерін бағалау",
    "ОН 5.2 Инвестицияларға тиімді бағалау жүргізу, Бизнесті қаржыландырудың балама көздерін анықтау және бағалау",
    "Кәсіптік практика КМ3. ОН3.2, ОН3.3; КМ4. ОН4.3; КМ5. ОН5.2, ОН5.3; КМ7. ОН7.1, ОН7.2, ОН7.3; КМ8. ОН8.1, ОН8.2, ОН8.3; КМ9. ОН9.1, ОН9.2, ОН9.3.",
    "Қорытынды аттестаттау :",
    "Ф1 Факультативтік ағылшын тілі",
    "Ф2 Факультативтік түрік тілі",
    "Ф3 Факультативтік Бизнес және бухгалтерлік есептегі жағдайлар (Cases in Business and Accounting)",
    "Ф4 Факультативтік Бизнес деректерін талдау (Business data analysis (excel, macros, google sheets, sql, python, power BI, tableau))",
    "Ф5 Факультативтік кәсіпкерлік қызмет негіздері (Enterpreneurship)",
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
COL_SUBJECT = 1   # B — Пән атауы
COL_HOURS   = 2   # C — Сағат
COL_CREDITS = 3   # D — Кредит
COL_POINTS  = 4   # E — Балл
COL_LETTER  = 5   # F — Әріп
COL_GPA     = 6   # G — GPA
COL_TRAD    = 7   # H — Дәстүрлі баға

# Approximate character width for subject column to calculate row height
SUBJECT_COL_CHAR_WIDTH = 42


# ─────────────────────────────────────────────────────────────
# 4. HELPER FUNCTIONS
# ─────────────────────────────────────────────────────────────

def calc_row_height(text, col_chars=SUBJECT_COL_CHAR_WIDTH, line_height=13):
    """Calculate the minimum row height needed to fit the text.

    Only increases height for multi-line subjects.
    Single-line subjects keep the default Excel row height (no stretching).
    """
    if not text:
        return None  # let Excel use default height
    num_lines = max(1, -(-len(text) // col_chars))  # ceiling division
    if num_lines <= 1:
        return None  # default height — no stretching needed
    return num_lines * line_height


def is_module_header(subject_name):
    """Check if a subject row is a module header (KM, Practice, Final) — displayed bold."""
    # Checks for "КМ", "Кәсіптік практика", "Қорытынды аттестаттау"
    return (subject_name.startswith("КМ ") or 
            subject_name.startswith("Кәсіптік практика") or 
            subject_name.startswith("Қорытынды аттестаттау"))


def setup_sheet(worksheet):
    """Apply common page/column settings to a worksheet."""
    worksheet.hide_gridlines(2)
    worksheet.set_portrait()
    worksheet.set_paper(9)  # A4
    worksheet.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)

    # Column widths
    worksheet.set_column(COL_INDEX,   COL_INDEX,   3)    # A — №
    worksheet.set_column(COL_SUBJECT, COL_SUBJECT, 32)   # B — Subject
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


def write_subjects(worksheet, styles, start_row, subjects, start_index, grades_data):
    """Write a subject block with empty grade cells."""
    current_row = start_row
    item_num = start_index
    
    # Pre-compute normalized keys for grades
    grades_map = {normalize_key(k): v for k, v in grades_data.items()}

    for subject in subjects:
        is_header = is_module_header(subject)

        # Set row height only if text overflows
        height = calc_row_height(subject)
        if height is not None:
            worksheet.set_row(current_row, height)

        # Write index
        worksheet.write(current_row, COL_INDEX, item_num, styles["index"])

        # Write subject (bold for КМ headers)
        fmt = styles["subject_bold"] if is_header else styles["subject"]
        worksheet.write(current_row, COL_SUBJECT, subject, fmt)

        # Grade cells (Cols C–H)
        if not is_header:
            # Try exact match first, then normalized
            grade = grades_data.get(subject)
            if not grade:
                grade = grades_map.get(normalize_key(subject), {})

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


def generate_diploma(data=None, output_path="diploma_v2.xlsx"):
    """Generate an Excel diploma supplement file.
    
    Args:
        data: dict with student info and grades. If None, uses empty placeholders.
        output_path: file path to save.
    """
    if data is None:
        data = get_empty_data()

    workbook = xlsxwriter.Workbook(output_path)
    styles = create_formats(workbook)

    # ═══════════════════ PAGE 1 ═══════════════════
    ws1 = workbook.add_worksheet("Бет 1")
    setup_sheet(ws1)

    # Write header with data
    header_end_row = write_header(ws1, styles, data)

    # 6 empty rows gap before subjects
    data_start = header_end_row + 6

    # Write Page 1 subjects (grades from data)
    write_subjects(ws1, styles, data_start, PAGE1_SUBJECTS, start_index=1, grades_data=data.get("grades", {}))

    # ═══════════════════ PAGE 2 ═══════════════════
    ws2 = workbook.add_worksheet("Бет 2")
    setup_sheet(ws2)

    # Page 2 starts subjects right away (row 1)
    write_subjects(ws2, styles, start_row=1, subjects=PAGE2_SUBJECTS, start_index=18, grades_data=data.get("grades", {}))

    # ═══════════════════ PAGE 3 ═══════════════════
    ws3 = workbook.add_worksheet("Бет 3")
    setup_sheet(ws3)
    write_subjects(ws3, styles, start_row=1, subjects=PAGE3_SUBJECTS, start_index=30, grades_data=data.get("grades", {}))

    # ═══════════════════ PAGE 4 ═══════════════════
    ws4 = workbook.add_worksheet("Бет 4")
    setup_sheet(ws4)
    write_subjects(ws4, styles, start_row=1, subjects=PAGE4_SUBJECTS, start_index=42, grades_data=data.get("grades", {}))
    
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
            # Header: Write ONLY Hours and Credits (if aggregation populated them)
            # Leave Grade/Points/Letter blank
            worksheet.write(current_row, COL_HOURS,   grade.get("hours", ""),   styles["grade_center"])
            worksheet.write(current_row, COL_CREDITS, grade.get("credits", ""), styles["grade_center"])
            
        else:
            # Regular Subject: Write ALL columns
            
            # Special Case: Electives (Факультативтік) -> Traditional = "сынақ"
            trad_val = grade.get("traditional", "")
            if subject.startswith("Ф") and "Факультативтік" in subject:
                trad_val = "сынақ"

            worksheet.write(current_row, COL_HOURS,   grade.get("hours", ""),   styles["grade_center"])
            worksheet.write(current_row, COL_CREDITS, grade.get("credits", ""), styles["grade_center"])
            worksheet.write(current_row, COL_POINTS,  grade.get("points", ""),  styles["grade_center"])
            worksheet.write(current_row, COL_LETTER,  grade.get("letter", ""),  styles["grade_center"])
            worksheet.write(current_row, COL_GPA,     grade.get("gpa", ""),     styles["grade_center"])
            worksheet.write(current_row, COL_TRAD,    trad_val,                 styles["grade_left"])

        item_num += 1
        current_row += 1

    return current_row


if __name__ == "__main__":
    # Example: Generate with empty data
    output_file = generate_diploma(output_path="diploma_v4.xlsx")

    # Example: How to likely use it later
    # my_data = {
    #     "full_name": "Иванов Иван",
    #     "grades": {
    #         "Қазақ тілі": {"hours": 96, "points": 90, ...}
    #     }
    # }
    # generate_diploma(my_data, "diploma_filled.xlsx")

    try:
        os.startfile(output_file)
    except Exception:
        pass
