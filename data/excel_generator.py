# -*- coding: utf-8 -*-
"""
Unified Diploma Generator
============================
Generates Excel diploma supplement files matching the official template format.

Supports:
- Bilingual output (KZ / RU)
- 8-column layout: №, Subject, Hours, Credits, Points, Letter, GPA, Traditional
- Proper Times New Roman fonts, merged header cells
- Module headers (КМ/ПМ/БМ) with summed hours/credits, no grade columns
- Elective subjects with "сынақ"/"зачтено" labels
- 4-page pagination matching the real diploma template

Usage:
    from data.excel_generator import DiplomaGenerator
    gen = DiplomaGenerator(Program.IT, Language.KZ)
    diploma_bytes = gen.generate(student, subjects, "letter")
"""

import io
import re
import xlsxwriter
from typing import List, Dict, Optional

from core.models import Student, Grade, Subject, Language, Program
from config.languages import ELECTIVE_GRADES
from core.utils import is_module_header, normalize_key


# ─────────────────────────────────────────────────────────────
# HARDCODED SUBJECT LISTS (from official templates)
# ─────────────────────────────────────────────────────────────

_PAGE1_KZ = [
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

_PAGE2_KZ = [
    "БМ 01 Дене қасиеттерін дамыту және жетілдіру",
    "БМ 02 Ақпараттық-коммуникациялық және цифрлық технологияларды қолдану",
    "БМ 03 Экономиканың базалық білімін және кәсіпкерлік негіздерін қолдану",
    "БМ 04 Қоғам мен еңбек ұжымында әлеуметтену және бейімделу үшін әлеуметтік ғылымдар негіздерін қолдану",
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

_PAGE3_KZ = [
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

_PAGE4_KZ = [
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

_PAGE1_RU = [
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

_PAGE2_RU = [
    "БМ 01 Развитие и совершенств. физических качеств",
    "БМ 02 Применение информационно-коммуникационных и цифровых технологий",
    "БМ 03 Применение базовых знаний экономики и основ предпринимательства",
    "БМ 04 Применение основ социальных наук для социализации и адаптации в обществе и трудовом коллективе",
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

_PAGE3_RU = [
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

_PAGE4_RU = [
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

_PAGES = {
    Language.KZ: [_PAGE1_KZ, _PAGE2_KZ, _PAGE3_KZ, _PAGE4_KZ],
    Language.RU: [_PAGE1_RU, _PAGE2_RU, _PAGE3_RU, _PAGE4_RU],
}


# ─────────────────────────────────────────────────────────────
# COLUMN LAYOUT
# ─────────────────────────────────────────────────────────────

COL_INDEX   = 0   # A — №
COL_SUBJECT = 1   # B — Subject name
COL_HOURS   = 2   # C — Hours
COL_CREDITS = 3   # D — Credits
COL_POINTS  = 4   # E — Points (percentage)
COL_LETTER  = 5   # F — Letter grade
COL_GPA     = 6   # G — GPA
COL_TRAD    = 7   # H — Traditional grade

SUBJECT_COL_CHAR_WIDTH = 45


# ─────────────────────────────────────────────────────────────
# EXCEPTIONS
# ─────────────────────────────────────────────────────────────

class DiplomaGenerationError(Exception):
    """Raised when diploma generation fails."""
    pass


# ─────────────────────────────────────────────────────────────
# GENERATOR CLASS
# ─────────────────────────────────────────────────────────────

class DiplomaGenerator:
    """
    Generates Excel diploma supplement files matching the official template.

    Parameters:
        program: Program type (IT, ACCOUNTING)
        language: Output language (KZ, RU)
        academic_year: Academic year string (default "2025-2026")
    """

    def __init__(
        self,
        program: Program,
        language: Language,
        academic_year: str = "2025-2026",
    ):
        self.program = program
        self.language = language
        self.academic_year = academic_year

    # ── Public API ──

    def generate(
        self,
        student: Student,
        subjects: List[Subject],
        grade_format: str = "letter",
    ) -> bytes:
        """
        Generate diploma as in-memory bytes.

        Args:
            student: Student object with grades
            subjects: List of Subject objects (with hours/credits)
            grade_format: Not used currently (kept for API compatibility)

        Returns:
            bytes: Excel file content
        """
        buf = io.BytesIO()
        self._build_workbook(buf, student, subjects)
        buf.seek(0)
        return buf.read()

    def generate_to_file(
        self,
        student: Student,
        subjects: List[Subject],
        output_path: str = "diploma.xlsx",
        grade_format: str = "letter",
    ) -> str:
        """
        Generate diploma and write to file.

        Args:
            student: Student object with grades
            subjects: List of Subject objects
            output_path: File path for output
            grade_format: Not used currently

        Returns:
            str: Path to generated file
        """
        data = self.generate(student, subjects, grade_format)
        with open(output_path, "wb") as f:
            f.write(data)
        return output_path

    # ── Internal ──

    def _build_workbook(self, output, student, subjects):
        """Build the complete workbook."""
        workbook = xlsxwriter.Workbook(output, {"in_memory": True})
        styles = self._create_formats(workbook)

        # Build grades_data dict (subject_name → {hours, credits, points, ...})
        grades_data = self._build_grades_data(student, subjects)

        # Get page subjects for this language
        pages = _PAGES.get(self.language, _PAGES[Language.KZ])

        # ═══ PAGE 1 (with header) ═══
        sheet_name_1 = self._sheet_name(1)
        ws1 = workbook.add_worksheet(sheet_name_1)
        self._setup_sheet(ws1)

        header_data = self._build_header_data(student)
        header_end = self._write_header(ws1, styles, header_data)
        data_start = header_end + 6
        idx = 1
        self._write_subjects(ws1, styles, data_start, pages[0], idx, grades_data)

        # ═══ PAGES 2-4 (no header) ═══
        idx = 1 + len(pages[0])
        for page_num in range(2, len(pages) + 1):
            ws = workbook.add_worksheet(self._sheet_name(page_num))
            self._setup_sheet(ws)
            self._write_subjects(ws, styles, 1, pages[page_num - 1], idx, grades_data)
            idx += len(pages[page_num - 1])

        workbook.close()

    def _sheet_name(self, page_num: int) -> str:
        """Get sheet name for page number."""
        if self.language == Language.KZ:
            return f"Бет {page_num}"
        else:
            return f"Лист {page_num}"

    def _setup_sheet(self, ws):
        """Apply common page/column settings."""
        ws.hide_gridlines(2)
        ws.set_portrait()
        ws.set_paper(9)  # A4
        ws.set_margins(left=0.5, right=0.5, top=0.5, bottom=0.5)
        ws.set_column(COL_INDEX, COL_INDEX, 4)
        ws.set_column(COL_SUBJECT, COL_SUBJECT, 30)
        ws.set_column(COL_HOURS, COL_HOURS, 5)
        ws.set_column(COL_CREDITS, COL_CREDITS, 4)
        ws.set_column(COL_POINTS, COL_POINTS, 5)
        ws.set_column(COL_LETTER, COL_LETTER, 4)
        ws.set_column(COL_GPA, COL_GPA, 5)
        ws.set_column(COL_TRAD, COL_TRAD, 13)

    def _organize_subjects_into_pages(self, subjects: List[Subject]) -> Dict[int, list]:
        """Organize subjects into pages. Returns dict of page_num → subject_list."""
        pages = _PAGES.get(self.language, _PAGES[Language.KZ])
        result = {}
        for i, page in enumerate(pages, 1):
            result[i] = page
        return result

    def _create_formats(self, workbook) -> dict:
        """Create all reusable cell formats."""
        styles = {}

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

        styles["index"] = workbook.add_format({
            "font_name": "Times New Roman",
            "font_size": 8,
            "align": "center",
            "valign": "top",
        })

        styles["subject"] = workbook.add_format({
            "font_name": "Times New Roman",
            "font_size": 8,
            "align": "left",
            "valign": "top",
            "text_wrap": True,
        })

        styles["subject_bold"] = workbook.add_format({
            "font_name": "Times New Roman",
            "font_size": 8,
            "bold": True,
            "align": "left",
            "valign": "top",
            "text_wrap": True,
        })

        styles["grade_center"] = workbook.add_format({
            "font_name": "Times New Roman",
            "font_size": 8,
            "align": "center",
            "valign": "top",
        })

        styles["grade_left"] = workbook.add_format({
            "font_name": "Times New Roman",
            "font_size": 8,
            "align": "left",
            "valign": "top",
        })

        return styles

    def _build_header_data(self, student: Student) -> dict:
        """Build header data dict from Student object."""
        start, end = self._split_academic_year()
        data = {
            "diploma_id": student.diploma_number or "",
            "start_year": start,
            "end_year": end,
        }

        if self.language == Language.KZ:
            data["full_name"] = student.full_name
            data["college_name"] = ""
            data["specialization"] = ""
            data["qualification"] = ""
        else:
            data["full_name"] = student.full_name
            data["college_name"] = ""
            data["specialization"] = ""
            data["qualification"] = ""

        return data

    def _split_academic_year(self):
        """Split '2025-2026' into ('2025', '2026')."""
        parts = self.academic_year.split("-")
        if len(parts) == 2:
            return parts[0].strip(), parts[1].strip()
        return self.academic_year, ""

    def _write_header(self, ws, styles, data) -> int:
        """Write header section. Returns row after header."""
        row = 1
        ws.merge_range(row, 0, row, 7, data.get("diploma_id", ""), styles["title_bold"])
        row += 2
        ws.merge_range(row, 0, row, 7, data.get("full_name", ""), styles["title_bold"])
        row += 2
        ws.merge_range(row, 0, row, 3, data.get("start_year", ""), styles["year_bold"])
        ws.merge_range(row, 4, row, 7, data.get("end_year", ""), styles["year_bold"])
        row += 1
        ws.merge_range(row, 0, row, 7, data.get("college_name", ""), styles["header_center"])
        row += 2
        ws.merge_range(row, 0, row, 7, data.get("specialization", ""), styles["header_center"])
        row += 3
        ws.merge_range(row, 0, row, 7, data.get("qualification", ""), styles["header_center"])
        return row + 1

    def _build_grades_data(self, student: Student, subjects: List[Subject]) -> dict:
        """
        Build grades_data dict from Student and Subject objects.

        Returns dict: subject_name → {hours, credits, points, letter, gpa, traditional}
        """
        grades_data = {}
        norm_map = {}  # normalized_key → grade_dict

        for subj in subjects:
            kz_name = subj.name_kz
            ru_name = subj.name_ru

            # Look up grade by KZ or RU name
            grade = student.grades.get(kz_name) or student.grades.get(ru_name)

            entry = {
                "hours": subj.hours or "",
                "credits": subj.credits or "",
                "points": "",
                "letter": "",
                "gpa": "",
                "traditional": "",
            }

            if grade and not grade.is_empty():
                entry["points"] = grade.points or ""
                entry["letter"] = grade.letter or ""
                entry["gpa"] = str(grade.gpa) if grade.gpa else ""
                if self.language == Language.KZ:
                    entry["traditional"] = grade.traditional_kz or ""
                else:
                    entry["traditional"] = grade.traditional_ru or ""

            # Store under both KZ and RU names
            grades_data[kz_name] = entry
            grades_data[ru_name] = entry
            norm_map[normalize_key(kz_name)] = entry
            norm_map[normalize_key(ru_name)] = entry

        # Merge norm_map into grades_data for fallback lookup
        grades_data["__norm__"] = norm_map
        return grades_data

    def _lookup_grade(self, subject_name: str, grades_data: dict) -> dict:
        """Look up grade data for a subject name with fuzzy matching."""
        # Direct lookup
        g = grades_data.get(subject_name)
        if g and g != grades_data.get("__norm__"):
            return g

        # Normalized lookup
        norm_map = grades_data.get("__norm__", {})
        nk = normalize_key(subject_name)
        g = norm_map.get(nk)
        if g:
            return g

        # Fuzzy: practice / attestation
        s_lower = subject_name.lower()
        if "кәсіптік практика" in s_lower or "профессиональная практика" in s_lower:
            for k, v in norm_map.items():
                if "кәсіптікпрактика" in k or "профессиональнаяпрактика" in k:
                    return v
        if "аттестаттау" in s_lower or "аттестация" in s_lower:
            for k, v in norm_map.items():
                if "аттестаттау" in k or "аттестация" in k:
                    return v

        return {}

    def _write_subjects(self, ws, styles, start_row, subjects, start_index, grades_data):
        """Write a block of subjects with grades."""
        current_row = start_row
        item_num = start_index
        norm_map = grades_data.get("__norm__", {})

        for i, subject in enumerate(subjects):
            is_header = is_module_header(subject)

            # Row height
            height = _calc_row_height(subject)
            if height is not None:
                ws.set_row(current_row, height)

            # Index column
            ws.write(current_row, COL_INDEX, item_num, styles["index"])

            # Subject name (bold for module headers)
            fmt = styles["subject_bold"] if is_header else styles["subject"]
            ws.write(current_row, COL_SUBJECT, subject, fmt)

            if is_header:
                self._write_module_header_grades(
                    ws, styles, current_row, i, subjects, grades_data, norm_map
                )
            elif subject.startswith("Ф") and "актив" in subject.lower():
                # Elective
                self._write_elective_grades(ws, styles, current_row, subject, grades_data)
            else:
                # Normal subject
                self._write_normal_grades(ws, styles, current_row, subject, grades_data)

            item_num += 1
            current_row += 1

        return current_row

    def _write_normal_grades(self, ws, styles, row, subject, grades_data):
        """Write grade cells for a normal subject."""
        grade = self._lookup_grade(subject, grades_data)

        ws.write(row, COL_HOURS, grade.get("hours", ""), styles["grade_center"])
        ws.write(row, COL_CREDITS, grade.get("credits", ""), styles["grade_center"])
        ws.write(row, COL_POINTS, grade.get("points", ""), styles["grade_center"])
        ws.write(row, COL_LETTER, grade.get("letter", ""), styles["grade_center"])
        ws.write(row, COL_GPA, grade.get("gpa", ""), styles["grade_center"])
        ws.write(row, COL_TRAD, grade.get("traditional", ""), styles["grade_left"])

    def _write_elective_grades(self, ws, styles, row, subject, grades_data):
        """Write grade cells for an elective (Факультатив)."""
        grade = self._lookup_grade(subject, grades_data)

        ws.write(row, COL_HOURS, grade.get("hours", ""), styles["grade_center"])
        ws.write(row, COL_CREDITS, grade.get("credits", ""), styles["grade_center"])
        ws.write(row, COL_POINTS, grade.get("points", ""), styles["grade_center"])
        ws.write(row, COL_LETTER, grade.get("letter", ""), styles["grade_center"])
        ws.write(row, COL_GPA, grade.get("gpa", ""), styles["grade_center"])

        # Elective traditional: "сынақ" (KZ) or "зачтено" (RU)
        trad = ELECTIVE_GRADES.get(self.language.value, "")
        ws.write(row, COL_TRAD, trad, styles["grade_left"])

    def _write_module_header_grades(self, ws, styles, row, i, subjects, grades_data, norm_map):
        """Write grade cells for a module header (КМ/ПМ/БМ)."""
        # Sum hours/credits from sub-subjects
        total_hours = 0.0
        total_credits = 0.0

        for sub_idx in range(i + 1, len(subjects)):
            sub_name = subjects[sub_idx]
            if is_module_header(sub_name):
                break
            g = self._lookup_grade(sub_name, grades_data)
            if g:
                try:
                    total_hours += float(str(g.get("hours") or "0").replace(",", "."))
                except (ValueError, TypeError):
                    pass
                try:
                    total_credits += float(str(g.get("credits") or "0").replace(",", "."))
                except (ValueError, TypeError):
                    pass

        def _fmt(v):
            return str(int(v)) if v == int(v) else str(v)

        if total_hours > 0:
            ws.write(row, COL_HOURS, _fmt(total_hours), styles["grade_center"])
            ws.write(row, COL_CREDITS, _fmt(total_credits), styles["grade_center"])
        else:
            # Direct lookup for terminal modules (БМ, practice, attestation)
            g = self._lookup_grade(subjects[i], grades_data)
            if g:
                ws.write(row, COL_HOURS, g.get("hours", ""), styles["grade_center"])
                ws.write(row, COL_CREDITS, g.get("credits", ""), styles["grade_center"])

        # Module headers NEVER show points/letter/GPA/traditional


# ─────────────────────────────────────────────────────────────
# MODULE-LEVEL HELPERS
# ─────────────────────────────────────────────────────────────

def _calc_row_height(text, col_chars=SUBJECT_COL_CHAR_WIDTH, line_height=11):
    """Calculate minimum row height for text wrapping."""
    if not text:
        return None
    lines = text.split("\n")
    total_lines = 0
    for line in lines:
        total_lines += max(1, -(-len(line) // col_chars))
    if total_lines <= 1:
        return 12
    return total_lines * line_height
