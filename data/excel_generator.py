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
import xlsxwriter
from typing import List, Dict, Optional

from core.models import Student, Grade, Subject, Language, Program
from config.languages import ELECTIVE_GRADES
from core.utils import is_module_header, normalize_key
from config.programs import PROGRAM_ACCOUNTING_PAGES, PROGRAM_IT_PAGES

# ─────────────────────────────────────────────────────────────
# HARDCODED SUBJECT LISTS ARE NOW IN config/programs.py
# ─────────────────────────────────────────────────────────────

_PAGES = {}
# ─────────────────────────────────────────────────────────────

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
        pages = {}
        if self.program == Program.ACCOUNTING:
            pages = PROGRAM_ACCOUNTING_PAGES
        elif self.program == Program.IT:
            pages = PROGRAM_IT_PAGES

        # ═══ PAGE 1 (with header) ═══
        sheet_name_1 = self._sheet_name(1)
        ws1 = workbook.add_worksheet(sheet_name_1)
        self._setup_sheet(ws1)

        header_data = self._build_header_data(student)
        header_end = self._write_header(ws1, styles, header_data)
        data_start = header_end + 6
        idx = 1
        self._write_subjects(ws1, styles, data_start, pages[1], idx, grades_data)

        # ═══ PAGES 2-4 (no header) ═══
        idx = 1 + len(pages[1])
        for page_num in range(2, len(pages) + 1):
            ws = workbook.add_worksheet(self._sheet_name(page_num))
            self._setup_sheet(ws)
            self._write_subjects(ws, styles, 1, pages[page_num], idx, grades_data)
            idx += len(pages[page_num])

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
        pages = {}
        if self.program == Program.ACCOUNTING:
            pages = PROGRAM_ACCOUNTING_PAGES
        elif self.program == Program.IT:
            pages = PROGRAM_IT_PAGES
            
        return dict(pages)

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
        if "аттестаттау" in s_lower or "аттестациялау" in s_lower or "аттестация" in s_lower:
            for k, v in norm_map.items():
                if "аттестаттау" in k or "аттестациялау" in k or "аттестация" in k:
                    return v

        return {}

    def _write_subjects(self, ws, styles, start_row, subjects, start_index, grades_data):
        """Write a block of subjects with grades."""
        current_row = start_row
        item_num = start_index
        norm_map = grades_data.get("__norm__", {})

        for i, subject_obj in enumerate(subjects):
            # Support both string and Subject dataclass
            is_obj = isinstance(subject_obj, Subject)
            subject_name = subject_obj.name_kz if is_obj and self.language == Language.KZ else (subject_obj.name_ru if is_obj else subject_obj)
            is_header = subject_obj.is_module_header if is_obj else is_module_header(subject_name)

            # Row height
            height = _calc_row_height(subject_name)
            if height is not None:
                ws.set_row(current_row, height)

            # Index column
            ws.write(current_row, COL_INDEX, item_num, styles["index"])

            # Subject name (bold for module headers)
            fmt = styles["subject_bold"] if is_header else styles["subject"]
            ws.write(current_row, COL_SUBJECT, subject_name, fmt)

            if is_header:
                self._write_module_header_grades(
                    ws, styles, current_row, i, subjects, grades_data, norm_map
                )
            elif (is_obj and subject_obj.is_elective) or (not is_obj and subject_name.startswith("Ф") and "актив" in subject_name.lower()):
                # Elective
                self._write_elective_grades(ws, styles, current_row, subject_name, grades_data, subject_obj if is_obj else None)
            else:
                # Normal subject
                self._write_normal_grades(ws, styles, current_row, subject_name, grades_data, subject_obj if is_obj else None)

            item_num += 1
            current_row += 1

        return current_row

    def _write_normal_grades(self, ws, styles, row, subject, grades_data, subj_obj=None):
        """Write grade cells for a normal subject."""
        grade = self._lookup_grade(subject, grades_data) if grades_data else {}

        # Prefer grades if available, else static obj info
        hours = grade.get("hours", "")
        if not hours and subj_obj and subj_obj.hours:
            hours = subj_obj.hours
            
        credits = grade.get("credits", "")
        if not credits and subj_obj and subj_obj.credits:
            credits = subj_obj.credits

        ws.write(row, COL_HOURS, hours, styles["grade_center"])
        ws.write(row, COL_CREDITS, credits, styles["grade_center"])
        ws.write(row, COL_POINTS, grade.get("points", ""), styles["grade_center"])
        ws.write(row, COL_LETTER, grade.get("letter", ""), styles["grade_center"])
        ws.write(row, COL_GPA, grade.get("gpa", ""), styles["grade_center"])
        ws.write(row, COL_TRAD, grade.get("traditional", ""), styles["grade_left"])

    def _write_elective_grades(self, ws, styles, row, subject, grades_data, subj_obj=None):
        """Write grade cells for an elective (Факультатив)."""
        grade = self._lookup_grade(subject, grades_data) if grades_data else {}

        # Prefer grades if available, else static obj info
        hours = grade.get("hours", "")
        if not hours and subj_obj and subj_obj.hours:
            hours = subj_obj.hours
            
        credits = grade.get("credits", "")
        if not credits and subj_obj and subj_obj.credits:
            credits = subj_obj.credits

        ws.write(row, COL_HOURS, hours, styles["grade_center"])
        ws.write(row, COL_CREDITS, credits, styles["grade_center"])
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
            sub_obj = subjects[sub_idx]
            is_obj = isinstance(sub_obj, Subject)
            sub_name = sub_obj.name_kz if is_obj and self.language == Language.KZ else (sub_obj.name_ru if is_obj else sub_obj)
            
            is_header = sub_obj.is_module_header if is_obj else is_module_header(sub_name)
            if is_header:
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
            subj_obj = subjects[i]
            is_obj = isinstance(subj_obj, Subject)
            sub_name = subj_obj.name_kz if is_obj and self.language == Language.KZ else (subj_obj.name_ru if is_obj else subj_obj)
            
            g = self._lookup_grade(sub_name, grades_data)
            
            hours = g.get("hours", "") if g else ""
            if not hours and is_obj and subj_obj.hours:
                hours = subj_obj.hours
                
            credits = g.get("credits", "") if g else ""
            if not credits and is_obj and subj_obj.credits:
                credits = subj_obj.credits
                
            ws.write(row, COL_HOURS, hours, styles["grade_center"])
            ws.write(row, COL_CREDITS, credits, styles["grade_center"])

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
