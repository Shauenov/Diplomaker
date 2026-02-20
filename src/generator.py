import os
import shutil
import xlsxwriter
import datetime
import math
from typing import Dict, Any, List

class DiplomaGenerator:
    """Единый класс для генерации Excel-дипломов на основе шаблона."""
    
    def __init__(self, template_path: str, output_path: str, config: Dict[str, Any], terms: Dict[str, str]):
        self.template_path = template_path
        self.output_path = output_path
        self.config = config
        self.terms = terms
        
        # Копируем шаблон
        shutil.copyfile(template_path, output_path)
        self.workbook = xlsxwriter.Workbook(output_path)
        self.worksheet = self.workbook.add_worksheet('Template')
        self.setup_styles()
        self.setup_page_layout()

    def setup_styles(self):
        """Создает стили для ячеек."""
        self.styles = {
            'text_center': self.workbook.add_format({
                'font_name': 'Times New Roman', 'font_size': 10,
                'align': 'center', 'valign': 'vcenter', 'text_wrap': True
            }),
            'text_left': self.workbook.add_format({
                'font_name': 'Times New Roman', 'font_size': 10,
                'align': 'left', 'valign': 'vcenter', 'text_wrap': True
            }),
            'text_bold_center': self.workbook.add_format({
                'font_name': 'Times New Roman', 'font_size': 10, 'bold': True,
                'align': 'center', 'valign': 'vcenter', 'text_wrap': True
            }),
            'text_bold_left': self.workbook.add_format({
                'font_name': 'Times New Roman', 'font_size': 10, 'bold': True,
                'align': 'left', 'valign': 'vcenter', 'text_wrap': True
            }),
            'header_large_bold': self.workbook.add_format({
                'font_name': 'Times New Roman', 'font_size': 11, 'bold': True,
                'align': 'center', 'valign': 'vcenter', 'text_wrap': True
            })
        }

    def setup_page_layout(self):
        """Настройка ширины колонок для A4 альбомной ориентации."""
        self.worksheet.set_paper(9) # A4
        self.worksheet.set_landscape()
        self.worksheet.set_margins(left=0.1, right=0.1, top=0.1, bottom=0.1)
        
        # Левая страница (стр 2/4)
        self.worksheet.set_column('A:A', 3.3)
        self.worksheet.set_column('B:B', 28.5)
        self.worksheet.set_column('C:C', 5.6)
        self.worksheet.set_column('D:D', 5.3)
        self.worksheet.set_column('E:E', 4.3)
        self.worksheet.set_column('F:F', 3.6)
        self.worksheet.set_column('G:G', 3.6)
        self.worksheet.set_column('H:H', 11.0)
        
        self.worksheet.set_column('I:I', 3.0) # Отступ
        
        # Правая страница (стр 3/1)
        self.worksheet.set_column('J:J', 3.3)
        self.worksheet.set_column('K:K', 28.5)
        self.worksheet.set_column('L:L', 5.6)
        self.worksheet.set_column('M:M', 5.3)
        self.worksheet.set_column('N:N', 4.3)
        self.worksheet.set_column('O:O', 3.6)
        self.worksheet.set_column('P:P', 3.6)
        self.worksheet.set_column('Q:Q', 11.0)

    def is_module_header(self, subj_name: str) -> bool:
        """Определяет, является ли предмет заголовком модуля."""
        s = subj_name.strip()
        # БМ не являются заголовками (это самостоятельные предметы)
        return s.startswith("КМ") or s.startswith("ПМ") or \
               s.startswith("Кәсіптік модуль") or s.startswith("Профессиональная практика") or \
               s.startswith("Қорытынды аттестаттау") or s.startswith("Итоговая аттестация")

    def calc_row_height(self, text: str, width_chars: int = 40) -> float:
        """Рассчитывает высоту строки в зависимости от длины текста."""
        chars = len(str(text))
        lines = math.ceil(chars / width_chars)
        return max(15.0, lines * 13.0)

    def fill_student_data(self, student: Dict[str, Any]):
        """Заполняет диплом конкретными данными."""
        # TODO: Заполнение шапки диплома (ФИО, дата, протокол)
        self.worksheet.write('K15', student['name'], self.styles['header_large_bold'])
        
        # Дата выдачи (текущая дата для примера)
        date_str = datetime.datetime.now().strftime("%d.%m.%Y")
        self.worksheet.write('L35', date_str, self.styles['text_center'])
        
        # Заполнение таблиц с предметами (Страницы 1, 2, 3, 4)
        self._write_subjects(student['grades'], self.config['p1'], start_row=45, col_offset=9) # Page 1 (справа внизу)
        self._write_subjects(student['grades'], self.config['p2'], start_row=1,  col_offset=0) # Page 2 (слева сверху)
        self._write_subjects(student['grades'], self.config['p3'], start_row=1,  col_offset=9) # Page 3 (справа сверху)
        self._write_subjects(student['grades'], self.config['p4'], start_row=46, col_offset=0) # Page 4 (слева внизу)

    def _write_subjects(self, grades: Dict[str, Any], template_subjects: List[str], start_row: int, col_offset: int):
        """Записывает блок предметов на определенную 'страницу'."""
        current_row = start_row
        from .utils import normalize_key
        
        # Инициализируем сумму часов и кредитов
        total_h, total_c = 0.0, 0.0
        
        for idx, subj in enumerate(template_subjects):
            if subj.strip() in ('', ' '):
                self.worksheet.set_row(current_row, 15)
                current_row += 1
                continue
                
            is_header = self.is_module_header(subj)
            style_name = self.styles['text_bold_left'] if is_header else self.styles['text_left']
            style_center = self.styles['text_bold_center'] if is_header else self.styles['text_center']
            
            # Пишем № п/п и название
            self.worksheet.write(current_row, col_offset, idx + 1 if not is_header else "", style_center)
            self.worksheet.write(current_row, col_offset + 1, subj, style_name)
            self.worksheet.set_row(current_row, self.calc_row_height(subj))
            
            nkey = normalize_key(subj)
            grade = grades.get(nkey)
            
            # Фолбэк для БМ (БМ 1 -> Базалық модульдер)
            if not grade and re.match(r'(БМ|КМ|ПМ)\s*\.?\s*\d+', subj):
                prefix = normalize_key(re.match(r'(БМ|КМ|ПМ)\s*\.?\s*\d+', subj).group(1))
                for gk, gv in grades.items():
                    if prefix in gk:
                        grade = gv
                        break

            if is_header:
                # Заголовки: собираем сумму часов до следующего заголовка
                # В Excel агрегациях эта сумма уже есть, используем ее, если найдем, иначе считаем сами
                if grade and grade.get('hours'):
                    self.worksheet.write(current_row, col_offset + 2, grade['hours'], style_center)
                    self.worksheet.write(current_row, col_offset + 3, grade['credits'], style_center)
                    total_h += float(grade['hours']) if str(grade['hours']).replace('.', '').isdigit() else 0
                    total_c += float(grade['credits']) if str(grade['credits']).replace('.', '').isdigit() else 0
            else:
                # Обычный предмет
                if grade:
                    # Факультатив
                    trad_val = grade.get('traditional_kz' if 'kz' in subj.lower() else 'traditional_ru', "")
                    if "факультатив" in subj.lower(): trad_val = self.terms.get('traditional_elective', 'зачтено')
                    if "практика" in subj.lower(): trad_val = self.terms.get('traditional_practice', 'зачтено')
                    
                    self.worksheet.write(current_row, col_offset + 2, grade.get('hours', ''), style_center)
                    self.worksheet.write(current_row, col_offset + 3, grade.get('credits', ''), style_center)
                    self.worksheet.write(current_row, col_offset + 4, grade.get('points', ''), style_center)
                    self.worksheet.write(current_row, col_offset + 5, grade.get('letter', ''), style_center)
                    self.worksheet.write(current_row, col_offset + 6, grade.get('gpa', ''), style_center)
                    self.worksheet.write(current_row, col_offset + 7, trad_val, self.styles['text_left'])
                    
            current_row += 1

    def close(self):
        """Закрывает и сохраняет файл."""
        self.workbook.close()
