import pandas as pd
from typing import Dict, Any, List
from .utils import normalize_key, parse_hours_credits, safe_str, clean_name

def parse_excel_sheet(df: pd.DataFrame, sheet_name: str, start_row: int = 4) -> List[Dict[str, Any]]:
    """
    Парсит лист DataFrame (данные группы) и возвращает список студентов.
    
    Ожидаемая архитектура:
    - row 1 (индекс 1): названия предметов (могут быть билингвальными \n)
    - row 3 (индекс 3): часы/кредиты '72с-3к'
    - Начиная со start_row (index=start_row) - данные конкретных студентов.
    """
    students = []
    
    # 1. Извлекаем предметы и часы (названия колонок)
    row_subjects = df.iloc[1]
    row_hours = df.iloc[3]
    
    # Кэш колонок: col_index -> { 'kz': .., 'ru': .., 'nkz':.., 'nru':.., 'hours':.., 'credits':.. }
    col_dict = {}
    col = 2
    while col < len(row_subjects):
        cv = row_subjects.iloc[col]
        if pd.isna(cv) or str(cv).strip() in ('', 'nan'):
            col += 4
            continue
            
        raw = str(cv).strip()
        parts = raw.split('\n')
        kz_name = parts[0].strip().rstrip(':')
        ru_name = parts[1].strip().rstrip(':') if len(parts) >= 2 else kz_name
        
        h_str = safe_str(row_hours.iloc[col]) if col < len(row_hours) else ""
        hours, credits = parse_hours_credits(h_str)
        
        col_dict[col] = {
            'kz': kz_name, 'ru': ru_name,
            'nkz': normalize_key(kz_name), 'nru': normalize_key(ru_name),
            'hours': hours, 'credits': credits
        }
        col += 4

    # 2. Итерируемся по строкам студентов
    for i in range(start_row, len(df)):
        row = df.iloc[i]
        
        # Индекс и ФИО
        s_idx = safe_str(row.iloc[0])
        s_name = clean_name(row.iloc[1])
        if not s_name or str(s_idx).lower() == 'nan':
            continue # Пустая строка
            
        # Паттерн "Руководитель практики" или подписи - прерываем парсинг
        if any(w in s_idx.lower() for w in ['руководитель', 'директор', 'заместитель']):
            break
            
        # Тема диплома
        diploma_kz = safe_str(row.iloc[-5]) if len(row) > 5 else ""
        diploma_ru = safe_str(row.iloc[-4]) if len(row) > 4 else ""
        
        # Собираем оценки
        grades = {}
        for c_idx, subj_info in col_dict.items():
            pts_val = row.iloc[c_idx]
            if pd.isna(pts_val):
                continue
            try:
                pts = float(str(pts_val).replace(',', '.'))
            except ValueError:
                # Если оценка "зачтено", "босатылды" и др. текстом
                val_str = str(pts_val).strip()
                if val_str:
                    grades[subj_info['nkz']] = {
                        'subject_kz': subj_info['kz'], 'subject_ru': subj_info['ru'],
                        'hours': subj_info['hours'], 'credits': subj_info['credits'],
                        'points': '', 'letter': '', 'gpa': '', 'traditional': val_str
                    }
                continue
                
            from .utils import calc_letter_grade, calc_gpa_grade, calc_traditional_grade
            
            letter = calc_letter_grade(pts)
            gpa = f"{calc_gpa_grade(pts):.2f}"
            gpa = str(int(float(gpa))) if gpa.endswith('.00') else gpa
            
            grades[subj_info['nkz']] = {
                'subject_kz': subj_info['kz'],
                'subject_ru': subj_info['ru'],
                'hours': subj_info['hours'],
                'credits': subj_info['credits'],
                'points': str(int(pts)) if pts.is_integer() else str(pts),
                'letter': letter,
                'gpa': gpa,
                'traditional_kz': calc_traditional_grade(pts, True),
                'traditional_ru': calc_traditional_grade(pts, False)
            }
            
            # Добавим и русский ключ для удобства
            grades[subj_info['nru']] = grades[subj_info['nkz']]

        students.append({
            'name': s_name,
            'sheet': sheet_name,
            'diploma_kz': diploma_kz,
            'diploma_ru': diploma_ru,
            'grades': grades
        })
        
    return students
