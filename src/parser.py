import pandas as pd
import re as _re
from typing import Dict, Any, List
from .utils import normalize_key, parse_hours_credits, safe_str, clean_name


def _parse_sheet_meta(df: pd.DataFrame) -> Dict[str, str]:
    """Извлекает метаданные листа: специальность, квалификация, годы обучения.
    
    Колонка 1 (B), строка 1 содержит многострочный текст вида:
        3F - 1
        Мамандық: 06130100 - Бағдарламалық қамтамасыз ету
        Біліктілік: 4S06130105 - Ақпараттық жүйелер технигі
        
        Специальность: 06130100 - Программное обеспечение
        Квалификация: 4S06130105 - Техник информационных систем

    Последние колонки (row 3 = header, row 4 = value) содержат:
        Год поступления / Год выпуска / Диплом (номер)
    """
    meta: Dict[str, str] = {
        'specialty_kz': '', 'qualification_kz': '',
        'specialty_ru': '', 'qualification_ru': ''
    }
    
    # --- Разбор Col 1 Row 1 ---
    raw = str(df.iloc[1, 1]) if not pd.isna(df.iloc[1, 1]) else ''
    for line in raw.split('\n'):
        line = line.strip()
        if not line:
            continue
        low = line.lower()
        if low.startswith('мамандық'):
            meta['specialty_kz'] = line.split(':', 1)[-1].strip()
        elif low.startswith('біліктілік'):
            meta['qualification_kz'] = line.split(':', 1)[-1].strip()
        elif low.startswith('спец'):       # "Специальность" / "Спецальность" (typo in source)
            meta['specialty_ru'] = line.split(':', 1)[-1].strip()
        elif low.startswith('квалификац'):
            meta['qualification_ru'] = line.split(':', 1)[-1].strip()
    
    return meta


def parse_excel_sheet(df: pd.DataFrame, sheet_name: str, start_row: int = 5) -> List[Dict[str, Any]]:
    """
    Парсит лист DataFrame (данные группы) и возвращает список студентов.
    
    Ожидаемая архитектура:
    - row 1 (индекс 1): названия предметов верхнего уровня (общеобразовательные предметы,
      а также заголовки модулей вида «Базалық модульдер», «Кәсіптік модуль 1» и т.д.)
    - row 2 (индекс 2): детальные названия предметов (БМ 1, ОН 1.1, ОН 1.2, ...)
      Если в row 2 есть значение — оно имеет приоритет (содержит точное название дисциплины).
    - row 3 (индекс 3): часы/кредиты '72с-3к'
    - row 4 (индекс 4): метки колонок (п, б, цэ, трад)
    - Начиная со start_row (index=5) - данные конкретных студентов.
    """
    students = []
    
    # Извлекаем мета-данные листа (специальность, квалификация, годы)
    sheet_meta = _parse_sheet_meta(df)
    
    # 1. Извлекаем предметы и часы (названия колонок)
    #    Row 1 — верхнеуровневые / общеобразовательные предметы
    #    Row 2 — детальные ОН / БМ (приоритетнее row 1)
    row_subjects = df.iloc[1]
    row_sub_subjects = df.iloc[2]
    
    # Динамический поиск строки с часами (Сағат саны)
    row_hours = df.iloc[3] # default
    for r in range(min(10, len(df))):
        val_c1 = str(df.iloc[r, 0]).strip().lower() if not pd.isna(df.iloc[r, 0]) else ''
        val_c2 = str(df.iloc[r, 1]).strip().lower() if not pd.isna(df.iloc[r, 1]) else ''
        if 'сағат' in val_c1 or 'часы' in val_c1 or 'сағат' in val_c2 or 'часы' in val_c2:
            row_hours = df.iloc[r]
            break
    
    # Кэш колонок: col_index -> { 'kz': .., 'ru': .., 'nkz':.., 'nru':.., 'hours':.., 'credits':.. }
    col_dict = {}
    ncols = max(len(row_subjects), len(row_sub_subjects))
    for col in range(2, ncols):
        # Сначала проверяем row 2 (детальное название предмета)
        cv_sub = row_sub_subjects.iloc[col] if col < len(row_sub_subjects) else None
        cv_main = row_subjects.iloc[col] if col < len(row_subjects) else None
        
        raw = None
        # Row 2 имеет приоритет (детальные ОН, БМ)
        if cv_sub is not None and not pd.isna(cv_sub):
            s = str(cv_sub).strip()
            # Пропускаем служебные строки вроде "Сабақтар"
            if s and s.lower() != 'nan' and 'Сабақтар' not in s and 'сағат' not in s.lower():
                raw = s
        
        # Если row 2 пустая, берём row 1 (общеобразовательные предметы)
        if raw is None:
            if cv_main is not None and not pd.isna(cv_main) and str(cv_main).strip() not in ('', 'nan'):
                raw = str(cv_main).strip()
        
        if raw is None:
            continue
            
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

    # Non-student row markers to skip (all lowercase, any language)
    NON_STUDENT_MARKERS = [
        'сағат саны', 'часы', 'итого', 'жиыны', 'барлығы',
        'руководитель', 'директор', 'заместитель', 'куратор',
        'жетекшісі', 'маманы', 'мамандығы'
    ]

    # Ищем индексы колонок с годами и номером диплома (ищем по имени листа)
    year_start_col_idx = -1
    year_end_col_idx = -1
    diploma_num_col_idx = -1
    
    sheet_name_low = sheet_name.lower()
    
    # Бухгалтеры: ET (149), EU (150), EV (151)
    if '3d' in sheet_name_low or '3д' in sheet_name_low:
        year_start_col_idx = 149
        year_end_col_idx = 150
        diploma_num_col_idx = 151
    # IT-шники: HH (215), HI (216), HJ (217)
    elif '3f' in sheet_name_low or '3ф' in sheet_name_low or '3ғ' in sheet_name_low:
        year_start_col_idx = 215
        year_end_col_idx = 216
        diploma_num_col_idx = 217
    else:
        # Fallback - динамический поиск если не нашли группу в названии листа
        row_headers = df.iloc[4]
        ncols = len(row_headers)
        for c in range(max(0, ncols - 10), ncols):
            hdr = str(row_headers.iloc[c]) if not pd.isna(row_headers.iloc[c]) else ''
            hdr_low = hdr.lower()
            if 'поступлен' in hdr_low:
                year_start_col_idx = c
            elif 'выпуск' in hdr_low:
                year_end_col_idx = c
            elif 'диплом' in hdr_low:
                diploma_num_col_idx = c

    # 2. Итерируемся по строкам студентов
    for i in range(start_row, len(df)):
        row = df.iloc[i]
        
        # Индекс и ФИО
        s_idx = safe_str(row.iloc[0])
        s_name = clean_name(row.iloc[1])

        # Skip empty / nan rows
        if not s_name or str(s_name).lower() in ('nan', ''):
            continue
        if str(s_idx).lower() in ('nan', ''):
            continue

        # Index must be numeric (1, 2, 3…) — non-numeric means header/footer row
        try:
            float(str(s_idx).replace(',', '.'))
        except ValueError:
            continue

        # Skip well-known non-student markers in name or index
        name_lower = str(s_name).lower()
        if any(marker in name_lower for marker in NON_STUDENT_MARKERS):
            continue

        # Stop at footer/signature rows
        if any(w in s_idx.lower() for w in ['руководитель', 'директор', 'заместитель']):
            break
            
        # Тема диплома
        diploma_kz = safe_str(row.iloc[-5]) if len(row) > 5 else ""
        diploma_ru = safe_str(row.iloc[-4]) if len(row) > 4 else ""
        
        # Год поступления, Год выпуска, Номер диплома
        # Безопасное чтение с проверкой границ
        year_start = str(row.iloc[year_start_col_idx]).strip() if year_start_col_idx != -1 and year_start_col_idx < len(row) and not pd.isna(row.iloc[year_start_col_idx]) else ""
        year_start = year_start.replace(".0", "") if year_start.endswith(".0") else year_start
        
        year_end = str(row.iloc[year_end_col_idx]).strip() if year_end_col_idx != -1 and year_end_col_idx < len(row) and not pd.isna(row.iloc[year_end_col_idx]) else ""
        year_end = year_end.replace(".0", "") if year_end.endswith(".0") else year_end
        
        diploma_num = str(row.iloc[diploma_num_col_idx]).strip() if diploma_num_col_idx != -1 and diploma_num_col_idx < len(row) and not pd.isna(row.iloc[diploma_num_col_idx]) else ""
        diploma_num = diploma_num.replace(" ", "")  # удаляем пробелы
        if diploma_num == "nan": diploma_num = ""

        # Собираем оценки и часы по всем предметам колонок
        grades = {}
        for c_idx, subj_info in col_dict.items():
            pts_val = row.iloc[c_idx]

            # Всегда сохраняем базовую инфу по часам и кредитам (важно для заголовков модулей)
            base_info = {
                'subject_kz': subj_info['kz'], 'subject_ru': subj_info['ru'],
                'hours': subj_info['hours'], 'credits': subj_info['credits'],
                'points': '', 'letter': '', 'gpa': '', 'traditional': '',
                'traditional_kz': '', 'traditional_ru': ''
            }

            if pd.isna(pts_val) or str(pts_val).strip() == '':
                grades[subj_info['nkz']] = base_info
                grades[subj_info['nru']] = base_info
                continue

            try:
                pts = float(str(pts_val).replace(',', '.'))
            except ValueError:
                # Если оценка "зачтено", "босатылды" и др. текстом
                val_str = str(pts_val).strip()
                if val_str:
                    # Если у предмета есть часы и кредиты, он требует числовой оценки.
                    # Случайный текст вроде "зачтено" игнорируем.
                    has_hours_credits = bool(subj_info.get('hours')) and bool(subj_info.get('credits'))
                    if has_hours_credits and val_str.lower() in ['зачтено', 'зачет', 'сынақ', 'сынак', 'өтті']:
                        val_str = ""
                    
                    if val_str:
                        base_info['traditional'] = val_str
                        grades[subj_info['nkz']] = base_info
                        grades[subj_info['nru']] = base_info
                continue
                
            from .utils import calc_letter_grade, calc_gpa_grade, calc_traditional_grade
            
            letter = calc_letter_grade(pts)
            gpa = f"{calc_gpa_grade(pts):.2f}"
            gpa = str(int(float(gpa))) if gpa.endswith('.00') else gpa
            
            base_info.update({
                'points': str(int(pts)) if pts.is_integer() else str(pts),
                'letter': letter,
                'gpa': gpa,
                'traditional_kz': calc_traditional_grade(pts, True),
                'traditional_ru': calc_traditional_grade(pts, False)
            })
            
            grades[subj_info['nkz']] = base_info
            grades[subj_info['nru']] = base_info

        students.append({
            'name': s_name,
            'sheet': sheet_name,
            'diploma_kz': diploma_kz,
            'diploma_ru': diploma_ru,
            'diploma_num': diploma_num,
            'year_start': year_start,
            'year_end': year_end,
            'grades': grades,
            'meta': sheet_meta,
        })
        
    return students
