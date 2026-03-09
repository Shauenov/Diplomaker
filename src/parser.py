import pandas as pd
from typing import Dict, Any, List
from .utils import normalize_key, safe_str, clean_name
from .columns_config import SUBJECT_COLUMNS, META_COLUMNS
from core.converters import convert_score_to_grade
from config.settings import META_FIELD_COLUMNS


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
    
    # 1. Определяем группу по имени листа и загружаем хардкод
    sheet_name_low = sheet_name.lower()
    if '3d' in sheet_name_low or '3д' in sheet_name_low:
        group_key = '3D'
    elif '3f' in sheet_name_low or '3ф' in sheet_name_low or '3ғ' in sheet_name_low:
        group_key = '3F'
    else:
        group_key = None

    # Кэш колонок из хардкод-конфига
    col_dict = {}
    if group_key and group_key in SUBJECT_COLUMNS:
        for col_idx, info in SUBJECT_COLUMNS[group_key].items():
            kz_name = info['kz']
            ru_name = info['ru'] if info['ru'] else kz_name
            col_dict[col_idx] = {
                'kz': kz_name, 'ru': ru_name,
                'nkz': normalize_key(kz_name), 'nru': normalize_key(ru_name),
                'hours': info['hours'], 'credits': info['credits']
            }

    # Non-student row markers to skip
    NON_STUDENT_MARKERS = [
        'сағат саны', 'часы', 'итого', 'жиыны', 'барлығы',
        'руководитель', 'директор', 'заместитель', 'куратор',
        'жетекшісі', 'маманы', 'мамандығы'
    ]

    # Хардкод индексов для года поступления, выпуска и номера диплома
    # Primary source: centralized config/settings.py
    # Fallback: legacy src/columns_config.py mapping
    meta_cols = {}
    if group_key:
        meta_cols = META_FIELD_COLUMNS.get(group_key, {})
        if not meta_cols:
            meta_cols = META_COLUMNS.get(group_key, {})
    year_start_col_idx = meta_cols.get('year_start', -1)
    year_end_col_idx = meta_cols.get('year_end', -1)
    diploma_num_col_idx = meta_cols.get('diploma_num', -1)

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
        sorted_keys = sorted(col_dict.keys())
        
        for i, c_idx in enumerate(sorted_keys):
            subj_info = col_dict[c_idx]
            
            # Find max offset (distance to next subject)
            max_offset = sorted_keys[i + 1] - c_idx if i < len(sorted_keys) - 1 else 4
            
            pts_val = None
            for offset in range(max_offset):
                chk_idx = c_idx + offset
                if chk_idx < len(row):
                    val = row.iloc[chk_idx]
                    if pd.notna(val) and str(val).strip() != '':
                        pts_val = val
                        break

            # Всегда сохраняем базовую инфу по часам и кредитам (важно для заголовков модулей)
            base_info = {
                'subject_kz': subj_info['kz'], 'subject_ru': subj_info['ru'],
                'hours': subj_info['hours'], 'credits': subj_info['credits'],
                'points': '', 'letter': '', 'gpa': '', 'traditional': '',
                'traditional_kz': '', 'traditional_ru': ''
            }

            if pts_val is None or str(pts_val).strip() == '':
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
            
            # Централизованная конвертация оценок через config/settings.GRADE_THRESHOLDS
            grade_obj = convert_score_to_grade(str(pts))
            gpa_str = f"{grade_obj.gpa:.2f}" if grade_obj.gpa is not None else ""
            gpa_str = str(int(float(gpa_str))) if gpa_str.endswith('.00') else gpa_str
            
            base_info.update({
                'points': str(int(pts)) if pts.is_integer() else str(pts),
                'letter': grade_obj.letter,
                'gpa': gpa_str,
                'traditional_kz': grade_obj.traditional_kz,
                'traditional_ru': grade_obj.traditional_ru
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
