import openpyxl
import re
import os

def parse_hc(h_str):
    if not h_str: return 0, 0
    h_str = str(h_str).lower().replace(' ', '')
    # '72с-3к'
    m = re.match(r'(\d+)с.*?(\d+(?:[.,]\d+)?)к', h_str)
    if m:
        c = float(m.group(2).replace(',', '.'))
        return int(m.group(1)), (int(c) if c.is_integer() else c)
    return 0, 0

def main():
    wb = openpyxl.load_workbook('2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (2).xlsx', read_only=True, data_only=True)
    with open('src/columns_config.py', 'w', encoding='utf-8') as out:
        out.write('SUBJECT_COLUMNS = {\n')
        
        for sheet in ['3D-1', '3Ғ-1']:
            group_key = '3D' if '3D' in sheet else '3F'
            out.write(f'    "{group_key}": {{\n')
            ws = wb[sheet]
            rows = list(ws.iter_rows(min_row=1, max_row=6, values_only=True))
            if len(rows) < 4: continue
            
            r2, r3 = rows[1], rows[2]
            
            # Dynamically find the hours row — search for "сағат"/"часы" in first 6 rows
            # This handles the case where 3D and 3F may have hours in different rows
            r4 = rows[3]  # default fallback
            for row_data in rows:
                c0 = str(row_data[0]).lower() if row_data and row_data[0] else ''
                c1 = str(row_data[1]).lower() if row_data and len(row_data) > 1 and row_data[1] else ''
                if 'сағат' in c0 or 'часы' in c0 or 'сағат' in c1 or 'часы' in c1:
                    r4 = row_data
                    break
            
            for i in range(2, len(r2)):
                cv_sub = r3[i] if i < len(r3) else None
                cv_main = r2[i] if i < len(r2) else None
                
                raw = None
                if cv_sub is not None and str(cv_sub).strip() != '':
                    s = str(cv_sub).strip()
                    if s.lower() != 'none' and 'Сабақтар' not in s and 'сағат' not in s.lower():
                        raw = s
                
                if raw is None:
                    if cv_main is not None and str(cv_main).strip() not in ('', 'None'):
                        raw = str(cv_main).strip()
                
                if raw is None or raw == 'None':
                    continue
                
                parts = raw.split('\n')
                kz_name = parts[0].strip().rstrip(':')
                ru_name = parts[1].strip().rstrip(':') if len(parts) >= 2 else kz_name
                
                h_str = str(r4[i]).strip() if i < len(r4) and r4[i] is not None else ""
                hours, credits = parse_hc(h_str)
                
                # Format strings safely
                kz_name = kz_name.replace('"', '\\"')
                ru_name = ru_name.replace('"', '\\"')
                
                out.write(f'        {i}: {{"kz": "{kz_name}", "ru": "{ru_name}", "hours": {hours}, "credits": {credits}}},\n')
            
            out.write('    },\n')
        out.write('}\n\n')
        
        out.write('META_COLUMNS = {\n')
        out.write('    "3D": {"year_start": 149, "year_end": 150, "diploma_num": 151},\n')
        out.write('    "3F": {"year_start": 215, "year_end": 216, "diploma_num": 217},\n')
        out.write('}\n')
    print("Columns config successfully generated in src/columns_config.py")

if __name__ == '__main__':
    main()
