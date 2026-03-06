import sys
import io
from src.parser import parse_excel_sheet
from src.utils import normalize_key
import pandas as pd

def main():
    excel_path = r"c:\Users\user\OneDrive\Рабочий стол\template\2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
    xl = pd.ExcelFile(excel_path)
    
    with open("debug_keys_out.txt", "w", encoding="utf-8") as f:
        f.write(f"Sheets available: {xl.sheet_names}\n")
        
        # 3Ғ-1 (IT каз)
        df_kz = pd.read_excel(excel_path, sheet_name='3Ғ-1', header=None)
        students_kz = parse_excel_sheet(df_kz, '3Ғ-1', start_row=5)
        if students_kz:
            st = students_kz[0]
            f.write(f"\n3Ғ-1 - Keys in grades:\n")
            for k, v in st['grades'].items():
                if '10.2' in k:
                    f.write(f"  {k} -> {v}\n")
                    
        # 3D-2 (BU rus)
        df_ru = pd.read_excel(excel_path, sheet_name='3D-2', header=None)
        students_ru = parse_excel_sheet(df_ru, '3D-2', start_row=5)
        if students_ru:
            st = students_ru[0]
            f.write(f"\n3D-2 - Keys in grades:\n")
            keys = list(st['grades'].keys())
            for k in keys:
                if 'ро1.3' in k or 'пм1' in k:
                    f.write(f"  {k} -> {st['grades'][k]['hours']} hours\n")
            # Also log all keys for 3D-2 just to check naming
            f.write("\nAll 3D-2 normalized keys:\n")
            for k in keys:
                f.write(f"  {k}\n")
               
if __name__ == '__main__':
    main()
