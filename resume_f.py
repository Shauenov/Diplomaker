import os
import gc
import sys
import openpyxl
from configs import get_config
from src.parser import parse_excel_sheet
from src.generator import DiplomaGenerator

# Force utf-8 for Windows console
if sys.stdout.encoding.lower() != 'utf-8':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

def main():
    import pandas as pd
    
    print("Loading excel...")
    wb_source = openpyxl.load_workbook('local_test_copy.xlsx', read_only=True, data_only=True)
    
    # Process only the sheets that failed
    sheets_to_do = ["3Ғ-4"]
    for s_name in sheets_to_do:
        ws = wb_source[s_name]
        data = []
        for r in ws.iter_rows(max_row=200, max_col=200, values_only=True):
            data.append(r)
        
        df = pd.DataFrame(data)
        students = parse_excel_sheet(df, s_name, start_row=5)
        print(f"Loaded {len(students)} students from {s_name}")
        
        for lang in ['KZ', 'RU']:
            config, terms, template_name = get_config("3F", lang.lower())
            template_path = os.path.join("templates", template_name)
            
            for student in students:
                safe_name = "".join([c for c in student['name'] if c.isalpha() or c.isspace() or c in "-."]).strip()
                out_name = f"{s_name}_{safe_name}_{lang}.xlsx"
                out_path = os.path.join("output", out_name)
                
                # Skip if already exists
                if os.path.exists(out_path):
                    continue
                
                print(f"Generating {out_name}...")
                generator = DiplomaGenerator(template_path, out_path, config, terms)
                generator.fill_student_data(student)
                generator.close()
                del generator
                gc.collect()

if __name__ == "__main__":
    main()
