import pandas as pd
import os

files = [
    "Diplom_IT_KZ_Template.xlsx", # Input template (desired structure)
    "Diplom_IT_RU_Template.xlsx"
]

base_path = r"c:\Users\user\OneDrive\Рабочий стол\template"

with open("template_structure.txt", "w", encoding="utf-8") as outfile:
    for f in files:
        path = os.path.join(base_path, f)
        outfile.write(f"\n=== {f} ===\n")
        try:
            xls = pd.ExcelFile(path)
            for sheet_name in xls.sheet_names:
                outfile.write(f"\n--- Sheet: {sheet_name} ---\n")
                df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                
                # Inspect column B (index 1)
                if df.shape[1] > 1:
                    subjects = df[1].dropna()
                    outfile.write(f"Count: {len(subjects)}\n")
                    for idx, val in subjects.items():
                         outfile.write(f"Row {idx+1}: {val}\n")
                else:
                    outfile.write("Sheet has fewer than 2 columns.\n")

        except Exception as e:
            outfile.write(f"Error: {e}\n")
print("Done writing to template_structure.txt")
