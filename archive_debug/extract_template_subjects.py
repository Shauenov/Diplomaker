import pandas as pd

TEMPLATES = [
    "Diplom_IT_KZ_Template.xlsx",
    "Diplom_IT_RU_Template.xlsx"
]

def clean(val):
    if pd.isna(val): return ""
    return str(val).strip()

def main():
    for path in TEMPLATES:
        print(f"--- Analyzing {path} ---")
        try:
            xl = pd.ExcelFile(path)
            for sheet in xl.sheet_names:
                df = pd.read_excel(path, sheet_name=sheet, header=None)
                print(f"  Sheet: {sheet}")
                subjects = []
                # Assuming Subject is in Col 1 (Index 1) and starts around row 0-5
                # Let's scan Col 1
                for i in range(df.shape[0]):
                    val = df.iloc[i, 1] 
                    if not pd.isna(val) and str(val).strip() != "":
                        subjects.append(str(val).strip())
                
                print(f"    Found {len(subjects)} subjects:")
                for s in subjects[:3]: print(f"      - {s[:50]}...")
                if len(subjects) > 3: print(f"      ... and {len(subjects)-3} more")
                
        except Exception as e:
            print(f"Error reading {path}: {e}")
            
if __name__ == "__main__":
    main()
