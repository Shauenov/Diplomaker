import pandas as pd
import os

TEMPLATE_PATH = "Diplom_IT_RU_Template.xlsx"
GENERATED_PATH = os.path.join("Diplomas_Batch", "Аймахан Балауса Абайханқызы_RU.xlsx")

def clean(val):
    if pd.isna(val): return ""
    return str(val).strip()

def main():
    if not os.path.exists(TEMPLATE_PATH):
        print(f"Template not found: {TEMPLATE_PATH}")
        return
    if not os.path.exists(GENERATED_PATH):
        print(f"Generated file not found: {GENERATED_PATH}")
        return

    print("Comparing Headers (Rows 0-15)...")
    
    try:
        df_tmpl = pd.read_excel(TEMPLATE_PATH, sheet_name=0, header=None)
        df_gen = pd.read_excel(GENERATED_PATH, sheet_name=0, header=None)
        
        # Compare first 15 rows
        for i in range(15):
            row_t = [clean(x) for x in df_tmpl.iloc[i].tolist()[:8]]
            row_g = [clean(x) for x in df_gen.iloc[i].tolist()[:8]]
            
            if row_t != row_g:
                print(f"Row {i} Mismatch:")
                print(f"  Template: {row_t}")
                print(f"  Generated: {row_g}")
            else:
                # print(f"Row {i} Match")
                pass
                
        print("\nComparing Sheet Names...")
        xl_t = pd.ExcelFile(TEMPLATE_PATH)
        xl_g = pd.ExcelFile(GENERATED_PATH)
        print(f"  Template: {xl_t.sheet_names}")
        print(f"  Generated: {xl_g.sheet_names}")
        
    except Exception as e:
        print(f"Error: {e}")

if __name__ == "__main__":
    main()
