import pandas as pd
import os

OUTPUT_DIR = "Diplomas_Batch"
FILE_NAME = "Аймахан Балауса Абайханқызы_KZ.xlsx"
FILE_PATH = os.path.join(OUTPUT_DIR, FILE_NAME)

def main():
    if not os.path.exists(FILE_PATH):
        print("File not found.")
        return

    print(f"Checking {FILE_PATH}...")
    try:
        # Read the first sheet (Page 1)
        df = pd.read_excel(FILE_PATH, sheet_name=0, header=None)
        
        # Check for subject names
        # Usually they are in some column.
        # Just dump non-empty cells.
        print("--- Content Sample (First 20 non-empty cells) ---")
        count = 0
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                val = df.iat[r, c]
                if pd.notna(val) and str(val).strip() != "":
                    print(f"({r},{c}): {str(val)[:50]}")
                    count += 1
                    if count > 20: break
            if count > 20: break
            
        # Check if subject count > 0 is implied by content
        
    except Exception as e:
        print(f"Error reading generated file: {e}")

if __name__ == "__main__":
    main()
