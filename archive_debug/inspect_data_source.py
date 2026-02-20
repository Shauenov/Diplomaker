
import pandas as pd
import os

# The specific file mentioned by the user
file_path = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
abs_path = os.path.abspath(file_path)

if not os.path.exists(abs_path):
    print(f"Error: File not found at {abs_path}")
    exit(1)

print(f"Reading: {abs_path}")

try:
    with open("data_dump.txt", "w", encoding="utf-8") as f:
        xl = pd.ExcelFile(abs_path)
        f.write(f"Sheets: {xl.sheet_names}\n")
        
        # Read first sheet
        df_raw = pd.read_excel(abs_path, sheet_name=0, header=None, nrows=50)
        
        header_row_idx = -1
        for i, row in df_raw.iterrows():
            row_str = " ".join([str(x) for x in row.values if pd.notna(x)])
            # Check for ANY common header keywords
            if any(k in row_str for k in ["Аты-жөні", "Ф.И.О", "Фамилия", "Student", "№"]):
                header_row_idx = i
                f.write(f"\n--- Header found at row {header_row_idx} ---\n")
                f.write(f"Row {i} content: {row_str}\n")
                break
                
        if header_row_idx != -1:
            # Reload with correct header
            df = pd.read_excel(abs_path, sheet_name=0, header=header_row_idx)
            f.write(f"Columns: {df.columns.tolist()}\n")
            
            # Print rows 14-16 (adjusting for header index if needed)
            # Row 9 in 0-indexed is the subject header.
            # Row 14 is likely the first student.
            # Let's verify by printing a slice around where we expect data.
            
            # Since we loaded with header=header_row_idx (which was 0), 
            # the index in df corresponds to Excel Row - 1 (header).
            # If Excel Row 14 is the first student, that is df index 13 ??
            # Wait, header_row_idx was 0.
            # So df index 14 is Row 15. 
            
            f.write("\n--- Data Rows 14-18 ---\n")
            f.write(df.iloc[14:19].to_string())
        else:
            f.write("\n--- Header NOT found in first 50 rows. Printing raw top 10 ---\n")
            f.write(df_raw.head(10).to_string())
            
        f.write(f"\n\n=== SHEET: 3Ғ-1 ===\n")
        try:
            df_gh = pd.read_excel(abs_path, sheet_name="3Ғ-1", header=None, nrows=20)
            f.write(df_gh.to_string())
        except Exception as e:
            f.write(f"Error reading 3Ғ-1: {e}")
            
    print("Dump written to data_dump.txt")

    
except Exception as e:
    print(f"Error reading Excel: {e}")
