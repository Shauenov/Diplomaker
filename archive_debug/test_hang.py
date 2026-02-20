print("1. Starting imports")
import pandas as pd
print("2. Imports done")

source_file = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
print(f"3. Opening {source_file}")
xl = pd.ExcelFile(source_file)
print("4. ExcelFile open")
sheets = xl.sheet_names
print("5. Sheets:", sheets)
target_sheets = [s for s in xl.sheet_names if s.startswith('3Ғ')]

for s in target_sheets:
    print("6. Parsing sheet", s)
    df = xl.parse(s, header=None)
    print("7. Parsed shape:", df.shape)
    
    # Check while loop
    row_subjects = df.iloc[1]
    col = 2
    while col < len(row_subjects):
        col += 4
    print("8. While loop done")
    break

print("Done")
