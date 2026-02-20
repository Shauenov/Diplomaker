import pandas as pd

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def main():
    print(f"Reading {SOURCE_FILE}...")
    try:
        df = pd.read_excel(SOURCE_FILE, sheet_name=SHEET_NAME, header=None)
    except FileNotFoundError:
        print("File not found.")
        return

    print("Searching for 'Қазақ тілі'...")
    found = False
    for r_idx in range(df.shape[0]):
        for c_idx in range(df.shape[1]):
            val = df.iat[r_idx, c_idx]
            if isinstance(val, str) and "Қазақ тілі" in val:
                print(f"FOUND 'Қазақ тілі' at Row {r_idx}, Col {c_idx}")
                print(f"Full content: {val}")
                found = True
                # Print adjacent cells to confirm structure
                print(f"Row {r_idx} First 10 cols: {df.iloc[r_idx, :10].tolist()}")
                
                # Check 2 rows below for Hours
                if r_idx + 2 < df.shape[0]:
                    print(f"Row {r_idx+2} (Possible Hours) Col {c_idx}: {df.iat[r_idx+2, c_idx]}")
                
                break
        if found: break

    if not found:
        print("Subject NOT found in first scan. Trying 'Kazakh' or other keywords?")

if __name__ == "__main__":
    main()
