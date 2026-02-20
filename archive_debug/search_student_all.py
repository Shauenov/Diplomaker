import pandas as pd

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"

def search_student_all():
    xl = pd.ExcelFile(SOURCE_FILE)
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        for r in range(df.shape[0]):
            val = str(df.iloc[r, 1])
            if "Төленді Қалжан" in val or "Болат Хамида" in val or "Бакирам Балнұр" in val:
                # Count non-empty technical columns (e.g. 50-150)
                non_empty = 0
                for c in range(2, min(200, df.shape[1])):
                    v = df.iloc[r, c]
                    if pd.notna(v) and str(v).strip() not in ["", "0", "0.0"]:
                        non_empty += 1
                if non_empty > 5:
                    print(f"--- SUCCESS: Found {val} in Sheet '{sheet}' Row {r} with {non_empty} grades ---")
                else:
                    print(f"--- EMPTY: Found {val} in Sheet '{sheet}' Row {r} (empty) ---")

if __name__ == "__main__":
    search_student_all()
