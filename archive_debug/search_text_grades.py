import pandas as pd

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"
SHEET_NAME = "3Ғ-1"

def search_text_all():
    xl = pd.ExcelFile(SOURCE_FILE)
    for sheet in xl.sheet_names:
        df = pd.read_excel(xl, sheet_name=sheet, header=None)
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                val = str(df.iloc[r, c])
                if "өте жақсы" in val or " жақсы" in val or "қанағаттанарлық" in val:
                    print(f"Sheet '{sheet}' Row {r} Col {c} has grade: {val}")
                    # Print first 2 cols of this row
                    print(f"  Row {r} info: {df.iloc[r, :2].tolist()}")
                    return # Stop after first sheet found to avoid spam

if __name__ == "__main__":
    search_text_all()
