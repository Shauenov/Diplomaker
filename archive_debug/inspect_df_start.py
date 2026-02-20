import pandas as pd

SOURCE_FILE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"

def inspect():
    df = pd.read_excel(SOURCE_FILE, header=3)
    print("Searching for Аймахан...")
    found = df[df.iloc[:, 1].astype(str).str.contains("Аймахан", na=False)]
    print(found)

if __name__ == "__main__":
    inspect()
