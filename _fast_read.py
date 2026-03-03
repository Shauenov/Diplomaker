import openpyxl
import time

t0 = time.time()
wb = openpyxl.load_workbook("2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx", read_only=True, data_only=True)
print("Sheets:", wb.sheetnames)
print(f"Time: {time.time() - t0:.2f}s")
