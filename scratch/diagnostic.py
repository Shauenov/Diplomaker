
import sys
import os
import pandas as pd
from pathlib import Path

# Try to force UTF-8 output
if sys.stdout.encoding != 'utf-8':
    try:
        import io
        sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding='utf-8', errors='replace')
    except Exception:
        pass

# Add current dir to sys.path
sys.path.append(os.getcwd())

def safe_print(msg):
    try:
        print(msg)
    except Exception:
        try:
            print(msg.encode('ascii', errors='replace').decode('ascii'))
        except Exception:
            pass

safe_print("--- System Diagnostic ---")
safe_print(f"Python version: {sys.version}")
safe_print(f"Current directory: {os.getcwd()}")

# 1. Check dependencies
safe_print("\nChecking dependencies...")
try:
    import pandas
    import openpyxl
    import PySide6
    safe_print("[OK] All major dependencies (pandas, openpyxl, PySide6) are installed.")
except ImportError as e:
    safe_print(f"[FAIL] Missing dependency: {e}")

# 2. Check files
safe_print("\nChecking essential files...")
files_to_check = [
    "desktop_app.py",
    "main.py",
    "requirements.txt",
    "config/settings.py",
    "templates/Diplom_F_KZ_Template(4).xlsx",
    "templates/Diplom_F_RU_Template(4).xlsx",
    "templates/Diplom_D_KZ_Template(4).xlsx",
    "templates/Diplom_D_RU_Template(4).xlsx",
]

for f in files_to_check:
    if os.path.exists(f):
        safe_print(f"[OK] Found: {f}")
    else:
        safe_print(f"[FAIL] Not found: {f}")

# 3. Check Source Data
safe_print("\nChecking source data Excel files...")
source_candidates = [
    "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx",
    "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (5).xlsx",
    "2025-2026 official grades.xlsx"
]
found_sources = []
for s in source_candidates:
    if os.path.exists(s):
        safe_print(f"[OK] Found source candidate: {s}")
        found_sources.append(s)
    else:
        # Avoid printing non-existent filenames with special chars if they might fail
        pass

# 4. Run Preflight Check
if found_sources:
    safe_print("\nRunning Preflight Check on the latest source file...")
    try:
        from core.app_service import DiplomaGenerationService
        service = DiplomaGenerationService()
        
        # Prefer (5) if available
        actual_source = found_sources[0]
        for s in found_sources:
            if "(5)" in s:
                actual_source = s
                break
        
        safe_print(f"Testing with: {actual_source}")
        report = service.preflight_checks(
            source_file=actual_source,
            group="3F", # IT
            lang="ALL",
            output_dir="output_test"
        )
        
        if report["ok"]:
            safe_print("[OK] Preflight validation passed!")
            safe_print(f"     Sheets found: {report.get('sheet_count')}")
            safe_print(f"     Target sheets: {report.get('target_sheets')}")
        else:
            safe_print("[FAIL] Preflight validation failed.")
            for err in report.get("errors", []):
                safe_print(f"      - {err}")
    except Exception as e:
        safe_print(f"[ERROR] Exception during preflight: {e}")
        import traceback
        # traceback.print_exc() # Might fail on encoding too

safe_print("\n--- Diagnostic Complete ---")
