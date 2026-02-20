# -*- coding: utf-8 -*-
"""
check_attestation_data.py
Check attestation row in source file
"""

import pandas as pd

SOURCE = "2025-2026 диплом бағалары (ТОЛЫҚ) қызыл диплом жазылған соңғысы точно (1).xlsx"

for sheet_name in ["3Ғ-1"]:
    df = pd.read_excel(SOURCE, sheet_name=sheet_name, header=None)
    
    print(f"=== Sheet: {sheet_name} ===\n")
    
    COL_START = 2
    ROW_NAMES = 1
    ROW_HOURS = 3
    
    for col_idx in range(COL_START, df.shape[1], 4):
        name_r2 = df.iloc[ROW_NAMES, col_idx]
        name_r3 = df.iloc[ROW_NAMES+1, col_idx] if ROW_NAMES+1 < len(df) else None
        
        r2_str = str(name_r2).strip() if pd.notna(name_r2) else ""
        r3_str = str(name_r3).strip() if pd.notna(name_r3) else ""
        
        if "аттест" in r2_str.lower() or "аттест" in r3_str.lower() or "Итоговая" in r2_str or "Итоговая" in r3_str:
            print(f"Found at col {col_idx}:")
            print(f"  Row 2 (module): {r2_str}")
            print(f"  Row 3 (sub): {r3_str}")
            
            # Check row 4 (hours)
            hours_val = df.iloc[ROW_HOURS, col_idx]
            print(f"  Row 4 (hours): {hours_val}")
            
            # Check some student rows for actual data
            for row_idx in range(5, min(8, len(df))):
                val = df.iloc[row_idx, col_idx]
                if pd.notna(val) and str(val).strip():
                    print(f"    Row {row_idx+1}: {val}")
            
            print()
