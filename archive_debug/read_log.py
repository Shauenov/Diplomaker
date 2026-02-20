try:
    print("--- FILTERED LOG ---")
    with open("debug_log_utf8_2.txt", "r", encoding="utf-16le", errors="replace") as f: # PS default
         for line in f:
             if "DEBUG" in line or "Ф1" in line:
                 print(line.strip())
except Exception as e:
    print(f"Error: {e}")
