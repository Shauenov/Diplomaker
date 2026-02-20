
import re

def parse_dump(filename):
    with open(filename, 'r', encoding='utf-8') as f:
        lines = [line.strip() for line in f if line.strip()]

    subjects = []
    current_subject = None
    
    # Simple state machine or regex search
    # Pattern: Digit. Subject Name
    # Then Hours
    # Then Credits
    
    # Regex for start of subject: "1. Қазақ тілі" or just "Қазақ тілі" if numbers are separate
    # Dump has lines like: "1" then "Қазақ тілі"
    
    i = 0
    while i < len(lines):
        line = lines[i]
        
        # Check if line is a number (index) or subject start
        # Many subjects start with upper case Cyrillic
        # Hours are usually integers like 72, 96, 120...
        # Credits are small integers 3, 4, 5...
        
        # It's hard to be perfect, but let's try to grab chunks
        
        # Heuristic: line with Kazakh letters is a subject
        # line with digits is hours/credits
        
        if re.match(r'^[0-9]+$', line) and int(line) < 1000 and int(line) > 10: 
            # Likely hours (e.g. 72, 96)
            hours = line
            # Next line might be credits
            if i+1 < len(lines) and re.match(r'^[0-9]+(\.[0-9]+)?$', lines[i+1]):
                credits = lines[i+1]
                # The PREVIOUS line(s) were likely the subject
                # Let's verify
                if i > 0:
                    subj = lines[i-1]
                    if re.match(r'^[0-9]+$', subj): # If previous is just a number (index)
                        if i > 1: subj = lines[i-2]
                    
                    if len(subj) > 3:
                        subjects.append((subj, hours, credits))
            i += 1
        elif re.match(r'^[0-9]+$', line) and int(line) < 100:
             # Might be just index
             pass
        
        i += 1
        
    # Write to file directly
    with open("extracted_it_list.py", "w", encoding="utf-8") as out:
        out.write("ALL_SUBJECTS = [\n")
        for s in subjects:
            out.write(f'    ("{s[0]}", {s[1]}, {s[2]}),\n')
        out.write("]\n")

if __name__ == "__main__":
    try:
        parse_dump("kaz_it_dump_v4.txt")
        print("Done.")
    except Exception as e:
        print(f"Error: {e}")
