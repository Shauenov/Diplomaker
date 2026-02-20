import zipfile
import xml.etree.ElementTree as ET
import sys
import os

def extract_text(docx_path, output_txt):
    print(f"Extracting from: {docx_path}")
    if not os.path.exists(docx_path):
        print("Error: File not found!")
        return

    try:
        with zipfile.ZipFile(docx_path) as z:
            xml_content = z.read('word/document.xml')
            tree = ET.fromstring(xml_content)
            
            # Namespace for Word
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            with open(output_txt, "w", encoding="utf-8") as f:
                # Find all text nodes
                for t in tree.iterfind('.//w:t', ns):
                    if t.text:
                        f.write(t.text + "\n")
        print(f"Successfully extracted to {output_txt}")
                        
    except Exception as e:
        print(f"Error extracting text: {e}")

if __name__ == "__main__":
    docx_file = "3Ф ПРИЛОЖ 2025 РУС ШАБЛОН соңғы.docx"
    output_file = "rus_it_text.txt"
    extract_text(docx_file, output_file)
