
import zipfile
import re
import xml.etree.ElementTree as ET

def extract_text(docx_path):
    with zipfile.ZipFile(docx_path) as z:
        xml_content = z.read('word/document.xml')
    
    root = ET.fromstring(xml_content)
    # Namespaces
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    text_content = []
    
    with open("kaz_it_dump_v4.txt", "w", encoding="utf-8") as f:
        # Iterate through all text nodes
        for node in root.findall('.//w:t', ns):
            if node.text and node.text.strip():
                f.write(node.text.strip() + "\n")

if __name__ == "__main__":
    try:
        extract_text("kaz_it.docx")
        print("Done.")
    except Exception as e:
        print(f"Error: {e}")
