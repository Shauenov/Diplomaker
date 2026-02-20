"""Quick verification: read one generated diploma and check headers."""
import zipfile
import xml.etree.ElementTree as ET

# Read via zipfile (xlsx is a zip)
fname = "Diplomas_Batch/3F-1_Аймахан Балауса Абайханқызы_KZ.xlsx"

with zipfile.ZipFile(fname, 'r') as z:
    # List sheets
    with z.open('xl/workbook.xml') as f:
        tree = ET.parse(f)
        root = tree.getroot()
        ns = {'x': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
        sheets = root.findall('.//x:sheet', ns)
        print("Sheets:", [s.attrib.get('name') for s in sheets])
    
    # Read shared strings
    with z.open('xl/sharedStrings.xml') as f:
        ss_tree = ET.parse(f)
        ss_root = ss_tree.getroot()
        strings = []
        for si in ss_root.findall('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}si'):
            text = ''.join(t.text or '' for t in si.iter('{http://schemas.openxmlformats.org/spreadsheetml/2006/main}t'))
            strings.append(text)
        print(f"\nShared strings (first 30):")
        for i, s in enumerate(strings[:30]):
            print(f"  [{i}]: {s[:80]}")
