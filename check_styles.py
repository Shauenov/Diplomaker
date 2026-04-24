"""
Диагностика: что реально записано в style_index 45, 46, 48 в styles.xml
"""
import zipfile, xml.etree.ElementTree as ET, glob, os, sys

files = sorted(glob.glob("output/*KZ*.xlsx"))
if not files:
    print("Нет KZ файлов в output/"); sys.exit(1)

target = files[0]
print(f"Проверяем: {os.path.basename(target)}\n")

NS = {"main": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}

with zipfile.ZipFile(target) as z:
    styles_xml = z.read("xl/styles.xml")
    styles_root = ET.fromstring(styles_xml)
    xfs = styles_root.findall(".//main:xf", NS)

for idx in [45, 46, 47, 48]:
    if idx < len(xfs):
        xf = xfs[idx]
        align = xf.find("main:alignment", NS)
        print(f"style xf[{idx}]: {xf.attrib}")
        print(f"  alignment: {align.attrib if align is not None else 'НЕТУ'}")
        print()
