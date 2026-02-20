"""Verify BM subjects in generated diplomas have hours/credits and are not bold."""
import zipfile, xml.etree.ElementTree as ET, re

NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'

def get_shared_strings(z):
    with z.open('xl/sharedStrings.xml') as f:
        root = ET.parse(f).getroot()
        return [''.join(t.text or '' for t in si.iter(f'{{{NS}}}t'))
                for si in root.findall(f'{{{NS}}}si')]

for lang, fname in [
    ('KZ', 'Diplomas_Batch/3F-1_Аймахан Балауса Абайханқызы_KZ.xlsx'),
    ('RU', 'Diplomas_Batch/3F-1_Аймахан Балауса Абайханқызы_RU.xlsx'),
]:
    with zipfile.ZipFile(fname) as z:
        strings = get_shared_strings(z)
    
    bm_entries = [(i, s) for i, s in enumerate(strings) if s.startswith('БМ')]
    print(f'\n=== {lang} Diploma ===')
    for idx, name in bm_entries:
        # Hours and credits should be near the BM entry in strings
        nearby = strings[idx:idx+5]
        print(f'  "{name[:60]}"')
        print(f'    -> Nearby strings: {nearby[1:4]}')
