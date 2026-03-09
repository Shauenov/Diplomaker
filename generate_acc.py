import json

with open('acc_subjects.json', 'r', encoding='utf-8') as f:
    data = json.load(f)

res_pages = {1: [], 2: [], 3: [], 4: []}

for i in range(17):
    s = data[i]
    is_mod = 'True' if str(s['code']).startswith('КМ') or str(s['code']).startswith('БМ') else 'False'
    res_pages[1].append(f'''                Subject(name_kz="{s['name_kz']}", name_ru="{s['name_ru']}", hours="{s['hours']}", credits="{s['credits']}", is_module_header={is_mod}, is_elective=False)''')

for i in range(17, 29):
    s = data[i]
    is_mod = 'True' if str(s['code']).startswith('КМ') or str(s['code']).startswith('БМ') else 'False'
    res_pages[2].append(f'''                Subject(name_kz="{s['name_kz']}", name_ru="{s['name_ru']}", hours="{s['hours']}", credits="{s['credits']}", is_module_header={is_mod}, is_elective=False)''')

for i in range(29, 40):
    s = data[i]
    is_mod = 'True' if str(s['code']).startswith('КМ') or str(s['code']).startswith('БМ') else 'False'
    res_pages[3].append(f'''                Subject(name_kz="{s['name_kz']}", name_ru="{s['name_ru']}", hours="{s['hours']}", credits="{s['credits']}", is_module_header={is_mod}, is_elective=False)''')

for i in range(40, 43):
    s = data[i]
    is_mod = 'True' if str(s['code']).startswith('КМ') or str(s['code']).startswith('БМ') else 'False'
    res_pages[4].append(f'''                Subject(name_kz="{s['name_kz']}", name_ru="{s['name_ru']}", hours="{s['hours']}", credits="{s['credits']}", is_module_header={is_mod}, is_elective=False)''')


electives = [
    ('Қорытынды аттестаттау :', 'Итоговая аттестация :'),
    ('Ф1 Факультативтік ағылшын тілі', 'Факультатив английский язык'),
    ('Ф2 Факультативтік түрік тілі', 'Факультатив турецкий язык'),
    ('Ф3 Факультативтік Бизнес және бухгалтерлік есептегі жағдайлар (Cases in Business and Accounting)', 'Ф3 Факультатив Ситуации в бизнесе и бухгалтерском учете (Cases in Business and Accounting)'),
    ('Ф4 Факультативтік Бизнес деректерін талдау (Business data analysis (excel, macros, google sheets, sql, python, power BI, tableau))', 'Ф4 Факультатив Анализ бизнес данных (Business data analysis (excel, macros, google sheets, sql, python, power BI, tableau))'),
    ('Ф5 Факультативтік кәсіпкерлік қызмет негіздері (Enterpreneurship)', 'Ф5 Факультатив основы предпринимательской деятельности (Enterpreneurship)')
]

for kz, ru in electives:
    is_elec = 'True' if kz.startswith('Ф') else 'False'
    res_pages[4].append(f'''                Subject(name_kz="{kz}", name_ru="{ru}", hours="", credits="", is_module_header=False, is_elective={is_elec})''')


out = 'PROGRAM_ACCOUNTING_PAGES = {\n'
for p in range(1, 5):
    out += f'    {p}: [\n' + ',\n'.join(res_pages[p]) + '\n    ],\n'
out += '}\n'

with open('acc_pages_config.py', 'w', encoding='utf-8') as f:
    f.write(out)
