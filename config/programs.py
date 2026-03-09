# -*- coding: utf-8 -*-
"""
Program Configuration
====================
Program-specific definitions (IT 3F, Accounting 3D, etc.)

Each program has:
- Sheet names in source Excel
- Subject structure (subjects per page, names)
- Column mappings
- Custom settings
"""

# ─────────────────────────────────────────────────────────────
from core.models import Subject

# IT PROGRAM (F-group / 3F)
# ─────────────────────────────────────────────────────────────

PROGRAM_IT_PAGES = {
    1: [
                Subject(name_kz="Қазақ тілі", name_ru="Казахский язык", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="Қазақ әдебиеті", name_ru="Казахская литература", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="Орыс тілі және әдебиеті", name_ru="Русский язык и литература", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="Ағылшын тілі", name_ru="Английский язык", hours="216", credits="9", is_module_header=False, is_elective=False),
                Subject(name_kz="Қазақстан тарихы", name_ru="История Казахстана", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="Математика", name_ru="Математика", hours="168", credits="7", is_module_header=False, is_elective=False),
                Subject(name_kz="Информатика", name_ru="Информатика", hours="48", credits="2", is_module_header=False, is_elective=False),
                Subject(name_kz="Алғашқы әскери және технологиялық дайындық", name_ru="Начальная военная и технологическая подготовка", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="Дене тәрбиесі", name_ru="Физическая культура", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="Физика", name_ru="Физика", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="Химия", name_ru="Химия", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="Биология", name_ru="Биология", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="Дүниежүзі тарихы", name_ru="Всемирная история", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="БМ 1 Дене қасиеттерін дамыту және жетілдіру", name_ru="БМ 1. Развитие и совершенствование физических качеств", hours="216", credits="9", is_module_header=False, is_elective=False),
                Subject(name_kz="БМ 2 Ақпараттық-коммуникациялық және цифрлық технологияларды қолдану", name_ru="БМ 2. Применение информационно-коммуникационных и цифровых технологий", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="БМ 3 Экономиканың базалық білімін және кәсіпкерлік негіздерін қолдану", name_ru="БМ 3. Применение базовых знаний экономики и основ предпринимательства", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="БМ 4 Қоғам мен еңбек ұжымында әлеуметтену және бейімделу үшін әлеуметтік ғылымдар негіздерін қолдану", name_ru="БМ 4. Применение основ социальных наук для социализации и адаптации в обществе и трудовом коллективе", hours="24", credits="1", is_module_header=False, is_elective=False),
    ],
    2: [
                Subject(name_kz="КМ 1 Кәсіптік қызмет саласында коммуникация үшін ауызша және жазбаша қарым-қатынас дағдыларын пайдалану", name_ru="ПМ 1 Использование устных и письменных навыков общения в профессиональной сфере", hours="264", credits="11", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 1.1 Кәсіптік салада қарым-қатынас жасауға қажетті ағылшын тілінің лексика-грамматикалық материалдарын меңгеру.", name_ru="РО 1.1 Освоение лексико-грамматического материала английского языка для профессионального общения", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 1.2 Кәсіптік бағыттағы мәтіндерді оқу және аудару.", name_ru="РО 1.2 Чтение и перевод профессионально ориентированных текстов", hours="48", credits="2", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 1.3 Ойды жеткізудегі негізгі сөйлеу формаларын қалыптастыру.", name_ru="РО 1.3 Формирование ключевых речевых форм для выражения мысли", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 1.4 Іскерлік мақсатта екінші шетел тілін қолдану", name_ru="РО 1.4 Использование второго иностранного языка в деловом общении", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="КМ 2 Кәсіби қызметте математикалық, статистикалық есептерді қолдану", name_ru="ПМ 2 Применение математических и статистических расчетов в профессиональной деятельности", hours="144", credits="6", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 2.1 Тізбекті алгебра негізгі түсінігі мен әдістерін қолдану.", name_ru="РО2.1 Применение основных понятий и методов алгебры последовательностей.", hours="48", credits="2", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 2.2 Аналитикалық геометрия негізгі түсінігі мен әдістерін қолдану.", name_ru="РО 2.2 Использование базовых знаний аналитической геометрии", hours="48", credits="2", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 2.3 Математикалық талдаудың негіздерін қалыптастыру.", name_ru="РО 2.3 Освоение основ математического анализа", hours="48", credits="2", is_module_header=False, is_elective=False),
                Subject(name_kz="КМ 3 Front-end web ресурстарды құру", name_ru="ПМ 3 Разработка web-ресурсов Front-end", hours="408", credits="17", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 3.1 Контентті басқару жүйесі үшін жеке шаблондар мен плагиндер жасау.", name_ru="РО 3.1 Создание шаблонов и плагинов для систем управления контентом", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 3.2 Веб-сайттың көрінісін өзгерту үшін CSS немесе басқа сыртқы файлдарды пайдалану", name_ru="РО 3.2 Изменение дизайна веб-сайтов с использованием CSS и других внешних файлов", hours="168", credits="7", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 3.3 Пайдаланушылар үшін web-сайттарды құру, жаңарту және іздеу жүйесін құрастыру.", name_ru="РО 3.3 Разработка, обновление и настройка веб-сайтов для пользователей", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="КМ 04 Графикалық дизайн жасау", name_ru="ПМ 04 Графический дизайн", hours="288", credits="12", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 4.1 Өнеркәсіптік дизайнның жан-жағынан жинақталған бұйымдардың жобасын жасау.", name_ru="РО 4.1 Проектирование промышленных изделий с учетом комплексного дизайна", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 4.2 Дизайн макеттерін дайындау", name_ru="РО 4.2 Создание макетов дизайна", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 4.3 Компьютерде растрлық және векторлық бейнелерді құрастыру.", name_ru="РО 4.3 Разработка растровых и векторных изображений на компьютере", hours="96", credits="4", is_module_header=False, is_elective=False),
    ],
    3: [
                Subject(name_kz="КМ 05 Алгоритмге кіріспе түрлерін пайдаланып, бағдарламалар жасау", name_ru="ПМ 05 Программирование и разработка алгоритмов", hours="288", credits="12", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 5.1 Клиент-сервер негізінде бағдарламалық шешімдердің кодтарын құрастыру.", name_ru="РО 5.1 Разработка клиент-серверных программных решений.", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 5.2 Кодтарды өзгерту үшін соңғы бағдарламалық жасақтаманы әзірлеу орталары мен құралдарын пайдалану.", name_ru="РО 5.2 Использование новейших сред разработки программного обеспечения и инструментов для модификации кодов.", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 5.3 Жүйені дамыту әдістемелерін пайдалану.", name_ru="РО 5.3 Применение методов развития программных систем", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="КМ 06 ІТ Менеджмент стандарттарымен экологиялық және қауіпсіздік шаралармен жобаларды басқару", name_ru="ПМ 06 Стандарты управления ИТ, меры по охране окружающей среды и безопасности управление проектом", hours="120", credits="5", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 6.1 Пайдаланушы сұхбат, сауалнама, құжаттарды іздеу және талдау, бірлескен бағдарламаны әзірлеу және бақылау талаптарын өңдеу.", name_ru="РО 6.1 Интервью с пользователями, опросы, поиск и анализ документов, совместная разработка программ и отслеживание обработки требований.", hours="24", credits="1", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 6.2 Шешім қабылдау үшін баламаларды әзірлеу, ең қолайлы баламаны таңдау және қажетті шешімді жасау.", name_ru="РО 6.2 Разработка альтернатив для принятия решений, выбор наиболее подходящей альтернативы и принятие необходимого решения.", hours="24", credits="1", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 6.3 Бағдарламалық қамтамасыздандырудың мақсат пен міндет қоюды жүзеге асыру және қойылатын талаптарды әзірлеу", name_ru="РО 6.3 Реализация целей и задач программного обеспечения с учетом требований производства", hours="36", credits="1.5", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 6.4 Өндірістегі бағдарламалық компоненттерге техникалық, экологиялық және қауіпсіздік шараларын ескере отырып IT шешімдерін әзірлеу.", name_ru="РО 6.4 Разрабатывать ИТ-решения для компонентов программного обеспечения в производстве с учетом технических, экологических и мер безопасности.", hours="36", credits="1.5", is_module_header=False, is_elective=False),
                Subject(name_kz="КМ 07 Back-end web ресурстарын құру", name_ru="ПМ 07 Разработка web-ресурсов Back-end", hours="216", credits="9", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 7.1 Сервер мен клиенттік жүйелер арасындағы байланысты басқару.", name_ru="РО 7.1 Управление связью между серверными и клиентскими системами", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН7.2 Деректер базасымен жұмыс істеу технологиясын пайдалану.", name_ru="РО7.2 Использование технологий работы с базами данных", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН7.3 Бағдарламалық жасақтаманың дизайн үлгілерін пайданалу.", name_ru="РО7.3 Применение шаблонов проектирования программного обеспечения", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="КМ 08 UX/UI визуалды дизайн жасау", name_ru="ПМ 08 Визуальный дизайн UX/UI", hours="216", credits="9", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 8.1 Мультимедиялық қосымшаларды, web-элементтерді құрастыру.", name_ru="РО 8.1 Разработка мультимедийных компонентов и веб-элементов", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 8.2 Графикалық интерфейске қосу үшін графикалық материалдарды дайындау.", name_ru="РО 8.2 Создание графических материалов для пользовательского интерфейса", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 8.3 Графикалық пайдаланушы интерфейс элементтерінің визуалды дизайнын жасау.", name_ru="РО 8.3 Создание визуального дизайна элементов графического пользовательского интерфейса.", hours="72", credits="3", is_module_header=False, is_elective=False),
    ],
    4: [
                Subject(name_kz="КМ 09 Мобильді қосымшаларды әзірлеу", name_ru="ПМ 09 Разработка мобильных приложений", hours="192", credits="8", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 9.1 Деректерді жасау, сақтау және басқару үшін дерекқорды басқару жүйесін пайдалану.", name_ru="РО 9.1 Использование систем управления базами данных для хранения и обработки данных", hours="48", credits="2", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 9.2 Деңгейлі құрылғыларды жасау.", name_ru="РО 9.2 Создание уровневых устройств.", hours="48", credits="2", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 9.3 Клиент-сервер негізіндегі жүйе үшін мобильді интерфейсті құрастыру.", name_ru="РО 9.3 Создание мобильных приложений на основе клиент-серверных решений", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="КМ 10 Жобалық іс-шараларды орындау", name_ru="ПМ 10 Реализация проектной деятельности", hours="216", credits="9", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 10.1 Өз бетінше жобалау жұмыстарына практикалық дағдыларды қалыптастыру.", name_ru="РО 10.1 Развитие практических навыков проектирования", hours="24", credits="1", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 10.2 Тәжірибе негізінде қолданылатын заманауи жобалау және есептеу әдістерін қолдану", name_ru="РО 10.2 Применение современных методов проектирования и расчетов", hours="48", credits="2", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 10.3 Дипломдық жоба тақырыбы бойынша техникалық шарттар мен техникалық ұсыныстар әзірлеу.", name_ru="РО 10.3 Разработка технических условий и предложений по дипломному проекту", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 10.4 Дипломдық жоба тақырыбы бойынша бағдарламалық өнімді әзірлеу.", name_ru="РО 10.4 Разработка программного продукта по теме дипломного проекта", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="Кәсіптік практика КМ3. ОН3.2, ОН3.3; КМ4. ОН4.3; КМ5. ОН5.2, ОН5.3; КМ7. ОН7.1, ОН7.2, ОН7.3; КМ8. ОН8.1, ОН8.2, ОН8.3; КМ9. ОН9.1, ОН9.2, ОН9.3.", name_ru="Профессиональная практика ПМ3. РО3.2, РО3.3; ПМ4. РО4.3; ПМ5. РО5.2, РО5.3; ПМ7. РО7.1, РО7.2, РО7.3; ПМ8. РО8.1, РО8.2, РО8.3; ПМ9. РО9.1, РО9.2, РО9.3.", hours="504", credits="21", is_module_header=False, is_elective=False),
                Subject(name_kz="Қорытынды аттестаттау:", name_ru="Итоговая аттестация:", hours="", credits="", is_module_header=False, is_elective=False),
                Subject(name_kz="Ф1 Факультативтік ағылшын тілі", name_ru="Факультатив английский язык", hours="", credits="", is_module_header=False, is_elective=True),
                Subject(name_kz="Ф2 Факультативтік түрік тілі", name_ru="Факультатив турецкий язык", hours="", credits="", is_module_header=False, is_elective=True),
                Subject(name_kz="Ф3 Факультативтік кәсіпкерлік қызмет негіздері", name_ru="Факультатив основы предпринимательской деятельности", hours="", credits="", is_module_header=False, is_elective=True),
    ],
}

PROGRAM_IT = {
    "code": "IT",
    "name_kz": "IT технологиялары",
    "name_ru": "IT технологии",
    "sheets": ["3Ғ-1", "3Ғ-2", "3Ғ-3", "3Ғ-4"],
    "pages": PROGRAM_IT_PAGES
}

# ─────────────────────────────────────────────────────────────
# ACCOUNTING PROGRAM (D-group / 3D)
# ─────────────────────────────────────────────────────────────

PROGRAM_ACCOUNTING_PAGES = {
    1: [
                Subject(name_kz="Қазақ тілі", name_ru="Казахский язык", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="Қазақ әдебиеті", name_ru="Казахская литература", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="Орыс тілі және әдебиеті", name_ru="Русский язык и литература", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="Ағылшын тілі", name_ru="Английский язык", hours="216", credits="9", is_module_header=False, is_elective=False),
                Subject(name_kz="Қазақстан тарихы", name_ru="История Казахстана", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="Математика", name_ru="Математика", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="Информатика", name_ru="Информатика", hours="48", credits="2", is_module_header=False, is_elective=False),
                Subject(name_kz="Алғашқы әскери және технологиялық дайындық", name_ru="Начальная военная и технологическая подготовка", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="Дене тәрбиесі", name_ru="Физическая культура", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="География", name_ru="География", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="Биология", name_ru="Биология", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="Физика", name_ru="Физика", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="Графика және жобалау", name_ru="Графика и проектирование", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="БМ1 Дене қасиеттерін дамыту және жетілдіру", name_ru="БМ 1. Развитие и совершенствование физических качеств", hours="240", credits="10", is_module_header=False, is_elective=False),
                Subject(name_kz="БМ2 Ақпараттық-коммуникациялық және цифрлық технологияларды қолдану", name_ru="БМ 2. Применение информационно-коммуникационных и цифровых технологий", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="БМ3 Экономиканың базалық білімін және кәсіпкерлік негіздерін қолдану", name_ru="БМ 3. Применение базовых знаний экономики и основ предпринимательства", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="БМ4 Қоғам мен еңбек ұжымында әлеуметтену және бейімделу үшін әлеуметтік ғылымдар нгіздерін қолдану", name_ru="БМ 4. Применение основ социальных наук для социализации и адаптации в обществе и трудовом коллективе", hours="24", credits="1", is_module_header=False, is_elective=False)
    ],
    2: [
                Subject(name_kz="КМ 1 Бизнестің мақсаттары мен түрлерін, негізгі мүдделі тараптармен өзара әрекеттесуін түсіну", name_ru="ПМ 1 Понимание целей и видов бизнеса, взаимодействие с ключевыми заинтересованными сторонами", hours="384", credits="16", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 1.1 Бизнестің мақсаттары мен түрлерін, олардың негізгі мүдделі тараптармен және сыртқы ортамен өзара әрекеттесуін түсіну", name_ru="РО1.1 Понимание целей и видов бизнеса, и их взаимодействия с ключевыми заинтересованными сторонами и внешней средой", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 1.2 Көрсеткіштік және логарифмдік функциялар, сызықтық теңдеулер мен матрицалар жүйелері, сызықтық теңсіздіктер және сызықтық бағдарламалау, ықтималдық математикасын білу, бизнес және қаржылық қолдану мәселелерінде ақпаратты талдау және түсіндіру үшін ұғымдарды қолдана білу", name_ru="РО 1.2 Знание показательных и логарифмических функций, систем линейных уравнений и матриц, линейных неравенств и линейного программирования, основ теории вероятностей; применение этих понятий для анализа и интерпретации информации в бизнесе и финансовых расчетах.", hours="96", credits="4", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 1.3 Қаржылық есептіліктің мәні мен мақсатын түсіну, қаржылық ақпараттың сапалық сипаттамаларын анықтау, қаржылық есептілікті дайындау", name_ru="РО 1.3Понимание сути и назначения финансовой отчетности, определение качественных характеристик финансовой информации, подготовка финансовой отчетности.", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 1.4 Маркетингтің негізгі тұжырымдамаларды түсіну, маркетингтік ортаны зерттеу, тұтынушылар мен ұйымның сатып алу тәртібін түсіну, нарықтарды сегменттеу және өнімдерді орналастыру, жаңа өнімдерді әзірлеу үшін қолданылатын құралдар мен әдістерді білу", name_ru="РО 1.4 Понимание основных маркетинговых концепций, исследование маркетинговой среды, изучение поведения потребителей и организаций, сегментация рынков, размещение товаров и разработка новых продуктов, знание инструментов и методов, используемых в этих процессах.", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="КМ 2Кәсіптік салада тілдік дағдыларды қолдану", name_ru="ПМ 2 Использование языковых навыков в профессиональной сфере", hours="336", credits="14", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 2.1 Академиялық деңгейде Ағылшын тілінің оқылым, айтылым және жазылым дағдыларын еркін меңгеру", name_ru="РО 2.1 Свободное владение навыками чтения, говорения и письма на английском языке на академическом уровне.", hours="168", credits="7", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 2.2 Кәсіби салада Ағылшын тілінің айтылым және жазылым дағдыларын B2 деңгейінде еркін меңгеру", name_ru="РО2.2 Свободное владение навыками говорения и письма на английском языке в профессиональной сфере на уровне B2.", hours="48", credits="2", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 2.3 Іскерлік мақсатта қазақ тілін қолдану", name_ru="РО2.3 Использование казахского языка в деловых целях.", hours="48", credits="2", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 2.4 Іскерлік мақсатта түрік тілін қолдану", name_ru="РО 2.4 Использование турецкого языка в деловых целях", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="КМ 3 Бухгалтерлік (қаржылық) есептілікті жасауға қатысу", name_ru="ПМ 3 Участие в составлении бухгалтерской (финансовой) отчетности", hours="288", credits="12", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 3.1Басқару ақпаратының сипатын, мақсатын түсіну, шығындарды есепке алу, жоспарлау, бизнестің тиімділігін бақылау", name_ru="РО 3.1 Понимание характеристик и целей управленческой информации, учет затрат, планирование и контроль эффективности бизнеса.отчетности", hours="120", credits="5", is_module_header=False, is_elective=False)
    ],
    3: [
                Subject(name_kz="ОН 3.2 Еңбек қатынастарына қатысты заңды түсіну, компаниялардың қалай басқарылатындығын және реттелетінін сипаттау және түсіну", name_ru="РО 3.2 Понимание законодательства о трудовых отношениях, принципов управления и регулирования деятельности компаний.", hours="48", credits="2", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 3.3 Iскерлік шешім қабылдау процесін қолдайтын жалпы математикалық құралдарды қолдану, аналитикалық әдістерді әртүрлі бизнес қолданбаларында қолдану", name_ru="РО 3.3 Использование математических инструментов, поддерживающих процесс принятия деловых решений, применение аналитических методов в различных бизнес-приложениях.", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 3.4 Негізгі экономикалық принциптерді, макроэкономикалық мәселелерді және көрсеткіштерді есептеуді білу, фискалдық және ақша-несие саясатының макроэкономикаға әсер ету механизмдерді талдау", name_ru="РО3.4 Знание основных экономических принципов, макроэкономических проблем и показателей, расчет фискальных и кредитно-денежных политик, анализ механизмов их влияния на макроэкономику.", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="КМ 4 Ұйымның және оның бөлімшелерінің шаруашылық-қаржылық қызметін кешенді талдауға қатысу", name_ru="ПМ 4 Участие в комплексном анализе хозяйственно-финансовой деятельности организации и ее подразделений", hours="528", credits="22", is_module_header=True, is_elective=False),
                Subject(name_kz="ОН 4.1 Инвестициялар мен қаржыландыруды бағалаудың баламалы тәсілдерін салыстыру, қаржы саласындағы проблемаларды шешудің әртүрлі тәсілдерінің орындылығын бағалау.", name_ru="РО4.1 Сравнение альтернативных методов оценки инвестиций и финансирования, оценка целесообразности различных способов решения проблем в финансовой сфере.", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 4.2 Ұйымдарға өнімділікті басқару және өлшеу үшін қажет ақпаратты, технологиялық жүйелерді анықтау, шығындарды есепке алу және басқару есебі әдістерін қолдану.", name_ru="РО 4.2 Определение информации и технологических систем, необходимых для управления продуктивностью организаций, учет затрат и применение методов управленческого учета.", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 4.3 Салық жүйесінің жұмыс істеуі мен көлемін және оны басқаруды түсіну", name_ru="РО4.3 Понимание функционирования и структуры налоговой системы, а также принципов ее управления.", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 4.4 ХҚЕС стандарттарына сәйкес операцияларды есепке алу, Қаржылық есептерді талдау және түсіндіру", name_ru="РО 4.4 Учет операций в соответствии со стандартами МСФО, анализ и интерпретация финансовой отчетности", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 4.5 Бизнес статистикадағы негізгі түсініктерді, деректер материалдарын жинау, қорытындылау және талдау әдістерін білу", name_ru="РО 4.5 Знание основных понятий бизнес-статистики, сбор, обобщение и методы анализа данных.", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 4.6 Бухгалтерлік есептің ақпараттық жүйелері", name_ru="РО 4.6 Информационные системы бухгалтерского учета.", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 4.7 Аудит ұғымының, функцияларының, корпоративтік басқарудың, оның ішінде этика мен кәсіби мінез-құлықтың анықтамасы, Халықаралық аудит стандарттарын (АХС) қолдану", name_ru="РО 4.7 Понимание концепции аудита, его функций, корпоративного управления, включая вопросы этики и профессионального поведения, применение Международных стандартов аудита (МСА).", hours="120", credits="5", is_module_header=False, is_elective=False),
                Subject(name_kz="КМ 5 Қаржы менеджментіне экономикалық ортаның әсерін бағалау", name_ru="ПМ 5 Оценка влияния экономической среды на финансовый менеджмент", hours="144", credits="6", is_module_header=True, is_elective=False)
    ],
    4: [
                Subject(name_kz="ОН 5.1Қаржылық басқару функциясының рөлі мен мақсатын түсіну, Қаржы менеджментіне экономикалық ортаның әсерін бағалау", name_ru="РО 5.1Понимание роли и целей функций финансового управления, анализ влияния экономической среды на финансовый менеджмент.", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="ОН 5.2Инвестицияларға тиімді бағалау жүргізу, Бизнесті қаржыландырудың балама көздерін анықтау және бағалау", name_ru="РО 5.2 Эффективная оценка инвестиций, определение и анализ альтернативных источников финансирования бизнеса.", hours="72", credits="3", is_module_header=False, is_elective=False),
                Subject(name_kz="Кәсіптік практика КМ1 ОН1.3; КМ3 ОН3.1, ОН3.2, ОН3.3, ОН3.4; КМ4 ОН 4.1, ОН 4.2, ОН 4.3, ОН 4.4, ОН 4.5, ОН 4.6, ОН 4.7; КМ5 ОН 5.1, ОН 5.2.", name_ru="Профессиональная практика ПМ1 ОН 1.3; КМ3 ОН3.1, ОН3.2, ОН3.3, ОН3.4; КМ4 ОН 4.1, ОН 4.2, ОН 4.3, ОН 4.4, ОН 4.5, ОН 4.6, ОН 4.7; КМ5 ОН 5.1, ОН 5.2.", hours="432", credits="18", is_module_header=False, is_elective=False),
                Subject(name_kz="Қорытынды аттестаттау:", name_ru="Итоговая аттестация:", hours="", credits="", is_module_header=False, is_elective=False),
                Subject(name_kz="Ф1 Факультативтік ағылшын тілі", name_ru="Факультатив английский язык", hours="", credits="", is_module_header=False, is_elective=True),
                Subject(name_kz="Ф2 Факультативтік түрік тілі", name_ru="Факультатив турецкий язык", hours="", credits="", is_module_header=False, is_elective=True),
                Subject(name_kz="Ф3 Факультативтік Бизнес және бухгалтерлік есептегі жағдайлар (Cases in Business and Accounting)", name_ru="Ф3 Факультатив Ситуации в бизнесе и бухгалтерском учете (Cases in Business and Accounting)", hours="", credits="", is_module_header=False, is_elective=True),
                Subject(name_kz="Ф4 Факультативтік Бизнес деректерін талдау (Business data analysis (excel, macros, google sheets, sql, python, power BI, tableau))", name_ru="Ф4 Факультатив Анализ бизнес данных (Business data analysis (excel, macros, google sheets, sql, python, power BI, tableau))", hours="", credits="", is_module_header=False, is_elective=True),
                Subject(name_kz="Ф5 Факультативтік кәсіпкерлік қызмет негіздері (Enterpreneurship)", name_ru="Ф5 Факультатив основы предпринимательской деятельности (Enterpreneurship)", hours="", credits="", is_module_header=False, is_elective=True)
    ],
}

PROGRAM_ACCOUNTING = {
    "code": "ACCOUNTING",
    "name_kz": "Бухгалтерлік есеп",
    "name_ru": "Бухгалтерский учет",
    "sheets": ["3D-1", "3D-2"],
    "pages": PROGRAM_ACCOUNTING_PAGES
}

# ─────────────────────────────────────────────────────────────
# PROGRAM REGISTRY
# ─────────────────────────────────────────────────────────────

PROGRAMS = {
    "IT": PROGRAM_IT,
    "ACCOUNTING": PROGRAM_ACCOUNTING,
}

# ─────────────────────────────────────────────────────────────
# UTILITY FUNCTIONS
# ─────────────────────────────────────────────────────────────

def get_program_config(program_code: str) -> dict:
    """Get configuration for a program.
    
    Args:
        program_code: Program code (e.g., "IT", "ACCOUNTING")
        
    Returns:
        Program configuration dictionary
        
    Raises:
        ValueError: If program not found
    """
    if program_code not in PROGRAMS:
        raise ValueError(f"Unknown program: {program_code}. Available: {list(PROGRAMS.keys())}")
    return PROGRAMS[program_code]


def get_sheets_for_program(program_code: str) -> list:
    """Get Excel sheet names for a program.
    
    Args:
        program_code: Program code
        
    Returns:
        List of sheet names
    """
    config = get_program_config(program_code)
    return config.get("sheets", [])


def get_program_pages(program_code: str) -> dict:
    """Get page definitions for a program.
    
    Args:
        program_code: Program code
        
    Returns:
        Dictionary of page definitions
    """
    config = get_program_config(program_code)
    return config.get("pages", {})
