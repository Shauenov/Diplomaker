from .subjects_it_kz import (
    PAGE1_SUBJECTS as KZ_P1,
    PAGE2_SUBJECTS as KZ_P2,
    PAGE3_SUBJECTS as KZ_P3,
    PAGE4_SUBJECTS as KZ_P4
)
from .subjects_it_ru import (
    PAGE1_SUBJECTS as RU_P1,
    PAGE2_SUBJECTS as RU_P2,
    PAGE3_SUBJECTS as RU_P3,
    PAGE4_SUBJECTS as RU_P4
)

# Переводы терминов для таблицы
TERMS = {
    'kz': {
        'traditional_elective': 'сынақ',
        'traditional_practice': 'сынақ',
        'traditional_attestation': 'өте жақсы',  # или брать из Excel
    },
    'ru': {
        'traditional_elective': 'зачтено',
        'traditional_practice': 'зачтено',
        'traditional_attestation': 'отлично',
    }
}

IT_CONFIG = {
    'kz': {
        'p1': KZ_P1,
        'p2': KZ_P2,
        'p3': KZ_P3,
        'p4': KZ_P4,
    },
    'ru': {
        'p1': RU_P1,
        'p2': RU_P2,
        'p3': RU_P3,
        'p4': RU_P4,
    }
}
