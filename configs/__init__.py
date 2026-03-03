from .it import IT_CONFIG, TERMS as IT_TERMS
from .acc import ACC_CONFIG, TERMS as ACC_TERMS

# Экспорт для удобного доступа
CONFIGS = {
    '3F': {
        'config': IT_CONFIG,
        'terms': IT_TERMS,
        'template_kz': 'Diplom_F_KZ_Template (4).xlsx',
        'template_ru': 'Diplom_F_RU_Template (4).xlsx',
    },
    '3D': {
        'config': ACC_CONFIG,
        'terms': ACC_TERMS,
        'template_kz': 'Diplom_D_KZ_Template(4).xlsx',
        'template_ru': 'Diplom_D_RU_Template(4).xlsx',
    }
}

def get_config(group_type: str, lang: str):
    """
    Возвращает (списки_предметов, термины, имя_шаблона) для заданной специальности и языка.
    group_type: '3F' (IT) или '3D' (Учет)
    lang: 'kz' или 'ru'
    """
    base = CONFIGS.get(group_type)
    if not base:
        raise ValueError(f"Unknown group type: {group_type}")
    
    config = base['config'].get(lang)
    terms  = base['terms'].get(lang)
    tmpl   = base['template_kz'] if lang == 'kz' else base['template_ru']
    
    if not config:
        raise ValueError(f"Unknown language: {lang}")
        
    return config, terms, tmpl
