# check_compound_objects.py
from db.database import SessionLocal
from db.models import Object
import re


def normalize_name(name: str) -> str:
    """Функция нормализации из парсера"""
    normalized = re.sub(r'[^\w\s\-\(\)]', '', name)
    normalized = normalized.lower()
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized.strip()


# Список объектов для проверки
test_objects = [
    "вытяжные зонты А",
    "вытяжные зонты Н",
    "листы для выпечки А",
    "листы для выпечки Н",
    "потолок (высота _____м) ОК",
    "потолок (высота _____м) П",
    "производственные столы Д",
    "производственные столы Н",
    "съёмные детали оборудования А",
    "съёмные детали оборудования Н",
    "формы для выпечки А",
    "формы для выпечки Н",
    "формы для выпечки С",
]

session = SessionLocal()

print("🔍 ПРОВЕРКА СООТВЕТСТВИЯ ПАРСЕР → БД")
print("=" * 60)

for obj_name in test_objects:
    normalized = normalize_name(obj_name)

    # Ищем в БД
    obj = session.query(Object).filter(Object.normalized_name == normalized).first()

    if obj:
        print(f"✅ {obj_name}")
        print(f"   normalized: {obj.normalized_name}")
        print(f"   display:    {obj.display_name}")
    else:
        print(f"❌ {obj_name} → НЕ НАЙДЕН в БД")
        print(f"   Ищем как: {normalized}")

session.close()