# check_fermentatory.py
from parser.xml_parser import parse_checklist
from db.database import SessionLocal
from db.models import Object
import re


# Функция нормализации как в парсере
def normalize_name(name: str) -> str:
    normalized = re.sub(r'[^\w\s\-\(\)]', '', name)
    normalized = normalized.lower()
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized.strip()


# Парсим чек-лист
data = parse_checklist("Обновленный чек-лист.docx")

print("\n🔍 НАЙДЕННЫЕ ОБЪЕКТЫ 'ФЕРМЕНТАТОР' В ЧЕК-ЛИСТЕ")
print("=" * 50)

fermentator_items = []
for item in data.get_checked_items():
    if "фермент" in item.name.lower():
        fermentator_items.append(item)
        print(f"  name: {item.name}")
        print(f"  category: {item.category.value}")
        print()

print("\n🔍 ПРОВЕРКА БД")
print("=" * 50)

session = SessionLocal()

# Ищем все объекты с "фермент" в БД
objects = session.query(Object).filter(Object.display_name.like("%фермент%")).all()
for obj in objects:
    print(f"  display_name: {obj.display_name}")
    print(f"  normalized_name: {obj.normalized_name}")
    print(f"  category: {obj.category.name if obj.category else '?'}")
    print()

# Сравнение
if fermentator_items:
    print("\n🔍 СРАВНЕНИЕ")
    print("=" * 50)
    for item in fermentator_items:
        normalized = normalize_name(item.name)
        print(f"Из чек-листа: {item.name}")
        print(f"  нормализованное: {normalized}")

        # Ищем в БД
        obj = session.query(Object).filter(Object.normalized_name == normalized).first()
        if obj:
            print(f"  ✅ Найдено в БД: {obj.display_name}")
        else:
            print(f"  ❌ НЕ НАЙДЕНО в БД")
            # Поиск похожих
            similar = session.query(Object).filter(Object.display_name.like("%фермент%")).all()
            if similar:
                print(f"  📋 Похожие в БД:")
                for s in similar:
                    print(f"     - {s.normalized_name} → {s.display_name}")
        print()

session.close()