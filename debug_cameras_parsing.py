# debug_potolok.py
from parser.xml_parser import parse_checklist
from db.database import SessionLocal
from db.models import Object

# Парсим чек-лист
data = parse_checklist("Обновленный чек-лист.docx")

print("\n🔍 НАЙДЕННЫЕ ОБЪЕКТЫ С 'ПОТОЛОК'")
print("=" * 50)
for item in data.get_checked_items():
    if "потолок" in item.name.lower():
        print(f"  {item.name} → категория: {item.category.value}")

print("\n🔍 ПРОВЕРКА БД")
print("=" * 50)
session = SessionLocal()

# Ищем все объекты с "потолок" в БД
objects = session.query(Object).filter(Object.display_name.like("%потолок%")).all()
for obj in objects:
    print(f"  display: {obj.display_name}")
    print(f"  normalized: {obj.normalized_name}")
    print(f"  category: {obj.category.name if obj.category else '?'}")
    print()

session.close()