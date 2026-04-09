# check_dishwasher_debug.py
from parser.xml_parser import parse_checklist
from db.database import SessionLocal
from generator.mapper import find_object_and_instructions

# Парсим чек-лист
data = parse_checklist("Чек-лист проблемные.docx")

session = SessionLocal()

print("\n🔍 ПОСУДОМОЕЧНОЕ ОБОРУДОВАНИЕ: ЧТО ПОПАДАЕТ В ТЕХКАРТУ")
print("=" * 60)

for item in data.get_checked_items():
    obj, instructions = find_object_and_instructions(session, item.name, item.category)

    # Проверяем только посудомоечное оборудование
    if item.category.value == "Посудомоечное оборудование":
        print(f"\n📌 Объект из чек-листа: {item.name}")
        print(f"   Категория: {item.category.value}")

        if obj:
            print(f"   ✅ Найден в БД:")
            print(f"      normalized_name: {obj.normalized_name}")
            print(f"      display_name:    {obj.display_name}")
            print(f"   🖨️ В техкарту пойдёт: {obj.display_name}")
        else:
            print(f"   ❌ НЕ НАЙДЕН в БД")
            print(f"   🖨️ В техкарту пойдёт: {item.name}")

session.close()