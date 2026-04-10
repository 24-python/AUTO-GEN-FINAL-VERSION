# debug_search.py
from db.database import SessionLocal
from db.models import Object
from generator.mapper import find_object_and_instructions
from parser.models import Category

session = SessionLocal()

# Тестируем разные варианты
test_names = [
    "камеры мороз.",
    "камеры мороз",
    "камеры шок. замор.",
    "камеры шок замор",
]

for name in test_names:
    obj, _ = find_object_and_instructions(session, name, Category.REFRIGERATION_EQUIPMENT)
    if obj:
        print(f"✅ '{name}' → {obj.normalized_name}")
    else:
        print(f"❌ '{name}' → не найден")

session.close()