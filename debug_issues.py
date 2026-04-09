# debug_issues.py
from db.database import SessionLocal
from db.models import Object
from generator.mapper import find_object_and_instructions
from parser.models import Category
import re


def normalize_name(name: str) -> str:
    return re.sub(r'[^\w\s\-\(\)]', '', name).lower().strip()


session = SessionLocal()

# Список проблемных объектов из чек-листа
test_cases = [
    ("корзины/диспенсеры для СИЗ", Category.SANITARY_POST),
    ("камеры холд.", Category.REFRIGERATION_EQUIPMENT),
    ("камеры мороз.", Category.REFRIGERATION_EQUIPMENT),
    ("камеры шок. замор.", Category.REFRIGERATION_EQUIPMENT),
]

print("🔍 ДИАГНОСТИКА ПОИСКА")
print("=" * 50)

for original_name, category in test_cases:
    normalized = normalize_name(original_name)
    print(f"\n📌 Исходное: {original_name}")
    print(f"   Нормализованное: {normalized}")

    # Ищем в БД по нормализованному имени
    obj = session.query(Object).filter(Object.normalized_name == normalized).first()
    if obj:
        print(f"   ✅ Найден в БД: {obj.normalized_name} → {obj.display_name}")
    else:
        print(f"   ❌ НЕ НАЙДЕН в БД")

        # Показываем похожие объекты для диагностики
        similar = session.query(Object).filter(
            Object.normalized_name.ilike(f"%{normalized[:10]}%")
        ).all()
        if similar:
            print(f"   📋 Похожие объекты в БД:")
            for s in similar:
                print(f"      - {s.normalized_name} → {s.display_name}")

    # Проверяем через mapper
    obj2, _ = find_object_and_instructions(session, original_name, category)
    if obj2:
        print(f"   ✅ mapper нашёл: {obj2.normalized_name}")
    else:
        print(f"   ❌ mapper НЕ нашёл")

session.close()