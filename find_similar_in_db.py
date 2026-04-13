# fix_cameras_normalized.py
from db.database import SessionLocal
from db.models import Object

session = SessionLocal()

# Находим объект с normalized_name = "камеры холд"
obj = session.query(Object).filter(Object.normalized_name == "камеры холд").first()

if obj:
    print(f"🔍 Найден объект:")
    print(f"   display_name: {obj.display_name}")
    print(f"   normalized_name (было): {obj.normalized_name}")

    # Заменяем
    obj.normalized_name = "камеры холод"
    session.commit()

    print(f"   normalized_name (стало): {obj.normalized_name}")
    print(f"✅ Готово!")
else:
    print("❌ Объект с normalized_name = 'камеры холд' не найден")

session.close()