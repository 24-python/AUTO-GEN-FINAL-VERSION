# fix_cameras_holod.py
from db.database import SessionLocal
from db.models import Object

session = SessionLocal()

obj = session.query(Object).filter(Object.normalized_name == "камеры холод").first()
if obj:
    print(f"🔍 Найден: {obj.normalized_name}")
    obj.normalized_name = "камеры холд"
    session.commit()
    print(f"✅ Исправлено: {obj.normalized_name}")
else:
    print("❌ Объект 'камеры холод' не найден")

session.close()