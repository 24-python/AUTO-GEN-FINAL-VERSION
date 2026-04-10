# fix_cameras.py
from db.database import SessionLocal
from db.models import Object

session = SessionLocal()

obj = session.query(Object).filter(Object.normalized_name == "камеры холод").first()
if obj:
    obj.normalized_name = "камеры холод."
    session.commit()
    print(f"✅ Исправлено: {obj.normalized_name}")

session.close()