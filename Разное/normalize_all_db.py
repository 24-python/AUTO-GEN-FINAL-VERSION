# normalize_all_db.py
from db.database import SessionLocal
from db.models import Object
import re

session = SessionLocal()

def normalize(name: str) -> str:
    """Удаляет всё, кроме букв, цифр, пробелов, дефисов, скобок"""
    return re.sub(r'[^\w\s\-\(\)]', '', name)

objects = session.query(Object).all()
for obj in objects:
    new_name = normalize(obj.normalized_name)
    if new_name != obj.normalized_name:
        print(f"  {obj.normalized_name} → {new_name}")
        obj.normalized_name = new_name

session.commit()
session.close()
print("✅ Готово")