# normalize_all_to_lowercase.py
from db.database import SessionLocal
from db.models import Object

session = SessionLocal()

objects = session.query(Object).all()
updated = 0

for obj in objects:
    lower_name = obj.normalized_name.lower()
    if obj.normalized_name != lower_name:
        print(f"  {obj.normalized_name} → {lower_name}")
        obj.normalized_name = lower_name
        updated += 1

session.commit()
session.close()

print(f"\n✅ Приведено к нижнему регистру {updated} объектов")