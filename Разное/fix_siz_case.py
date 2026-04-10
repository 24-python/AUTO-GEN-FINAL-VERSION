# fix_case_all.py
from db.database import SessionLocal
from db.models import Object

session = SessionLocal()

# Список объектов для исправления регистра
objects_to_fix = [
    "АВД",
    "завесы ПВХ",
]

fixed_count = 0

for old_name in objects_to_fix:
    obj = session.query(Object).filter(Object.normalized_name == old_name).first()
    if obj:
        new_name = old_name.lower()
        print(f"🔍 {old_name} → {new_name}")
        obj.normalized_name = new_name
        fixed_count += 1
    else:
        print(f"❌ {old_name} не найден")

# Дополнительно: все normalized_name привести к нижнему регистру
print("\n🔄 Приведение всех normalized_name к нижнему регистру...")
all_objects = session.query(Object).all()
for obj in all_objects:
    if obj.normalized_name != obj.normalized_name.lower():
        print(f"  {obj.normalized_name} → {obj.normalized_name.lower()}")
        obj.normalized_name = obj.normalized_name.lower()
        fixed_count += 1

session.commit()
session.close()

print(f"\n✅ Исправлено {fixed_count} объектов")