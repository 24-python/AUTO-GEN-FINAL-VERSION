# check_refrigerator_names.py
from db.database import SessionLocal
from db.models import Object, Category

session = SessionLocal()

# Находим категорию "Холодильное оборудование"
category = session.query(Category).filter(Category.name == "Холодильное оборудование").first()

if category:
    objects = session.query(Object).filter(Object.category_id == category.id).all()

    print(f"\n📋 ХОЛОДИЛЬНОЕ ОБОРУДОВАНИЕ ({len(objects)} объектов)")
    print("=" * 50)

    for obj in objects:
        print(f"  normalized_name: {obj.normalized_name}")
        print(f"  display_name:   {obj.display_name}")
        print()
else:
    print("❌ Категория 'Холодильное оборудование' не найдена")

session.close()