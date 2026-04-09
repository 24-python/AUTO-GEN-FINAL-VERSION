# add_simple_names.py
from db.database import SessionLocal
from db.models import Object, Category

session = SessionLocal()

# Находим категории
inventory_cat = session.query(Category).filter(Category.name == "Инвентарь, посуда и т.д.").first()
tech_cat = session.query(Category).filter(Category.name == "Технологическое оборудование").first()
refrig_cat = session.query(Category).filter(Category.name == "Холодильное оборудование").first()

# Добавляем объекты с именами, которые выдает парсер
missing_objects = [
    ("внутрицеховая тара (вёдра ящики)", "внутрицеховая тара (вёдра ящики)", None, inventory_cat.id),
    ("машина для резки конд изделий", "машина для резки конд изделий", None, tech_cat.id),
    ("холод столы", "холод столы", None, refrig_cat.id),
    ("холод шкафы", "холод шкафы", None, refrig_cat.id),
]

for name, base_name, modifier, cat_id in missing_objects:
    existing = session.query(Object).filter(Object.name == name).first()
    if not existing:
        obj = Object(
            name=name,
            base_name=base_name,
            modifier=modifier,
            category_id=cat_id,
            sort_priority=0
        )
        session.add(obj)
        print(f"➕ Добавлен: {name}")
    else:
        print(f"⚠️ Уже существует: {name}")

session.commit()
session.close()
print("✅ Готово")