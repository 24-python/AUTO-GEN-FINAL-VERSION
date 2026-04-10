# add_telezki_uborochnye.py
from db.database import SessionLocal
from db.models import Object, Category, Instruction

session = SessionLocal()

# Находим категорию "Моечный, уборочный инвентарь и оборудование"
category = session.query(Category).filter(Category.name == "Моечный, уборочный инвентарь и оборудование").first()
if not category:
    print("❌ Категория не найдена")
    session.close()
    exit()

# Проверяем, нет ли уже такого объекта
existing = session.query(Object).filter(Object.normalized_name == "тележки уборочные").first()
if existing:
    print(f"⚠️ Объект уже существует: {existing.display_name} (id={existing.id})")
    session.close()
    exit()

# Создаём объект
obj = Object(
    normalized_name="тележки уборочные",
    display_name="тележки уборочные",
    base_name="тележки уборочные",
    modifier=None,
    category_id=category.id,
    sort_priority=0
)
session.add(obj)
session.flush()
print(f"✅ Добавлен объект: {obj.display_name} (id={obj.id}, категория={category.name})")

# Добавляем инструкции (если нужны)
instructions_data = [
    {
        "cleaning_method": "мойка",
        "product_name": "ХИМИТЕК УНИВЕРСАЛ-ПД-Н",
        "cleaning_technique": "ручной (замачивание)",
        "concentration": "3,0 % – 4,0 %, дозирующая система ProMax",
        "temperature": "не ниже 40°C",
        "inventory": "промаркированный",
        "frequency": "1 раз в смену",
        "executor": "сотрудник внутреннего клининга",
        "control_method": "визуальный контроль/сотрудник отдела качества"
    },
    {
        "cleaning_method": "дезинфекция",
        "product_name": "ХИМИТЕК ПОЛИДЕЗ®-СУПЕР",
        "cleaning_technique": "ручной (протирание)",
        "concentration": "0,02 %–0,04 % (по НУК), спрей-система ProTwin",
        "temperature": "18-20 ºС",
        "exposure_time": "30 мин",
        "inventory": "промаркированный",
        "frequency": "1 раз в смену",
        "executor": "сотрудник внутреннего клининга",
        "control_method": "визуальный контроль/сотрудник отдела качества"
    }
]

for instr_data in instructions_data:
    instr = Instruction(
        object_id=obj.id,
        cleaning_method=instr_data["cleaning_method"],
        instruction_number="",
        product_name=instr_data["product_name"],
        cleaning_technique=instr_data["cleaning_technique"],
        concentration=instr_data["concentration"],
        temperature=instr_data["temperature"],
        exposure_time=instr_data.get("exposure_time", ""),
        inventory=instr_data["inventory"],
        frequency=instr_data["frequency"],
        executor=instr_data["executor"],
        control_method=instr_data["control_method"]
    )
    session.add(instr)
    print(f"   ➕ Добавлена инструкция: {instr_data['cleaning_method']} | {instr_data['product_name']}")

session.commit()
session.close()

print("\n✅ Готово!")