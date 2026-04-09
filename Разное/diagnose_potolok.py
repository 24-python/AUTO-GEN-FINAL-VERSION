# diagnose_potolok.py
from db.database import SessionLocal
from db.models import Object, Instruction

session = SessionLocal()

# Ищем объекты "потолок"
objects = session.query(Object).filter(Object.name.like("%потолок%")).all()
print(f"Найдено объектов с 'потолок': {len(objects)}")
for obj in objects:
    instr_count = session.query(Instruction).filter(Instruction.object_id == obj.id).count()
    print(f"  - {obj.name} (id={obj.id}): {instr_count} инструкций")

session.close()