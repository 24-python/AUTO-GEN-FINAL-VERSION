# find_duplicates_in_one_object.py
# !!!!!ПОИСК ПО СТРОЧНО ДУБЛИКАТОВ И КОЛИЧЕСТВА ИНСТРУКЦИЙ ПРИ СОВПАДЕНИИ
from db.database import SessionLocal
from db.models import Instruction, Object
from collections import defaultdict

def get_full_key(instr):
    """Полный ключ инструкции по всем полям"""
    return (
        instr.cleaning_method or "",
        instr.instruction_number or "",
        instr.product_name or "",
        instr.cleaning_technique or "",
        instr.concentration or "",
        instr.temperature or "",
        instr.exposure_time or "",
        instr.inventory or "",
        instr.frequency or "",
        instr.executor or "",
        instr.control_method or "",
    )

session = SessionLocal()

# Находим объект "пол"
obj = session.query(Object).filter(Object.display_name == "тележки подкатные").first()
if not obj:
    print("❌ Объект не найден")
    session.close()
    exit()

print(f"🔍 Объект: {obj.display_name} (id={obj.id})")
print("=" * 60)

instructions = session.query(Instruction).filter(Instruction.object_id == obj.id).all()

# Группируем по полному ключу
grouped = defaultdict(list)
for instr in instructions:
    key = get_full_key(instr)
    grouped[key].append(instr)

print(f"\n📋 Всего инструкций: {len(instructions)}")
print(f"📋 Уникальных инструкций: {len(grouped)}")
print("-" * 40)

duplicates_found = False
for key, instr_list in grouped.items():
    if len(instr_list) > 1:
        duplicates_found = True
        print(f"\n❌ Найдено {len(instr_list)} дубликатов:")
        for instr in instr_list:
            print(f"   • id={instr.id}, метод={instr.cleaning_method}, средство={instr.product_name}, периодичность={instr.frequency}")

if not duplicates_found:
    print("\n✅ Дубликатов нет")

session.close()