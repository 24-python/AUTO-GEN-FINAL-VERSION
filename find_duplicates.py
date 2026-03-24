# find_duplicates.py
import json
from collections import defaultdict

with open("normalized_data.json", 'r', encoding='utf-8') as f:
    data = json.load(f)

# Группируем по object_name
grouped = defaultdict(list)
for instr in data["instructions"]:
    grouped[instr["object_name"]].append(instr)

print("=" * 80)
print("ДУБЛИКАТЫ ИНСТРУКЦИЙ")
print("=" * 80)

total_duplicates = 0

for obj_name, instructions in grouped.items():
    if len(instructions) <= 1:
        continue

    # Создаем ключ для сравнения (убираем object_name и category)
    unique = {}
    for instr in instructions:
        key = (
            instr["cleaning_method"],
            instr["product_name"],
            instr["concentration"],
            instr["temperature"],
            instr["exposure_time"],
            instr["frequency"]
        )
        if key not in unique:
            unique[key] = []
        unique[key].append(instr)

    # Если есть дубликаты
    duplicates = sum(len(v) - 1 for v in unique.values() if len(v) > 1)
    if duplicates > 0:
        print(f"\n📌 {obj_name}: {len(instructions)} записей, дубликатов: {duplicates}")
        for key, vals in unique.items():
            if len(vals) > 1:
                print(f"   🔁 Дублируется {len(vals)} раз:")
                print(f"      метод: {vals[0]['cleaning_method']}")
                print(f"      средство: {vals[0]['product_name']}")
                print(f"      концентрация: {vals[0]['concentration']}")
        total_duplicates += duplicates

print(f"\n" + "=" * 80)
print(f"ВСЕГО ДУБЛИКАТОВ: {total_duplicates}")
print("=" * 80)