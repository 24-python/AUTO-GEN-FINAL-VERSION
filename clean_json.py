#!/usr/bin/env python3
"""
СКРИПТ ДЛЯ ОЧИСТКИ JSON ОТ ПУСТЫХ ДУБЛЕЙ

Правила:
1. Если для объекта есть несколько записей:
   - оставляем ТОЛЬКО те, у которых есть хотя бы одно заполненное поле
   - пустые дубли УДАЛЯЕМ
2. Если объект уникальный (одна запись):
   - оставляем даже пустым (для базы данных)
"""

import json
from pathlib import Path


def is_instruction_empty(instr):
    """Проверяет, пустая ли инструкция"""
    # Поля, которые проверяем на заполненность
    fields_to_check = [
        "cleaning_method",
        "product_name",
        "concentration",
        "temperature",
        "exposure_time",
        "frequency",
        "executor",
        "control_method"
    ]

    for field in fields_to_check:
        if instr.get(field, "").strip():
            return False
    return True


def clean_instructions(json_data):
    """Очищает инструкции от пустых дублей"""

    # Группируем инструкции по object_name
    grouped = {}
    for instr in json_data["instructions"]:
        key = instr["object_name"]
        if key not in grouped:
            grouped[key] = []
        grouped[key].append(instr)

    cleaned_instructions = []
    removed_count = 0

    for obj_name, instructions in grouped.items():
        if len(instructions) == 1:
            # Уникальный объект - оставляем даже пустым
            cleaned_instructions.append(instructions[0])
            if is_instruction_empty(instructions[0]):
                print(f"  ➖ УНИКАЛЬНЫЙ ПУСТОЙ: {obj_name} (оставлен для БД)")
        else:
            # Несколько записей - оставляем только непустые
            non_empty = [instr for instr in instructions if not is_instruction_empty(instr)]
            empty_count = len(instructions) - len(non_empty)

            if non_empty:
                cleaned_instructions.extend(non_empty)
                if empty_count > 0:
                    print(f"  ✅ {obj_name}: оставлено {len(non_empty)}, удалено {empty_count} пустых")
                    removed_count += empty_count
            else:
                # Все записи пустые - оставляем одну (для БД)
                cleaned_instructions.append(instructions[0])
                print(f"  ⚠️ {obj_name}: все записи пустые, оставлена одна")
                removed_count += len(instructions) - 1

    return cleaned_instructions, removed_count


def main():
    print("🔧 ОЧИСТКА JSON ОТ ПУСТЫХ ДУБЛЕЙ")
    print("=" * 60)

    input_file = Path("normalized_data.json")
    output_file = Path("cleaned_data.json")

    if not input_file.exists():
        print(f"❌ Файл не найден: {input_file}")
        return

    # Загружаем JSON
    print(f"\n📖 Чтение файла: {input_file}")
    with open(input_file, 'r', encoding='utf-8') as f:
        data = json.load(f)

    print(f"   Всего инструкций: {len(data['instructions'])}")

    # Очищаем
    print("\n🔄 Обработка...")
    cleaned_instructions, removed = clean_instructions(data)

    # Обновляем данные
    data["instructions"] = cleaned_instructions
    data["statistics"]["total_instructions"] = len(cleaned_instructions)
    data["statistics"]["removed_empty_duplicates"] = removed

    # Сохраняем
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"\n✅ Очищенный JSON сохранен: {output_file}")
    print(f"\n📊 ИТОГ:")
    print(f"   - Было инструкций: {len(data['instructions']) + removed}")
    print(f"   - Удалено пустых дублей: {removed}")
    print(f"   - Стало инструкций: {len(data['instructions'])}")
    print(f"   - Объектов с инструкциями: {data['statistics']['objects_with_instructions']}")
    print(f"   - Объектов без инструкций: {data['statistics']['objects_without_instructions']}")


if __name__ == "__main__":
    main()