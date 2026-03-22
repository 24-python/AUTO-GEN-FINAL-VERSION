#!/usr/bin/env python3
"""
Скрипт для конвертации Excel-файла с инструкциями в JSON-формат для загрузки в БД.
Запуск: python excel_to_json.py "Объекты с определенной обработкой.xlsx"
"""

import json
import sys
from pathlib import Path
from openpyxl import load_workbook


def clean_text(value):
    """Очищает текст от лишних пробелов и переносов строк"""
    if value is None:
        return ""
    return str(value).strip().replace('\n', ' ').replace('\r', ' ')


def parse_excel_to_json(excel_path, output_json="seed_data.json"):
    """
    Парсит Excel и создает JSON в нужном формате
    """
    print(f"📂 Чтение файла: {excel_path}")

    # Загружаем книгу и активный лист
    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active

    # Словарь для хранения уникальных объектов
    objects_dict = {}
    instructions_list = []

    current_category = None
    current_object = None

    print("🔄 Обработка строк...")

    for row_idx, row in enumerate(ws.iter_rows(min_row=1, values_only=True), 1):
        # Пропускаем полностью пустые строки
        if not any(row):
            continue

        object_name = clean_text(row[0]) if len(row) > 0 else ""

        # Пропускаем заголовки разделов (они без инструкций)
        if object_name.endswith(':'):
            current_category = object_name.rstrip(':').strip()
            print(f"  📌 Раздел: {current_category}")
            continue

        # Если есть название объекта - это новая инструкция
        if object_name and object_name not in ["Объект обработки", ""]:
            cleaning_method = clean_text(row[1]) if len(row) > 1 else ""

            # Создаем запись инструкции со ВСЕМИ полями (даже пустыми)
            instruction = {
                "object_name": object_name,
                "category": current_category or "",
                "cleaning_method": cleaning_method,
                "instruction_number": clean_text(row[2]) if len(row) > 2 else "",
                "product_name": clean_text(row[3]) if len(row) > 3 else "",
                "cleaning_technique": clean_text(row[4]) if len(row) > 4 else "",
                "concentration": clean_text(row[5]) if len(row) > 5 else "",
                "temperature": clean_text(row[6]) if len(row) > 6 else "",
                "exposure_time": clean_text(row[7]) if len(row) > 7 else "",
                "inventory": clean_text(row[8]) if len(row) > 8 else "",
                "frequency": clean_text(row[9]) if len(row) > 9 else "",
                "executor": clean_text(row[10]) if len(row) > 10 else "",
                "control_method": clean_text(row[11]) if len(row) > 11 else ""
            }

            # Убираем None, заменяем на пустые строки
            for key, value in instruction.items():
                if value is None:
                    instruction[key] = ""

            instructions_list.append(instruction)

            # Сохраняем уникальные объекты (для objects.json)
            if object_name not in objects_dict:
                # Пытаемся определить base_name и modifier
                base_name = object_name
                modifier = None

                # Простая эвристика: если есть скобки - выделяем модификатор
                if '(' in object_name and ')' in object_name:
                    parts = object_name.split('(')
                    base_name = parts[0].strip()
                    modifier = parts[1].replace(')', '').strip()

                objects_dict[object_name] = {
                    "name": object_name,
                    "base_name": base_name,
                    "modifier": modifier,
                    "category": current_category or "",
                    "sort_priority": 0
                }

    print(f"\n📊 Статистика:")
    print(f"  - Найдено инструкций: {len(instructions_list)}")
    print(f"  - Найдено уникальных объектов: {len(objects_dict)}")

    # Формируем итоговый JSON
    result = {
        "categories": [],  # Пока пусто, добавим отдельно
        "objects": list(objects_dict.values()),
        "instructions": instructions_list
    }

    # Сохраняем JSON
    with open(output_json, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"✅ JSON сохранен в: {output_json}")

    # Также сохраняем отдельно список объектов для удобства
    objects_json = output_json.replace('.json', '_objects.json')
    with open(objects_json, 'w', encoding='utf-8') as f:
        json.dump(list(objects_dict.values()), f, ensure_ascii=False, indent=2)

    print(f"✅ Список объектов сохранен в: {objects_json}")

    return result


def validate_json_structure(json_data):
    """
    Проверяет соответствие JSON утвержденной структуре
    """
    print("\n🔍 Проверка структуры JSON...")

    required_keys = {
        "object_name": str,
        "category": str,
        "cleaning_method": str,
        "instruction_number": str,
        "product_name": str,
        "cleaning_technique": str,
        "concentration": str,
        "temperature": str,
        "exposure_time": str,
        "inventory": str,
        "frequency": str,
        "executor": str,
        "control_method": str
    }

    errors = 0
    for i, instr in enumerate(json_data["instructions"][:5]):  # Проверяем первые 5
        print(f"\n  Проверка инструкции {i + 1}: '{instr['object_name']}'")

        for key, expected_type in required_keys.items():
            if key not in instr:
                print(f"    ❌ Отсутствует ключ: {key}")
                errors += 1
            elif not isinstance(instr[key], expected_type):
                print(f"    ❌ Неверный тип для {key}: ожидался {expected_type}, получен {type(instr[key])}")
                errors += 1
            else:
                print(
                    f"    ✅ {key}: {instr[key][:30]}..." if len(str(instr[key])) > 30 else f"    ✅ {key}: {instr[key]}")

    if errors == 0:
        print("\n✅ Все проверки пройдены! Структура корректна.")
    else:
        print(f"\n⚠️ Найдено {errors} ошибок.")

    return errors == 0


def main():
    if len(sys.argv) < 2:
        print("Использование: python excel_to_json.py <путь_к_excel_файлу> [выходной_json]")
        return

    excel_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else "seed_data.json"

    if not Path(excel_file).exists():
        print(f"❌ Файл не найден: {excel_file}")
        return

    # Конвертируем
    json_data = parse_excel_to_json(excel_file, output_file)

    # Проверяем структуру
    validate_json_structure(json_data)


if __name__ == "__main__":
    main()