#!/usr/bin/env python3
"""
Импорт данных из JSON в базу данных
JSON структура: {
    "categories": [...],
    "objects": [
        {
            "normalized_name": "...",
            "display_name": "...",
            "base_name": "...",
            "modifier": "...",
            "category": "...",
            "sort_priority": 0,
            "instructions": [...]
        }
    ]
}
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import SessionLocal
from db.models import Category, Object, Instruction


def import_from_json(json_path: str):
    """Импортирует данные из JSON в БД"""

    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    session = SessionLocal()

    # Словарь категорий {name: id}
    categories = {cat.name: cat.id for cat in session.query(Category).all()}

    stats = {
        "objects_created": 0,
        "objects_updated": 0,
        "instructions_created": 0,
        "errors": 0
    }

    print("\n📥 ИМПОРТ ОБЪЕКТОВ")
    print("-" * 40)

    for obj_data in data["objects"]:
        category_id = categories.get(obj_data["category"])
        if not category_id:
            print(f"  ❌ Категория '{obj_data['category']}' не найдена")
            stats["errors"] += 1
            continue

        # Проверяем, существует ли объект по normalized_name
        existing = session.query(Object).filter_by(normalized_name=obj_data["normalized_name"]).first()

        if existing:
            # Обновляем
            existing.display_name = obj_data["display_name"]
            existing.base_name = obj_data["base_name"]
            existing.modifier = obj_data.get("modifier")
            existing.category_id = category_id
            existing.sort_priority = obj_data.get("sort_priority", 0)
            stats["objects_updated"] += 1
            print(f"  🔄 Обновлен: {obj_data['normalized_name']} → {obj_data['display_name']}")
            obj_id = existing.id
        else:
            # Создаем
            obj = Object(
                normalized_name=obj_data["normalized_name"],
                display_name=obj_data["display_name"],
                base_name=obj_data["base_name"],
                modifier=obj_data.get("modifier"),
                category_id=category_id,
                sort_priority=obj_data.get("sort_priority", 0)
            )
            session.add(obj)
            session.flush()
            stats["objects_created"] += 1
            print(f"  ➕ Создан: {obj_data['normalized_name']} → {obj_data['display_name']}")
            obj_id = obj.id

        # Импорт инструкций для этого объекта
        instructions = obj_data.get("instructions", [])
        for instr_data in instructions:
            # Проверяем, есть ли уже такая инструкция
            existing_instr = session.query(Instruction).filter(
                Instruction.object_id == obj_id,
                Instruction.cleaning_method == instr_data.get("cleaning_method", ""),
                Instruction.product_name == instr_data.get("product_name", "")
            ).first()

            if existing_instr:
                # Обновляем существующую
                existing_instr.cleaning_technique = instr_data.get("cleaning_technique", "")
                existing_instr.concentration = instr_data.get("concentration", "")
                existing_instr.temperature = instr_data.get("temperature", "")
                existing_instr.exposure_time = instr_data.get("exposure_time", "")
                existing_instr.inventory = instr_data.get("inventory", "")
                existing_instr.frequency = instr_data.get("frequency", "")
                existing_instr.executor = instr_data.get("executor", "")
                existing_instr.control_method = instr_data.get("control_method", "")
                existing_instr.instruction_number = instr_data.get("instruction_number", "")
            else:
                instruction = Instruction(
                    object_id=obj_id,
                    cleaning_method=instr_data.get("cleaning_method", ""),
                    instruction_number=instr_data.get("instruction_number", ""),
                    product_name=instr_data.get("product_name", ""),
                    cleaning_technique=instr_data.get("cleaning_technique", ""),
                    concentration=instr_data.get("concentration", ""),
                    temperature=instr_data.get("temperature", ""),
                    exposure_time=instr_data.get("exposure_time", ""),
                    inventory=instr_data.get("inventory", ""),
                    frequency=instr_data.get("frequency", ""),
                    executor=instr_data.get("executor", ""),
                    control_method=instr_data.get("control_method", "")
                )
                session.add(instruction)
                stats["instructions_created"] += 1

    session.commit()
    session.close()

    print("\n📊 СТАТИСТИКА ИМПОРТА")
    print("-" * 40)
    print(f"  Создано объектов: {stats['objects_created']}")
    print(f"  Обновлено объектов: {stats['objects_updated']}")
    print(f"  Создано инструкций: {stats['instructions_created']}")
    print(f"  Ошибок: {stats['errors']}")

    return stats


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Импорт данных из JSON в БД")
    parser.add_argument("json_file", help="Путь к JSON файлу")
    args = parser.parse_args()

    json_path = Path(args.json_file)
    if not json_path.exists():
        print(f"❌ Файл не найден: {json_path}")
        return

    print(f"🔧 ИМПОРТ ИЗ {json_path}")
    print("=" * 50)

    import_from_json(json_path)


if __name__ == "__main__":
    main()