#!/usr/bin/env python3
"""
Экспорт данных из базы данных в JSON
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import SessionLocal
from db.models import Category, Object, Instruction


def export_to_json(output_path: str):
    """Экспортирует все данные из БД в JSON"""
    session = SessionLocal()

    # Получаем категории
    categories = session.query(Category).order_by(Category.sort_order).all()

    # Получаем объекты
    objects = session.query(Object).order_by(Object.id).all()

    # Получаем инструкции
    instructions = session.query(Instruction).all()

    # Словарь объектов {id: name}
    object_names = {obj.id: obj.name for obj in objects}

    # Формируем JSON
    data = {
        "categories": [
            {"name": cat.name, "sort_order": cat.sort_order}
            for cat in categories
        ],
        "objects": [
            {
                "name": obj.name,
                "base_name": obj.base_name,
                "modifier": obj.modifier,
                "category": obj.category.name,
                "sort_priority": obj.sort_priority
            }
            for obj in objects
        ],
        "instructions": [
            {
                "object_name": object_names[instr.object_id],
                "category": obj.category.name,
                "cleaning_method": instr.cleaning_method,
                "instruction_number": instr.instruction_number,
                "product_name": instr.product_name,
                "cleaning_technique": instr.cleaning_technique,
                "concentration": instr.concentration,
                "temperature": instr.temperature,
                "exposure_time": instr.exposure_time,
                "inventory": instr.inventory,
                "frequency": instr.frequency,
                "executor": instr.executor,
                "control_method": instr.control_method
            }
            for instr in instructions
            for obj in objects if obj.id == instr.object_id
        ]
    }

    session.close()

    # Сохраняем
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"✅ Экспорт завершен: {output_path}")
    print(f"   Категорий: {len(data['categories'])}")
    print(f"   Объектов: {len(data['objects'])}")
    print(f"   Инструкций: {len(data['instructions'])}")

    return data


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Экспорт данных из БД в JSON")
    parser.add_argument("--output", "-o", default="exported_data.json", help="Выходной JSON файл")
    args = parser.parse_args()

    print("🔧 ЭКСПОРТ ДАННЫХ ИЗ БАЗЫ ДАННЫХ")
    print("=" * 50)

    export_to_json(args.output)


if __name__ == "__main__":
    main()