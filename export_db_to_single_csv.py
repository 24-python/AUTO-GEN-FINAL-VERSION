#!/usr/bin/env python3
"""
export_db_to_single_csv.py

Экспортирует все данные из БД в один CSV файл.
Каждая строка = один объект со всеми его инструкциями.
Добавлены колонки: room_category_name, maintenance_type, surface_type.
"""

import csv
import sys
from pathlib import Path
from collections import defaultdict

sys.path.insert(0, str(Path(__file__).parent))

from db.database import SessionLocal
from db.models import Category, Object, Instruction, RoomCategory


def export_all_to_single_csv(output_path: str = "db_export_all1111.csv"):
    """Экспортирует все данные в один CSV"""

    session = SessionLocal()

    # Получаем все данные
    categories = {cat.id: cat for cat in session.query(Category).all()}
    room_categories = {rc.id: rc for rc in session.query(RoomCategory).all()}
    objects = session.query(Object).order_by(Object.category_id, Object.display_name).all()
    instructions = session.query(Instruction).all()

    # Группируем инструкции по object_id
    instr_by_object = defaultdict(list)
    for instr in instructions:
        instr_by_object[instr.object_id].append(instr)

    # Создаём CSV
    with open(output_path, 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f, delimiter=';')

        # Заголовок
        writer.writerow([
            'object_id',
            'category_name',
            'normalized_name',
            'display_name',
            'base_name',
            'modifier',
            'sort_priority',
            'instruction_id',
            'room_category_name',
            'maintenance_type',
            'surface_type',                # новое поле
            'cleaning_method',
            'instruction_number',
            'product_name',
            'cleaning_technique',
            'concentration',
            'temperature',
            'exposure_time',
            'inventory',
            'frequency',
            'executor',
            'control_method',
        ])

        # Данные
        total_rows = 0
        for obj in objects:
            cat = categories.get(obj.category_id)
            cat_name = cat.name if cat else ''

            obj_instrs = instr_by_object.get(obj.id, [])

            if obj_instrs:
                for instr in obj_instrs:
                    rc = room_categories.get(instr.room_category_id) if instr.room_category_id else None
                    room_cat_name = rc.name if rc else ''

                    writer.writerow([
                        obj.id,
                        cat_name,
                        obj.normalized_name,
                        obj.display_name,
                        obj.base_name,
                        obj.modifier or '',
                        obj.sort_priority,
                        instr.id,
                        room_cat_name,
                        instr.maintenance_type or '',
                        instr.surface_type or '',      # новое поле
                        instr.cleaning_method or '',
                        instr.instruction_number or '',
                        instr.product_name or '',
                        instr.cleaning_technique or '',
                        instr.concentration or '',
                        instr.temperature or '',
                        instr.exposure_time or '',
                        instr.inventory or '',
                        instr.frequency or '',
                        instr.executor or '',
                        instr.control_method or '',
                    ])
                    total_rows += 1
            else:
                writer.writerow([
                    obj.id,
                    cat_name,
                    obj.normalized_name,
                    obj.display_name,
                    obj.base_name,
                    obj.modifier or '',
                    obj.sort_priority,
                    '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''
                ])
                total_rows += 1

    session.close()

    print("=" * 80)
    print("📊 ЭКСПОРТ БАЗЫ ДАННЫХ В CSV")
    print("=" * 80)
    print(f"📁 Файл: {output_path}")
    print(f"📋 Объектов: {len(objects)}")
    print(f"📋 Инструкций: {len(instructions)}")
    print(f"📋 Строк в CSV: {total_rows}")
    print(f"📋 Категорий помещений: {len(room_categories)}")
    print("=" * 80)

    return output_path


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Экспорт БД в один CSV")
    parser.add_argument("--output", "-o", default="db_export_all1111.csv", help="Выходной CSV файл")
    args = parser.parse_args()

    export_all_to_single_csv(args.output)
    print("\n✅ Готово!")


if __name__ == "__main__":
    main()