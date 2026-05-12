#!/usr/bin/env python3
"""Проверка сопоставления имён из чек-листа с БД + детали инструкций."""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from parser.xml_parser import parse_checklist
from db.database import SessionLocal
from db.models import Object, Instruction, RoomCategory

def check_names(checklist_path: str, filter_text: str = None):
    data = parse_checklist(checklist_path)
    checked_items = data.get_checked_items()
    session = SessionLocal()

    print(f"📄 Чек-лист: {checklist_path}")
    print(f"🏢 Категория помещения: {data.room_category}")
    print(f"📋 Отмечено объектов: {len(checked_items)}")
    print("-" * 60)

    not_found = []
    for item in checked_items:
        norm_name = item.name
        if filter_text and filter_text.lower() not in norm_name.lower():
            continue

        obj = session.query(Object).filter(Object.normalized_name == norm_name).first()
        if obj:
            instructions = obj.instructions
            print(f"✅ Чек-лист: '{norm_name}' (id={obj.id}, display='{obj.display_name}')")
            print(f"   Инструкций всего: {len(instructions)}")
            for instr in instructions:
                room_cat_name = instr.room_category.name if instr.room_category else "Общая (NULL)"
                maint_type = instr.maintenance_type if instr.maintenance_type else "(пусто)"
                print(f"   - ID={instr.id}: метод={instr.cleaning_method}, "
                      f"maintenance_type='{maint_type}', room_category='{room_cat_name}'")
        else:
            print(f"❌ Чек-лист: '{norm_name}' — НЕ НАЙДЕН в БД")
            not_found.append(norm_name)
        print()

    session.close()

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Использование: python check_names.py <путь_к_чеклисту.docx> [фильтр]")
        sys.exit(1)
    path = sys.argv[1]
    filter_text = sys.argv[2] if len(sys.argv) > 2 else None
    check_names(path, filter_text)