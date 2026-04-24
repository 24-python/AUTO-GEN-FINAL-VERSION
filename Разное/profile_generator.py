#!/usr/bin/env python3
"""
profile_generator.py

Профилирование генератора для поиска реального узкого места.
"""

import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from parser.xml_parser import parse_checklist
from Разное.docx_generator import TechCardGenerator


def profile_generation(file_path: str):
    print("=" * 80)
    print("🔍 ПРОФИЛИРОВАНИЕ ГЕНЕРАТОРА")
    print("=" * 80)

    # 1. Парсинг
    print("\n1️⃣ Парсинг чек-листа...")
    t1 = time.time()
    checklist_data = parse_checklist(file_path)
    t2 = time.time()
    print(f"   ✅ {t2 - t1:.2f} сек")

    # 2. Загрузка из БД
    print("\n2️⃣ Загрузка из БД...")
    from db.database import SessionLocal
    from db.models import Object as DBObject, Instruction

    session = SessionLocal()
    t1 = time.time()

    checked_items = checklist_data.get_checked_items()
    all_names = [item.name for item in checked_items]
    db_objects = session.query(DBObject).filter(DBObject.normalized_name.in_(all_names)).all()
    objects_dict = {obj.normalized_name: obj for obj in db_objects}
    object_ids = [obj.id for obj in db_objects]
    db_instructions = session.query(Instruction).filter(Instruction.object_id.in_(object_ids)).all()

    t2 = time.time()
    print(f"   ✅ {t2 - t1:.2f} сек (объектов: {len(db_objects)}, инструкций: {len(db_instructions)})")
    session.close()

    # 3. Генерация документа
    print("\n3️⃣ Генерация документа...")
    generator = TechCardGenerator()

    # Замеряем по этапам внутри генератора
    import tempfile
    with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as tmp:
        output_path = tmp.name

    t1 = time.time()
    generator.generate(checklist_data, output_path)
    t2 = time.time()
    print(f"   ✅ {t2 - t1:.2f} сек")

    # Размер файла
    size_mb = Path(output_path).stat().st_size / (1024 * 1024)
    print(f"   📁 Размер файла: {size_mb:.2f} MB")

    # Очистка
    Path(output_path).unlink()

    print("\n" + "=" * 80)
    print(f"📊 ОБЩЕЕ ВРЕМЯ: {t2 - t1:.2f} сек")
    print("=" * 80)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("file", help="Путь к чек-листу")
    args = parser.parse_args()

    profile_generation(args.file)