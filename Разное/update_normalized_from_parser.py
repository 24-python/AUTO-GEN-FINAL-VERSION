#!/usr/bin/env python3
"""
fix_machine_normalized.py

Исправляет normalized_name для "машина для резки конд. изделий".
Удаляет точку, чтобы совпадало с парсером.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from db.database import SessionLocal
from db.models import Object


def main():
    session = SessionLocal()

    obj = session.query(Object).filter(Object.id == 155).first()

    if not obj:
        print("❌ Объект ID=155 не найден!")
        session.close()
        return

    print("=" * 80)
    print("🔧 ИСПРАВЛЕНИЕ NORMALIZED_NAME")
    print("=" * 80)

    print(f"\n   Объект: {obj.display_name}")
    print(f"   Старое normalized_name: '{obj.normalized_name}'")

    # Новое имя — БЕЗ ТОЧКИ
    new_norm = "машина для резки конд изделий"

    print(f"   Новое normalized_name:  '{new_norm}'")

    # Проверяем, нет ли уже объекта с таким именем
    existing = session.query(Object).filter(
        Object.normalized_name == new_norm,
        Object.id != obj.id
    ).first()

    if existing:
        print(f"\n   ⚠️ Найден дубликат (ID={existing.id})")
        print(f"      Переносим инструкции...")
        for instr in existing.instructions:
            instr.object_id = obj.id
        session.delete(existing)
        print(f"      ✅ Дубликат удалён")

    obj.normalized_name = new_norm
    session.commit()

    print(f"\n   ✅ ИСПРАВЛЕНО!")

    session.close()

    print("\n" + "=" * 80)
    print("✅ ГОТОВО")
    print("=" * 80)


if __name__ == "__main__":
    main()