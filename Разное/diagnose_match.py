# diagnose_match.py
"""
Диагностика сопоставления Excel и БД.
Показывает первые объекты из Excel и результаты поиска в БД.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from match_schedule import read_excel, find_object_in_db, split_compound_objects, normalize_name
from db.database import SessionLocal
from db.models import Object


def main():
    excel_file = "../периодичность обработки всех помещений.xlsx"

    print(f"📖 Чтение Excel: {excel_file}")
    data = read_excel(excel_file)
    print(f"Категорий в Excel: {list(data.keys())}")
    print()

    # Показываем Бытовое
    byt = data.get('Бытовое', {})
    print(f"Объектов в категории 'Бытовое': {len(byt)}")
    print()

    for i, (name, levels) in enumerate(byt.items()):
        if i >= 5:
            break

        print(f"Excel: \"{name}\"")
        print(f"  Уровней: {len(levels)}")
        for l in levels:
            w = l['frequency_wash'] or '-'
            d = l['frequency_des'] or '-'
            e = l['executor'] or '-'
            print(f"    {l['maintenance_type']}: мойка={w}, дез={d}, исп={e}")

        # Поиск в БД
        session = SessionLocal()
        found = find_object_in_db(session, name)
        session.close()

        if found:
            for obj in found:
                print(f"  ✅ Найден: {obj.display_name} (id={obj.id})")
        else:
            print(f"  ❌ Не найден как целое")
            parts = split_compound_objects(name)
            print(f"  Составные части: {parts}")
            for p in parts:
                session = SessionLocal()
                f = find_object_in_db(session, p)
                session.close()
                if f:
                    for obj in f:
                        print(f"    ✅ \"{p}\" → {obj.display_name} (id={obj.id})")
                else:
                    print(f"    ❌ \"{p}\" → НЕ НАЙДЕНА")
        print()


if __name__ == "__main__":
    main()