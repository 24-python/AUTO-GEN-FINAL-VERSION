#!/usr/bin/env python3
"""
debug_fetch_instructions_interactive.py - Отладка поиска инструкций в БД с выбором файла
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from parser.xml_parser import parse_checklist
from db.database import SessionLocal
from generator.mapper import find_object_and_instructions


def find_checklists():
    """Ищет все .docx файлы в папке проекта и uploads"""
    project_root = Path(__file__).parent
    checklist_files = []

    for f in project_root.glob("*.docx"):
        if not f.name.startswith("~") and "tech_card" not in f.name.lower():
            checklist_files.append(f)

    uploads_dir = project_root / "uploads"
    if uploads_dir.exists():
        for f in uploads_dir.glob("*.docx"):
            if not f.name.startswith("~"):
                checklist_files.append(f)

    return checklist_files


def debug_fetch_instructions(checklist_path: str):
    """Отлаживает поиск инструкций для каждого объекта из чек-листа"""

    print("=" * 70)
    print("ОТЛАДКА ПОИСКА ИНСТРУКЦИЙ В БАЗЕ ДАННЫХ")
    print("=" * 70)

    # 1. Парсим чек-лист
    print(f"\n📄 Парсинг чек-листа: {checklist_path}")
    print("-" * 50)

    try:
        checklist_data = parse_checklist(checklist_path)
    except Exception as e:
        print(f"❌ Ошибка парсинга: {e}")
        return

    checked_items = checklist_data.get_checked_items()
    print(f"\n✅ Отмечено объектов: {len(checked_items)}")

    if not checked_items:
        print("❌ Нет отмеченных объектов")
        return

    # 2. Для каждого объекта ищем в БД
    print("\n" + "=" * 70)
    print("ПОИСК КАЖДОГО ОБЪЕКТА В БАЗЕ ДАННЫХ")
    print("=" * 70)

    session = SessionLocal()

    results = {
        "found": [],
        "not_found": [],
        "found_no_instructions": []
    }

    for idx, item in enumerate(checked_items, 1):
        print(f"\n[{idx}] Объект: '{item.name}'")
        print(f"    Категория (Enum): {item.category.value}")

        obj, instructions = find_object_and_instructions(session, item.name, item.category)

        if obj:
            instr_count = len(instructions)
            print(f"    ✅ НАЙДЕН в БД: id={obj.id}, name='{obj.name}'")
            print(f"    📋 Инструкций: {instr_count}")

            if instr_count > 0:
                results["found"].append((item.name, obj.name, instr_count))
                for i, instr in enumerate(instructions[:2]):
                    print(
                        f"       - {i + 1}. {instr.cleaning_method or 'нет метода'} | {instr.product_name or 'нет средства'}")
                if instr_count > 2:
                    print(f"       ... и ещё {instr_count - 2}")
            else:
                results["found_no_instructions"].append((item.name, obj.name))
                print(f"    ⚠️ НЕТ ИНСТРУКЦИЙ для этого объекта")
        else:
            print(f"    ❌ НЕ НАЙДЕН в БД")
            results["not_found"].append(item.name)

    session.close()

    # 3. Вывод статистики
    print("\n" + "=" * 70)
    print("СТАТИСТИКА")
    print("=" * 70)
    print(f"✅ Найдено с инструкциями: {len(results['found'])}")
    print(f"⚠️ Найдено без инструкций: {len(results['found_no_instructions'])}")
    print(f"❌ Не найдено в БД: {len(results['not_found'])}")

    if results['not_found']:
        print("\n❌ СПИСОК ОБЪЕКТОВ, НЕ НАЙДЕННЫХ В БД:")
        for name in results['not_found']:
            print(f"   - '{name}'")

    if results['found_no_instructions']:
        print("\n⚠️ СПИСОК ОБЪЕКТОВ, НАЙДЕННЫХ БЕЗ ИНСТРУКЦИЙ:")
        for orig_name, db_name in results['found_no_instructions']:
            print(f"   - '{orig_name}' → в БД: '{db_name}' (нет инструкций)")

    return results


def main():
    checklists = find_checklists()

    if not checklists:
        print("❌ Не найдено файлов чек-листов (.docx)")
        print("📌 Положите файл чек-листа в корень проекта или в папку uploads/")
        return

    print("\n📋 Найденные чек-листы:")
    print("-" * 40)
    for i, f in enumerate(checklists, 1):
        size_kb = f.stat().st_size / 1024
        print(f"   {i}. {f.name} ({size_kb:.1f} KB)")
    print("-" * 40)
    print("   0. Выйти")

    while True:
        try:
            choice = input("\nВыберите номер чек-листа: ").strip()
            if choice == "0":
                print("Выход.")
                return
            idx = int(choice) - 1
            if 0 <= idx < len(checklists):
                selected = checklists[idx]
                break
            else:
                print(f"❌ Введите число от 1 до {len(checklists)}")
        except ValueError:
            print("❌ Введите число")
        except KeyboardInterrupt:
            print("\nВыход.")
            return

    debug_fetch_instructions(str(selected))


if __name__ == "__main__":
    main()