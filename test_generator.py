#!/usr/bin/env python3
"""
test_generator.py - Тестирование генератора технологических карт
"""

import sys
from pathlib import Path

# Добавляем корень проекта в путь
sys.path.insert(0, str(Path(__file__).parent))

from parser.xml_parser import parse_checklist
from generator.docx_generator import TechCardGenerator


def find_checklists():
    """Ищет все .docx файлы в папке проекта и uploads"""
    project_root = Path(__file__).parent
    checklist_files = []

    # Ищем в корне
    for f in project_root.glob("*.docx"):
        if not f.name.startswith("~") and "tech_card" not in f.name.lower():
            checklist_files.append(f)

    # Ищем в папке uploads
    uploads_dir = project_root / "uploads"
    if uploads_dir.exists():
        for f in uploads_dir.glob("*.docx"):
            if not f.name.startswith("~"):
                checklist_files.append(f)

    return checklist_files


def test_generator(checklist_path: Path, output_path: Path = None):
    """Тестирует генератор на конкретном чек-листе"""
    if output_path is None:
        output_dir = Path("test_output")
        output_dir.mkdir(exist_ok=True)
        output_path = output_dir / f"{checklist_path.stem}_tech_card.docx"

    print("=" * 60)
    print("ТЕСТИРОВАНИЕ ГЕНЕРАТОРА ТЕХНОЛОГИЧЕСКИХ КАРТ")
    print("=" * 60)

    # ШАГ 1: Парсинг чек-листа
    print(f"\n📄 ШАГ 1: Парсинг чек-листа")
    print(f"   Файл: {checklist_path.name}")

    try:
        checklist_data = parse_checklist(str(checklist_path))
        print(f"   ✅ Предприятие: {checklist_data.enterprise or 'не указано'}")
        print(f"   ✅ Помещение: {checklist_data.room_name or 'не указано'}")
        checked_items = checklist_data.get_checked_items()
        print(f"   ✅ Отмечено объектов: {len(checked_items)}")

        if checked_items:
            print(f"\n   📋 Первые {min(5, len(checked_items))} объектов:")
            for i, item in enumerate(checked_items[:5]):
                print(f"      {i + 1}. {item.name} [{item.category.value}]")
            if len(checked_items) > 5:
                print(f"      ... и ещё {len(checked_items) - 5}")

    except Exception as e:
        print(f"   ❌ Ошибка парсинга: {e}")
        import traceback
        traceback.print_exc()
        return False

    # ШАГ 2: Генерация техкарты
    print(f"\n📝 ШАГ 2: Генерация технологической карты")
    print(f"   Шаблон: tech_card_templates/шаблон.docx")
    print(f"   Выходной файл: {output_path}")

    try:
        generator = TechCardGenerator()
        result_path = generator.generate(checklist_data, str(output_path))
        print(f"   ✅ Генерация завершена")
        print(f"   📁 Файл сохранен: {result_path}")

    except FileNotFoundError as e:
        print(f"   ❌ Шаблон не найден: {e}")
        print(f"   📌 Убедитесь, что файл шаблона существует:")
        print(f"      tech_card_templates/шаблон.docx")
        return False

    except Exception as e:
        print(f"   ❌ Ошибка генерации: {e}")
        import traceback
        traceback.print_exc()
        return False

    # ШАГ 3: Проверка результата
    print(f"\n🔍 ШАГ 3: Проверка результата")

    result_file = Path(result_path)
    if result_file.exists():
        size_kb = result_file.stat().st_size / 1024
        print(f"   ✅ Файл создан (размер: {size_kb:.1f} KB)")
        print(f"   📁 Откройте файл для визуальной проверки:")
        print(f"      {result_file.absolute()}")
    else:
        print(f"   ❌ Файл не создан")
        return False

    print("\n" + "=" * 60)
    print("✅ ТЕСТ ЗАВЕРШЕН")
    print("=" * 60)

    return True


def main():
    print("🔧 ТЕСТИРОВАНИЕ ГЕНЕРАТОРА ТЕХКАРТ")
    print("=" * 60)

    checklists = find_checklists()

    if not checklists:
        print("\n❌ Не найдено файлов чек-листов (.docx)")
        print("\n📌 Что делать:")
        print("   1. Положите файл чек-листа в корень проекта")
        print("   2. Или создайте папку 'uploads' и положите файл туда")
        return

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

    test_generator(selected)


if __name__ == "__main__":
    main()