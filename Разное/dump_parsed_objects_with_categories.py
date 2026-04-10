#!/usr/bin/env python3
"""
dump_parsed_objects_with_categories.py
Собирает все объекты (нормализованные имена) из чек-листов в корне проекта,
группирует по категориям и сохраняет в текстовый файл.
"""

import sys
from pathlib import Path
from collections import defaultdict

# Добавляем корень проекта в путь
sys.path.insert(0, str(Path(__file__).parent))

from parser.xml_parser import parse_checklist


def find_checklists(root_dir: Path) -> list:
    """Находит все .docx файлы в корне (не рекурсивно, без подпапок)."""
    checklists = []
    for pattern in ['*.docx', '*.DOCX']:
        for f in root_dir.glob(pattern):
            # Игнорируем временные файлы и уже сгенерированные техкарты
            if not f.name.startswith('~') and 'tech_card' not in f.name.lower():
                checklists.append(f)
    return checklists


def collect_objects_from_checklists(checklist_paths: list) -> dict:
    """
    Парсит чек-листы и возвращает словарь:
    {category_name: set(normalized_object_names)}
    """
    category_objects = defaultdict(set)

    for idx, path in enumerate(checklist_paths, 1):
        print(f"[{idx}/{len(checklist_paths)}] Обработка: {path.name}")
        try:
            data = parse_checklist(str(path))
            for item in data.get_checked_items():
                category_objects[item.category.value].add(item.name)
        except Exception as e:
            print(f"  Ошибка: {e}")

    return category_objects


def save_to_file(category_objects: dict, output_path: Path):
    """Сохраняет категории и объекты в текстовый файл."""
    with open(output_path, 'w', encoding='utf-8') as f:
        for category, objects in sorted(category_objects.items()):
            f.write(f"{category}:\n")
            for obj in sorted(objects):
                f.write(f"  {obj}\n")
            f.write("\n")  # пустая строка между категориями


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Сбор нормализованных объектов с категориями из чек-листов")
    parser.add_argument("--input-dir", default=".", help="Папка с чек-листами (по умолчанию текущая)")
    parser.add_argument("--output", "-o", default="parsed_objects_with_categories.txt", help="Выходной файл")
    args = parser.parse_args()

    root_dir = Path(args.input_dir).resolve()
    if not root_dir.exists():
        print(f"❌ Папка не найдена: {root_dir}")
        return

    print(f"🔍 Поиск .docx файлов в: {root_dir}")
    checklists = find_checklists(root_dir)

    if not checklists:
        print("❌ Не найдено ни одного .docx файла (кроме техкарт).")
        return

    print(f"✅ Найдено чек-листов: {len(checklists)}")

    print("\n📄 Сбор объектов...")
    category_objects = collect_objects_from_checklists(checklists)

    total_objects = sum(len(v) for v in category_objects.values())
    print(f"\n📊 Категорий: {len(category_objects)}")
    print(f"📊 Уникальных объектов: {total_objects}")

    output_path = Path(args.output)
    save_to_file(category_objects, output_path)
    print(f"\n✅ Результат сохранён в: {output_path}")


if __name__ == "__main__":
    main()