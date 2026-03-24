#!/usr/bin/env python3
"""
Импорт данных из JSON в базу данных
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import SessionLocal
from db.models import Category, Object, Instruction


def import_from_json(json_path):
    """
    Импортирует данные из JSON в БД

    Правила:
    - Категории: сопоставляются по имени (должны существовать)
    - Объекты: если name уже существует → обновляем, иначе создаем
    - Инструкции: привязываем к object_id
    """

    # Загружаем JSON
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

    # Импорт объектов
    for obj_data in data["objects"]:
        category_id = categories.get(obj_data["category"])
        if not category_id:
            print(f"  ❌ Категория '{obj_data['category']}' не найдена")
            stats["errors"] += 1
            continue

        # Проверяем, существует ли объект
        existing = session.query(Object).filter_by(name=obj_data["name"]).first()

        if existing:
            # Обновляем
            existing.category_id = category_id
            existing.base_name = obj_data["base_name"]
            existing.modifier = obj_data["modifier"]
            existing.sort_priority = obj_data.get("sort_priority", 0)
            stats["objects_updated"] += 1
            print(f"  🔄 Обновлен: {obj_data['name']}")
        else:
            # Создаем
            obj = Object(
                name=obj_data["name"],
                base_name=obj_data["base_name"],
                modifier=obj_data["modifier"],
                category_id=category_id,
                sort_priority=obj_data.get("sort_priority", 0)
            )
            session.add(obj)
            stats["objects_created"] += 1
            print(f"  ➕ Создан: {obj_data['name']}")

    session.flush()  # чтобы получить id объектов

    print("\n📥 ИМПОРТ ИНСТРУКЦИЙ")
    print("-" * 40)

    # Словарь объектов {name: id}
    objects = {obj.name: obj.id for obj in session.query(Object).all()}

    # Импорт инструкций
    for instr_data in data["instructions"]:
        object_id = objects.get(instr_data["object_name"])
        if not object_id:
            print(f"  ❌ Объект '{instr_data['object_name']}' не найден")
            stats["errors"] += 1
            continue

        # Создаем инструкцию
        instruction = Instruction(
            object_id=object_id,
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
    main()#!/usr/bin/env python3
"""
Импорт данных из JSON в базу данных
"""

import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import SessionLocal
from db.models import Category, Object, Instruction


def import_from_json(json_path):
    """
    Импортирует данные из JSON в БД

    Правила:
    - Категории: сопоставляются по имени (с маппингом)
    - Объекты: если name уже существует → обновляем, иначе создаем
    - Инструкции: привязываем к object_id
    """

    # Загружаем JSON
    with open(json_path, 'r', encoding='utf-8') as f:
        data = json.load(f)

    session = SessionLocal()

    # Словарь категорий {name: id}
    categories = {cat.name: cat.id for cat in session.query(Category).all()}

    # МАППИНГ для проблемных категорий
    category_mapping = {
        "просеиватели мука сахар": "Технологическое оборудование",
        "Моечный, уборочный инвентарь": "Моечный, уборочный инвентарь и оборудование"
    }

    stats = {
        "objects_created": 0,
        "objects_updated": 0,
        "instructions_created": 0,
        "errors": 0
    }

    print("\n📥 ИМПОРТ ОБЪЕКТОВ")
    print("-" * 40)

    # Импорт объектов
    for obj_data in data["objects"]:
        category_name = obj_data["category"]

        # Применяем маппинг
        if category_name in category_mapping:
            category_name = category_mapping[category_name]

        category_id = categories.get(category_name)
        if not category_id:
            print(f"  ❌ Категория '{obj_data['category']}' не найдена (маппинг: {category_name})")
            stats["errors"] += 1
            continue

        # Проверяем, существует ли объект
        existing = session.query(Object).filter_by(name=obj_data["name"]).first()

        if existing:
            # Обновляем
            existing.category_id = category_id
            existing.base_name = obj_data["base_name"]
            existing.modifier = obj_data["modifier"]
            existing.sort_priority = obj_data.get("sort_priority", 0)
            stats["objects_updated"] += 1
            print(f"  🔄 Обновлен: {obj_data['name']}")
        else:
            # Создаем
            obj = Object(
                name=obj_data["name"],
                base_name=obj_data["base_name"],
                modifier=obj_data["modifier"],
                category_id=category_id,
                sort_priority=obj_data.get("sort_priority", 0)
            )
            session.add(obj)
            stats["objects_created"] += 1
            print(f"  ➕ Создан: {obj_data['name']}")

    session.flush()  # чтобы получить id объектов

    print("\n📥 ИМПОРТ ИНСТРУКЦИЙ")
    print("-" * 40)

    # Словарь объектов {name: id}
    objects = {obj.name: obj.id for obj in session.query(Object).all()}

    # Специальный маппинг для объектов с разными названиями в JSON
    object_name_mapping = {
        "мопы": "мопы",
        "ветошь": "ветошь",
        "щётки": "щётки",
        "сгоны": "сгоны",
        "ведра": "ведра",
        "АВД": "АВД",
        "пылесосы": "пылесосы",
        "поломоечная машина": "поломоечная машина",
        "солодоварки ферментаторы": "солодоварки",
        "рентгеновские системы контроля": "рентгеновские системы контроля",
        "вакуумные роторные шприцы": "вакуумные роторные шприцы",
        "овощерезки": "овощерезки",
        "овощечистки": "овощечистки",
        "измельчители": "измельчители",
        "протирочные машины": "протирочные машины",
        "картофелечистки": "картофелечистки",
        "депозитор волюметрический": "депозитор волюметрический",
        "металлодетектор": "металлодетектор",
    }

    # Импорт инструкций
    for instr_data in data["instructions"]:
        object_name = instr_data["object_name"]

        # Применяем маппинг если нужно
        if object_name in object_name_mapping:
            object_name = object_name_mapping[object_name]

        object_id = objects.get(object_name)
        if not object_id:
            print(f"  ❌ Объект '{instr_data['object_name']}' не найден")
            stats["errors"] += 1
            continue

        # Создаем инструкцию
        instruction = Instruction(
            object_id=object_id,
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