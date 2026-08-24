#!/usr/bin/env python3
"""
Инициализация базы данных: создание таблиц и категорий
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import engine, SessionLocal
from db.models import Base, Category, RoomCategory, CleaningMethodOrder, InventoryColor, CleaningTechnique


def create_tables():
    """Создание таблиц (без удаления существующих)"""
    print("🔄 Создание таблиц (если не существуют)...")
    Base.metadata.create_all(bind=engine)
    print("✅ Таблицы проверены/созданы")


def seed_categories():
    """Заполнение категорий объектов (15 штук)"""
    print("🔄 Проверка категорий объектов...")

    session = SessionLocal()

    existing_count = session.query(Category).count()
    if existing_count > 0:
        print(f"✅ Категории объектов уже существуют ({existing_count} шт.), пропускаем")
        session.close()
        return

    categories_data = [
        ("Поверхности", 1),
        ("Сантехническое оборудование", 2),
        ("Санитарный пост", 3),
        ("Мебель", 4),
        ("Офисная техника", 5),
        ("Многоразовые резиновые СИЗ", 6),
        ("Бытовая техника", 7),
        ("Инвентарь, посуда и т.д.", 8),
        ("Моечный, уборочный инвентарь и оборудование", 9),
        ("Посудомоечное оборудование", 10),
        ("Холодильное оборудование", 11),
        ("Дозирующее оборудование", 12),
        ("Тепловое оборудование", 13),
        ("Технологическое оборудование", 14),
        ("Упаковочное оборудование", 15),
        ("Контактные поверхности", 16),
    ]

    for name, sort_order in categories_data:
        category = Category(name=name, sort_order=sort_order)
        session.add(category)
        print(f"  ➕ {name}")

    session.commit()
    session.close()
    print(f"✅ Добавлено {len(categories_data)} категорий")


def seed_room_categories():
    """Заполнение категорий помещений (актуальный список)"""
    print("🔄 Проверка категорий помещений...")

    session = SessionLocal()

    existing_count = session.query(RoomCategory).count()
    if existing_count > 0:
        print(f"✅ Категории помещений уже существуют ({existing_count} шт.), пропускаем")
        session.close()
        return

    # Все категории в нижнем регистре для единообразия
    room_categories = [
        "общего назначения",
        "производственное",
        "бытовое",
        "складское",
        "санитарное",
        "вспомогательное",
        "моечное",
        "техническое",
        "офисное",
    ]

    for name in room_categories:
        rc = RoomCategory(name=name)
        session.add(rc)
        print(f"  ➕ {name}")

    session.commit()
    session.close()
    print(f"✅ Добавлено {len(room_categories)} категорий помещений")


def seed_cleaning_method_order():
    """Заполнение порядка способов обработки (CleaningMethodOrder)"""
    print("🔄 Проверка порядка способов обработки...")

    session = SessionLocal()

    existing_count = session.query(CleaningMethodOrder).count()
    if existing_count > 0:
        print(f"✅ Порядок способов обработки уже задан ({existing_count} методов), пропускаем")
        session.close()
        return

    methods = [
        ("прочистка", 1),
        ("очистка", 2),
        ("очистка (обеспыливание поверхностей)", 3),
        ("удаление маркировки", 4),
        ("мойка", 5),
        ("мойка жаропрочного стекла", 6),
        ("ополаскивание", 7),
        ("стерилизация", 8),
        ("дезинфекция", 9),
        ("машинная стирка", 10),
    ]

    for name, sort_order in methods:
        m = CleaningMethodOrder(method_name=name, sort_order=sort_order)
        session.add(m)
        print(f"  ➕ {name}: {sort_order}")

    session.commit()
    session.close()
    print(f"✅ Добавлено {len(methods)} методов")


# ===================== ДОБАВЛЕНО: SEED-ЗАПОЛНЕНИЕ ЦВЕТОВ ИНВЕНТАРЯ =====================

def seed_inventory_colors():
    """Заполняет таблицу inventory_colors стандартными цветами, если она пуста."""
    print("🔄 Проверка цветов инвентаря...")

    session = SessionLocal()
    try:
        existing_count = session.query(InventoryColor).count()
        if existing_count > 0:
            print(f"✅ Цвета инвентаря уже существуют ({existing_count} шт.), пропускаем")
            return

        standard_colors = [
            {"name": "чёрный", "hex": "#000000"},
            {"name": "красный", "hex": "#FF0000"},
            {"name": "жёлтый", "hex": "#FFFF00"},
            {"name": "зелёный", "hex": "#008000"},
            {"name": "синий", "hex": "#0000FF"},
            {"name": "голубой", "hex": "#00FFFF"},
            {"name": "белый", "hex": "#FFFFFF"},
            {"name": "коричневый", "hex": "#8B4513"},
            {"name": "лаймовый", "hex": "#32CD32"},
            {"name": "оранжевый", "hex": "#FFA500"},
            {"name": "розовый", "hex": "#FFC0CB"},
            {"name": "серый", "hex": "#808080"},
            {"name": "сиреневый", "hex": "#C8A2C8"},
            {"name": "фиолетовый", "hex": "#800080"},
        ]

        for color in standard_colors:
            session.add(InventoryColor(name=color["name"], hex_color=color["hex"]))
            print(f"  ➕ {color['name']} – {color['hex']}")

        session.commit()
        print(f"✅ Добавлено {len(standard_colors)} цветов инвентаря")
    finally:
        session.close()


# ===================== ДОБАВЛЕНО: SEED-ЗАПОЛНЕНИЕ МЕТОДОВ УБОРКИ =====================

def seed_cleaning_techniques():
    """Заполняет таблицу cleaning_techniques стандартными методами уборки, если она пуста."""
    print("🔄 Проверка методов уборки...")

    session = SessionLocal()
    try:
        existing_count = session.query(CleaningTechnique).count()
        if existing_count > 0:
            print(f"✅ Методы уборки уже существуют ({existing_count} шт.), пропускаем")
            return

        standard_techniques = [
            "протирка",
            "обработка щёткой",
            "распыление",
            "погружение",
            "замачивание",
            "орошение",
            "обработка ветошью",
            "обработка губкой",
            "обработка мопом",
            "обработка скребком",
            "обработка паром",
            "обработка губкой с абразивным слоем",
        ]

        for name in standard_techniques:
            session.add(CleaningTechnique(name=name))
            print(f"  ➕ {name}")

        session.commit()
        print(f"✅ Добавлено {len(standard_techniques)} методов уборки")
    finally:
        session.close()


def main():
    print("\n🔧 ИНИЦИАЛИЗАЦИЯ БАЗЫ ДАННЫХ")
    print("=" * 50)

    create_tables()
    seed_categories()
    seed_room_categories()
    seed_cleaning_method_order()
    seed_inventory_colors()
    seed_cleaning_techniques()  # <-- ДОБАВЛЕН ВЫЗОВ

    print(f"\n📁 Файл БД: {engine.url.database}")
    print("🎉 Готово!")


if __name__ == "__main__":
    main()