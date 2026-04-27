#!/usr/bin/env python3
"""
Инициализация базы данных: создание таблиц и категорий
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from db.database import engine, SessionLocal
from db.models import Base, Category, RoomCategory


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
    ]

    for name, sort_order in categories_data:
        category = Category(name=name, sort_order=sort_order)
        session.add(category)
        print(f"  ➕ {name}")

    session.commit()
    session.close()
    print(f"✅ Добавлено {len(categories_data)} категорий")


def seed_room_categories():
    """Заполнение категорий помещений"""
    print("🔄 Проверка категорий помещений...")

    session = SessionLocal()

    existing_count = session.query(RoomCategory).count()
    if existing_count > 0:
        print(f"✅ Категории помещений уже существуют ({existing_count} шт.), пропускаем")
        session.close()
        return

    room_categories = [
        "Производственное",
        "Складское",
        "Инженерное",
        "Вспомогательное",
    ]

    for name in room_categories:
        rc = RoomCategory(name=name)
        session.add(rc)
        print(f"  ➕ {name}")

    session.commit()
    session.close()
    print(f"✅ Добавлено {len(room_categories)} категорий помещений")


def main():
    print("\n🔧 ИНИЦИАЛИЗАЦИЯ БАЗЫ ДАННЫХ")
    print("=" * 50)

    create_tables()
    seed_categories()
    seed_room_categories()

    print(f"\n📁 Файл БД: {engine.url.database}")
    print("🎉 Готово!")


if __name__ == "__main__":
    main()