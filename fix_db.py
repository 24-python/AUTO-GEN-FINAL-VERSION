#!/usr/bin/env python3
"""
fix_db.py – скрипт для добавления недостающих колонок в таблицу instructions
и создания таблицы inventory_colors, если она отсутствует.
Запускать один раз при возникновении ошибок типа "no such column".
"""

import sqlite3
import sys
from pathlib import Path

# Путь к БД (берём из database.py или указываем явно)
DB_PATH = Path("tech_cards.db")  # по умолчанию в корне проекта

# Если БД лежит в другой папке, можно указать через переменную окружения DB_DIR
# или раскомментировать строку ниже и указать путь вручную:
# DB_PATH = Path("db_storage") / "tech_cards.db"

def column_exists(cursor, table_name, column_name):
    """Проверяет, существует ли колонка в таблице."""
    cursor.execute(f"PRAGMA table_info({table_name})")
    columns = [row[1] for row in cursor.fetchall()]
    return column_name in columns

def table_exists(cursor, table_name):
    """Проверяет, существует ли таблица."""
    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
    return cursor.fetchone() is not None

def main():
    if not DB_PATH.exists():
        print(f"❌ Файл БД не найден: {DB_PATH}")
        print("   Убедитесь, что путь к БД указан правильно.")
        print("   Если БД находится в другом месте, измените переменную DB_PATH в скрипте.")
        sys.exit(1)

    conn = sqlite3.connect(str(DB_PATH))
    cursor = conn.cursor()

    print("🔍 Проверка структуры базы данных...")

    # ----- Таблица instructions -----
    print("\n📋 Таблица instructions:")
    # Колонка enterprise
    if not column_exists(cursor, "instructions", "enterprise"):
        print("   ➕ Добавляем колонку enterprise...")
        cursor.execute("ALTER TABLE instructions ADD COLUMN enterprise VARCHAR(200) NULL")
        print("   ✅ enterprise добавлена")
    else:
        print("   ✅ enterprise уже существует")

    # Колонка subgroup
    if not column_exists(cursor, "instructions", "subgroup"):
        print("   ➕ Добавляем колонку subgroup...")
        cursor.execute("ALTER TABLE instructions ADD COLUMN subgroup VARCHAR(50) NULL")
        print("   ✅ subgroup добавлена")
    else:
        print("   ✅ subgroup уже существует")

    # Колонка room_name (самая важная – её не хватало)
    if not column_exists(cursor, "instructions", "room_name"):
        print("   ➕ Добавляем колонку room_name...")
        cursor.execute("ALTER TABLE instructions ADD COLUMN room_name VARCHAR(200) NULL")
        print("   ✅ room_name добавлена")
    else:
        print("   ✅ room_name уже существует")

    # ----- Таблица inventory_colors -----
    print("\n🎨 Таблица inventory_colors:")
    if not table_exists(cursor, "inventory_colors"):
        print("   ➕ Создаём таблицу inventory_colors...")
        cursor.execute("""
            CREATE TABLE inventory_colors (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name VARCHAR(50) NOT NULL UNIQUE,
                hex_color VARCHAR(7) NOT NULL
            )
        """)
        print("   ✅ Таблица создана")

        # Добавляем стандартные цвета
        print("   ➕ Добавляем стандартные цвета...")
        colors = [
            ('чёрный', '#000000'),
            ('красный', '#FF0000'),
            ('жёлтый', '#FFFF00'),
            ('зелёный', '#008000'),
            ('синий', '#0000FF'),
            ('голубой', '#00FFFF'),
            ('белый', '#FFFFFF'),
            ('коричневый', '#8B4513'),
            ('лаймовый', '#32CD32'),
            ('оранжевый', '#FFA500'),
            ('розовый', '#FFC0CB'),
            ('серый', '#808080'),
            ('сиреневый', '#C8A2C8'),
            ('фиолетовый', '#800080'),
        ]
        cursor.executemany("INSERT INTO inventory_colors (name, hex_color) VALUES (?, ?)", colors)
        print(f"   ✅ Добавлено {len(colors)} цветов")
    else:
        # Проверяем, есть ли цвета
        cursor.execute("SELECT COUNT(*) FROM inventory_colors")
        count = cursor.fetchone()[0]
        if count == 0:
            print("   ⚠️ Таблица существует, но пуста. Добавляем стандартные цвета...")
            colors = [
                ('чёрный', '#000000'),
                ('красный', '#FF0000'),
                ('жёлтый', '#FFFF00'),
                ('зелёный', '#008000'),
                ('синий', '#0000FF'),
                ('голубой', '#00FFFF'),
                ('белый', '#FFFFFF'),
                ('коричневый', '#8B4513'),
                ('лаймовый', '#32CD32'),
                ('оранжевый', '#FFA500'),
                ('розовый', '#FFC0CB'),
                ('серый', '#808080'),
                ('сиреневый', '#C8A2C8'),
                ('фиолетовый', '#800080'),
            ]
            cursor.executemany("INSERT OR IGNORE INTO inventory_colors (name, hex_color) VALUES (?, ?)", colors)
            print(f"   ✅ Добавлено {len(colors)} цветов")
        else:
            print(f"   ✅ Таблица содержит {count} цветов")

    conn.commit()
    conn.close()

    print("\n✅ Все изменения применены. Теперь перезапустите приложение.")
    print("   Ошибка 'no such column: instructions.room_name' должна исчезнуть.")


if __name__ == "__main__":
    main()