import sqlite3
from pathlib import Path

DB_PATH = Path("tech_cards.db")  # или ваш путь к БД

conn = sqlite3.connect(str(DB_PATH))
cursor = conn.cursor()

# Добавляем колонку enterprise
try:
    cursor.execute("ALTER TABLE instructions ADD COLUMN enterprise VARCHAR(200) NULL")
    print("✅ Колонка enterprise добавлена")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✅ Колонка enterprise уже существует")
    else:
        print(f"⚠️ Ошибка: {e}")

# Добавляем колонку subgroup
try:
    cursor.execute("ALTER TABLE instructions ADD COLUMN subgroup VARCHAR(50) NULL")
    print("✅ Колонка subgroup добавлена")
except sqlite3.OperationalError as e:
    if "duplicate column name" in str(e):
        print("✅ Колонка subgroup уже существует")
    else:
        print(f"⚠️ Ошибка: {e}")

# Создаём таблицу inventory_colors, если нет
cursor.execute("""
    CREATE TABLE IF NOT EXISTS inventory_colors (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        name VARCHAR(50) NOT NULL UNIQUE,
        hex_color VARCHAR(7) NOT NULL
    )
""")
print("✅ Таблица inventory_colors создана (если не существовала)")

# Добавляем стандартные цвета
cursor.execute("SELECT COUNT(*) FROM inventory_colors")
if cursor.fetchone()[0] == 0:
    colors = [
        ('чёрный', '#000000'),
        ('красный', '#FF0000'),
        ('жёлтый', '#FFFF00'),
        ('зелёный', '#008000'),
        ('синий', '#0000FF'),
        ('голубой', '#00FFFF'),
    ]
    cursor.executemany("INSERT INTO inventory_colors (name, hex_color) VALUES (?, ?)", colors)
    print("✅ Стандартные цвета добавлены")
else:
    print("✅ Цвета уже есть в таблице")

conn.commit()
conn.close()
print("🎉 Готово! Перезапустите приложение.")