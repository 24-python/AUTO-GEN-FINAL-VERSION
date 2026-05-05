"""
Вывод реальной структуры таблиц базы данных
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from db.database import engine
from sqlalchemy import inspect

inspector = inspect(engine)

for table_name in inspector.get_table_names():
    print(f"\n{'=' * 60}")
    print(f"📋 Таблица: {table_name}")
    print(f"{'=' * 60}")

    # Колонки
    print(f"\nКолонки:")
    for col in inspector.get_columns(table_name):
        nullable = "NULL" if col['nullable'] else "NOT NULL"
        default = f" DEFAULT {col['default']}" if col['default'] else ""
        pk = " [PK]" if col.get('primary_key') else ""
        print(f"  {col['name']:<30} {str(col['type']):<20} {nullable}{default}{pk}")

    # Foreign Keys
    fks = inspector.get_foreign_keys(table_name)
    if fks:
        print(f"\nForeign Keys:")
        for fk in fks:
            cols = ', '.join(fk['constrained_columns'])
            ref = f"{fk['referred_table']}.{', '.join(fk['referred_columns'])}"
            print(f"  {cols} → {ref}")

    # Индексы
    indexes = inspector.get_indexes(table_name)
    if indexes:
        print(f"\nИндексы:")
        for idx in indexes:
            cols = ', '.join(idx['column_names'])
            unique = "UNIQUE " if idx['unique'] else ""
            print(f"  {unique}{idx['name']}: {cols}")

print(f"\n{'=' * 60}")
print(f"Всего таблиц: {len(inspector.get_table_names())}")