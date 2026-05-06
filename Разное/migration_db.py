import sys; from pathlib import Path; sys.path.insert(0, '..')
from db.database import engine
with engine.connect() as conn:
    try:
        conn.exec_driver_sql('ALTER TABLE instructions ADD COLUMN maintenance_type VARCHAR(50)')
        print('✅ Колонка maintenance_type добавлена в instructions')
    except Exception as e:
        if 'duplicate column' in str(e).lower():
            print('⚠️ Колонка уже существует')
        else:
            print(f'❌ Ошибка: {e}')