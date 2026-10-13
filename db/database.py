"""
Подключение к базе данных SQLite
"""

import os
from sqlalchemy import create_engine, event
from sqlalchemy.orm import sessionmaker, scoped_session
from pathlib import Path

# Путь к папке с БД (по умолчанию – корень проекта)
DB_DIR = os.environ.get('DB_DIR', str(Path(__file__).parent.parent))
DB_PATH = Path(DB_DIR) / "tech_cards.db"

# Создаём папку, если её нет
DB_PATH.parent.mkdir(parents=True, exist_ok=True)

DATABASE_URL = f"sqlite:///{DB_PATH}"

# Создаем движок SQLAlchemy
engine = create_engine(
    DATABASE_URL,
    connect_args={"check_same_thread": False, "timeout": 30},
    echo=False
)


# ===== ВКЛЮЧАЕМ WAL-РЕЖИМ И СОПУТСТВУЮЩИЕ PRAGMA =====
# Событие 'connect' срабатывает при каждом новом соединении.
# WAL и synchronous — это настройки, которые НЕ сохраняются между соединениями,
# поэтому их нужно применять каждый раз.
@event.listens_for(engine, "connect")
def _set_sqlite_pragmas(dbapi_connection, connection_record):
    """
    Настройка SQLite-соединения:
      - journal_mode=WAL       — параллельное чтение во время записи
      - synchronous=NORMAL     — безопасно с WAL, быстрее чем FULL
      - busy_timeout=30000     — ждать до 30 секунд при блокировке
    """
    cursor = dbapi_connection.cursor()
    cursor.execute("PRAGMA journal_mode=WAL")
    cursor.execute("PRAGMA synchronous=NORMAL")
    cursor.execute("PRAGMA busy_timeout=30000")
    cursor.close()


# Фабрика сессий
SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine)
db_session = scoped_session(SessionLocal)


def get_db():
    """Генератор сессий для Flask (dependency injection)"""
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()


def init_db():
    """Создание таблиц (импортируется из init_db.py)"""
    from .models import Base
    Base.metadata.create_all(bind=engine)