"""
Подключение к базе данных SQLite
"""

import os
from sqlalchemy import create_engine
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
    connect_args={"check_same_thread": False},
    echo=False  # True для отладки SQL
)

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