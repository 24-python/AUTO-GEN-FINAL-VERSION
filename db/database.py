"""
Подключение к базе данных SQLite
"""

from sqlalchemy import create_engine
from sqlalchemy.orm import sessionmaker, scoped_session
from pathlib import Path

# Путь к файлу БД в корне проекта
DB_PATH = Path(__file__).parent.parent / "tech_cards.db"
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