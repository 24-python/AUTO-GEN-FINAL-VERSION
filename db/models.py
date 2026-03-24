"""
Модели данных для базы данных технологических карт
"""

from sqlalchemy import Column, Integer, String, Text, ForeignKey, DateTime, Index
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from datetime import datetime

Base = declarative_base()


class Category(Base):
    """Категории объектов"""
    __tablename__ = 'categories'

    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False, unique=True)
    sort_order = Column(Integer, default=0)

    # Связи
    objects = relationship("Object", back_populates="category", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<Category(id={self.id}, name='{self.name}')>"


class Object(Base):
    """Объекты обработки из чек-листа"""
    __tablename__ = 'objects'

    id = Column(Integer, primary_key=True)
    category_id = Column(Integer, ForeignKey('categories.id'), nullable=False)

    # Основные поля
    name = Column(String(250), nullable=False, unique=True)  # из текстового файла
    base_name = Column(String(200), nullable=False)
    modifier = Column(String(50), nullable=True)
    sort_priority = Column(Integer, default=0)

    # Служебные поля
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    # Связи
    category = relationship("Category", back_populates="objects")
    instructions = relationship("Instruction", back_populates="object", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<Object(id={self.id}, name='{self.name}')>"


class Instruction(Base):
    """Инструкции по обработке объектов"""
    __tablename__ = 'instructions'

    id = Column(Integer, primary_key=True)
    object_id = Column(Integer, ForeignKey('objects.id'), nullable=False)

    # Поля из Excel
    cleaning_method = Column(String(100), nullable=True)  # мойка/дезинфекция
    instruction_number = Column(String(50), nullable=True)  # № инструкции
    product_name = Column(String(200), nullable=True)  # средство
    cleaning_technique = Column(String(200), nullable=True)  # метод уборки
    concentration = Column(String(200), nullable=True)  # концентрация
    temperature = Column(String(50), nullable=True)  # температура
    exposure_time = Column(String(50), nullable=True)  # время выдержки
    inventory = Column(String(100), nullable=True)  # инвентарь
    frequency = Column(String(100), nullable=True)  # периодичность
    executor = Column(String(200), nullable=True)  # исполнитель
    control_method = Column(String(200), nullable=True)  # метод контроля

    # Служебные поля
    created_at = Column(DateTime, default=datetime.now)

    # Связи
    object = relationship("Object", back_populates="instructions")

    def __repr__(self):
        return f"<Instruction(id={self.id}, object_id={self.object_id}, method='{self.cleaning_method}')>"

    def is_empty(self) -> bool:
        """Проверяет, пустая ли инструкция (нет данных)"""
        fields = [
            self.cleaning_method, self.product_name, self.concentration,
            self.temperature, self.exposure_time, self.frequency,
            self.executor, self.control_method
        ]
        return not any(field for field in fields if field and field.strip())


# Индексы для ускорения поиска
Index('idx_object_name', Object.name)
Index('idx_object_category', Object.category_id)
Index('idx_instruction_object', Instruction.object_id)