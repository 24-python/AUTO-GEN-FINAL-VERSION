"""
Модели данных для базы данных технологических карт
"""

from sqlalchemy import Column, Integer, String, Text, ForeignKey, DateTime, Index
from sqlalchemy.ext.declarative import declarative_base
from sqlalchemy.orm import relationship
from datetime import datetime

Base = declarative_base()


class Category(Base):
    """Категории объектов (15 штук, порядок сортировки задает умный маппинг)"""
    __tablename__ = 'categories'

    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False, unique=True)
    sort_order = Column(Integer, default=0)

    objects = relationship("Object", back_populates="category", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<Category(id={self.id}, name='{self.name}')>"


class RoomCategory(Base):
    """Категории помещений (производственное, складское, инженерное, вспомогательное)"""
    __tablename__ = 'room_categories'

    id = Column(Integer, primary_key=True)
    name = Column(String(100), nullable=False, unique=True)

    instructions = relationship("Instruction", back_populates="room_category")

    def __repr__(self):
        return f"<RoomCategory(id={self.id}, name='{self.name}')>"


class Object(Base):
    """Объекты обработки из чек-листа"""
    __tablename__ = 'objects'

    id = Column(Integer, primary_key=True)
    category_id = Column(Integer, ForeignKey('categories.id'), nullable=False)

    # Основные поля
    normalized_name = Column(String(250), nullable=False, unique=True)  # для поиска (из чек-листа)
    display_name = Column(String(250), nullable=False)                  # для вывода в техкарту
    base_name = Column(String(200), nullable=False)                    # базовая часть
    modifier = Column(String(50), nullable=True)                       # модификатор
    sort_priority = Column(Integer, default=0)                         # приоритет сортировки

    # Служебные поля
    created_at = Column(DateTime, default=datetime.now)
    updated_at = Column(DateTime, default=datetime.now, onupdate=datetime.now)

    # Связи
    category = relationship("Category", back_populates="objects")
    instructions = relationship("Instruction", back_populates="object", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<Object(id={self.id}, normalized_name='{self.normalized_name}', display_name='{self.display_name}')>"


class Instruction(Base):
    """Инструкции по обработке объектов"""
    __tablename__ = 'instructions'

    id = Column(Integer, primary_key=True)
    object_id = Column(Integer, ForeignKey('objects.id'), nullable=False)
    room_category_id = Column(Integer, ForeignKey('room_categories.id'), nullable=True)  # NULL = общая инструкция

    # Поля из Excel
    maintenance_type = Column(String(50), nullable=True)          # основная / поддерживающая / генеральная
    cleaning_method = Column(String(100), nullable=True)          # мойка/дезинфекция/очистка
    instruction_number = Column(String(50), nullable=True)        # № инструкции
    product_name = Column(String(200), nullable=True)             # наименование средства
    cleaning_technique = Column(String(200), nullable=True)       # метод уборки
    concentration = Column(String(200), nullable=True)            # концентрация
    temperature = Column(String(50), nullable=True)               # температура раствора
    exposure_time = Column(String(50), nullable=True)             # время выдержки
    inventory = Column(String(100), nullable=True)                # инвентарь
    frequency = Column(String(100), nullable=True)                # периодичность
    executor = Column(String(200), nullable=True)                 # исполнитель
    control_method = Column(String(200), nullable=True)           # метод контроля

    # Новое поле для типа поверхности (внешняя/внутренняя/очистка от мин. отложений)
    surface_type = Column(String(50), nullable=True)

    created_at = Column(DateTime, default=datetime.now)

    # Связи
    object = relationship("Object", back_populates="instructions")
    room_category = relationship("RoomCategory", back_populates="instructions")

    def __repr__(self):
        return f"<Instruction(id={self.id}, object_id={self.object_id}, method='{self.cleaning_method}', surface='{self.surface_type}', room_cat='{self.room_category.name if self.room_category else 'общая'}')>"

    def is_empty(self) -> bool:
        """Проверяет, пустая ли инструкция (нет данных)"""
        fields = [
            self.maintenance_type, self.cleaning_method, self.product_name, self.concentration,
            self.temperature, self.exposure_time, self.frequency,
            self.executor, self.control_method
        ]
        return not any(field for field in fields if field and field.strip())


# Индексы для ускорения поиска
Index('idx_object_normalized_name', Object.normalized_name)
Index('idx_object_category', Object.category_id)
Index('idx_instruction_object', Instruction.object_id)
Index('idx_instruction_room_category', Instruction.room_category_id)