"""
Модели данных для базы данных технологических карт
"""

from sqlalchemy import Column, Integer, String, Text, ForeignKey, DateTime, Index, Boolean
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
    object_property = relationship("ObjectProperty", back_populates="object", uselist=False, cascade="all, delete-orphan")
    object_groups = relationship("ObjectGroup", back_populates="object", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<Object(id={self.id}, normalized_name='{self.normalized_name}', display_name='{self.display_name}')>"


class Instruction(Base):
    """Инструкции по обработке объектов"""
    __tablename__ = 'instructions'

    id = Column(Integer, primary_key=True)
    object_id = Column(Integer, ForeignKey('objects.id'), nullable=False)
    room_category_id = Column(Integer, ForeignKey('room_categories.id'), nullable=True)

    maintenance_type = Column(String(50), nullable=True)
    # ===== ДОБАВЛЕНО: предприятие и подуровень (группа) =====
    enterprise = Column(String(200), nullable=True)          # <-- ДОБАВЛЕНО
    subgroup = Column(String(50), nullable=True)             # <-- ДОБАВЛЕНО
    cleaning_method = Column(String(100), nullable=True)
    instruction_number = Column(String(50), nullable=True)
    product_name = Column(String(200), nullable=True)
    cleaning_technique = Column(String(200), nullable=True)
    concentration = Column(String(200), nullable=True)
    application_method = Column(String(200), nullable=True)
    temperature = Column(String(50), nullable=True)
    exposure_time = Column(String(50), nullable=True)
    inventory = Column(String(100), nullable=True)
    frequency = Column(String(100), nullable=True)
    executor = Column(String(200), nullable=True)
    control_method = Column(String(200), nullable=True)
    surface_type = Column(String(50), nullable=True)

    created_at = Column(DateTime, default=datetime.now)

    object = relationship("Object", back_populates="instructions")
    room_category = relationship("RoomCategory", back_populates="instructions")

    def __repr__(self):
        return f"<Instruction(id={self.id}, object_id={self.object_id}, method='{self.cleaning_method}', surface='{self.surface_type}', room_cat='{self.room_category.name if self.room_category else 'общая'}')>"

    def is_empty(self) -> bool:
        fields = [
            self.maintenance_type, self.cleaning_method, self.product_name, self.concentration,
            self.temperature, self.exposure_time, self.frequency,
            self.executor, self.control_method, self.application_method
        ]
        return not any(field for field in fields if field and field.strip())


class Product(Base):
    """Моющие и дезинфицирующие средства"""
    __tablename__ = 'products'

    id = Column(Integer, primary_key=True)
    name = Column(String(200), nullable=False, unique=True)   # название средства
    product_type = Column(String(50), nullable=True)          # тип: дезинфицирующее, нейтральное, щелочное, кислотное, специальное
    color = Column(String(6), nullable=True)                  # HEX-код цвета (без #)

    def __repr__(self):
        return f"<Product(id={self.id}, name='{self.name}', type='{self.product_type}', color='{self.color}')>"


# ========== НОВЫЕ ТАБЛИЦЫ КОНФИГУРАЦИИ ==========

class ObjectProperty(Base):
    """Флаги специальных свойств объектов"""
    __tablename__ = 'object_properties'

    id = Column(Integer, primary_key=True)
    object_id = Column(Integer, ForeignKey('objects.id'), nullable=False, unique=True)
    is_split = Column(Boolean, default=False)                 # SPLIT_SURFACE_OBJECTS
    is_multi_method = Column(Boolean, default=False)          # MULTI_METHOD_OBJECTS
    has_support_maintenance = Column(Boolean, default=False)  # SUPPORT_MAINTENANCE_OBJECTS
    special_product_type = Column(String(20), nullable=True)  # floor, thermal, glass

    object = relationship("Object", back_populates="object_property")

    def __repr__(self):
        return f"<ObjectProperty(object_id={self.object_id}, split={self.is_split}, multi={self.is_multi_method}, support={self.has_support_maintenance}, special={self.special_product_type})>"


class ObjectGroup(Base):
    """Группы объектов для заголовков (GROUP_HEADERS)"""
    __tablename__ = 'object_groups'

    id = Column(Integer, primary_key=True)
    group_name = Column(String(250), nullable=False)          # Название группы (заголовок)
    object_id = Column(Integer, ForeignKey('objects.id'), nullable=False)

    object = relationship("Object", back_populates="object_groups")

    def __repr__(self):
        return f"<ObjectGroup(group='{self.group_name}', object_id={self.object_id})>"


class CleaningMethodOrder(Base):
    """Порядок способов обработки (CLEANING_METHOD_ORDER)"""
    __tablename__ = 'cleaning_method_order'

    id = Column(Integer, primary_key=True)
    method_name = Column(String(150), nullable=False, unique=True)
    sort_order = Column(Integer, nullable=False)

    def __repr__(self):
        return f"<CleaningMethodOrder(method='{self.method_name}', order={self.sort_order})>"


# ===================== ДОБАВЛЕНО: ЦВЕТА ИНВЕНТАРЯ =====================

class InventoryColor(Base):
    """Цвета уборочного инвентаря (редактируемый справочник)"""
    __tablename__ = 'inventory_colors'

    id = Column(Integer, primary_key=True)
    name = Column(String(50), nullable=False, unique=True)      # название цвета (чёрный, жёлтый, ...)
    hex_color = Column(String(7), nullable=False)               # HEX-код (например, #000000)

    def __repr__(self):
        return f"<InventoryColor(id={self.id}, name='{self.name}', hex='{self.hex_color}')>"


# Индексы для ускорения поиска
Index('idx_object_normalized_name', Object.normalized_name)
Index('idx_object_category', Object.category_id)
Index('idx_instruction_object', Instruction.object_id)
Index('idx_instruction_room_category', Instruction.room_category_id)