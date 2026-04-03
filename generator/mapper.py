"""
Маппинг между парсером (Enum Category) и базой данных
"""

from parser.models import Category as ParserCategory
from db.models import Category as DBCategory, Object, Instruction
from db.database import SessionLocal


class CategoryMapper:
    """Маппинг категорий из парсера в БД"""

    # Соответствие Enum Category → название в БД
    ENUM_TO_DB = {
        ParserCategory.SURFACE: "Поверхности",
        ParserCategory.HOUSEHOLD_APPLIANCES: "Бытовая техника",
        ParserCategory.THERMAL_EQUIPMENT: "Тепловое оборудование",
        ParserCategory.PACKAGING_EQUIPMENT: "Упаковочное оборудование",
        ParserCategory.TECH_EQUIPMENT: "Технологическое оборудование",
        ParserCategory.INVENTORY: "Инвентарь, посуда и т.д.",
        ParserCategory.CLEANING_EQUIPMENT: "Моечный, уборочный инвентарь и оборудование",
        ParserCategory.DISHWASHING_EQUIPMENT: "Посудомоечное оборудование",
        ParserCategory.REFRIGERATION_EQUIPMENT: "Холодильное оборудование",
        ParserCategory.DOSING_EQUIPMENT: "Дозирующее оборудование",
        ParserCategory.PLUMBING: "Сантехническое оборудование",
        ParserCategory.FURNITURE: "Мебель",
        ParserCategory.OFFICE_EQUIPMENT: "Офисная техника",
        ParserCategory.SANITARY_POST: "Санитарный пост",
        ParserCategory.PPE: "Многоразовые резиновые СИЗ",
        ParserCategory.OTHER: "Прочее",
    }

    @classmethod
    def get_db_category_name(cls, parser_category: ParserCategory) -> str:
        """Возвращает название категории в БД"""
        return cls.ENUM_TO_DB.get(parser_category, "Прочее")


def get_object_by_name(session, name: str):
    """Ищет объект в БД по точному имени"""
    return session.query(Object).filter(Object.name == name).first()


def get_instructions_for_object(session, object_id: int):
    """Возвращает все инструкции для объекта"""
    return session.query(Instruction).filter(Instruction.object_id == object_id).all()


def find_object_and_instructions(session, item_name: str, parser_category: ParserCategory):
    """
    Ищет объект в БД и возвращает его с инструкциями.
    Сначала по точному имени, потом по категории (если не найден).

    Returns:
        tuple: (object, list_of_instructions) - может быть (None, [])
    """
    # 1. Поиск по точному имени
    obj = get_object_by_name(session, item_name)

    # 2. Если не найден, поиск по категории (для базовых объектов)
    if not obj:
        db_category_name = CategoryMapper.get_db_category_name(parser_category)
        obj = session.query(Object).filter(
            Object.category.has(name=db_category_name),
            Object.name.ilike(f"%{item_name}%")
        ).first()

    if not obj:
        return None, []

    instructions = get_instructions_for_object(session, obj.id)
    return obj, instructions