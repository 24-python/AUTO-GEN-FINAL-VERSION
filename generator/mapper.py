"""
Маппинг между парсером (Enum Category) и базой данных
"""

from parser.models import Category as ParserCategory
from db.models import Category as DBCategory, Object, Instruction
from db.database import SessionLocal


class CategoryMapper:
    """Маппинг категорий из парсера в БД"""

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
        ParserCategory.CONTACT_SURFACES: "Контактные поверхности",
        ParserCategory.OTHER: "Прочее",
    }

    @classmethod
    def get_db_category_name(cls, parser_category: ParserCategory) -> str:
        return cls.ENUM_TO_DB.get(parser_category, "Прочее")


def get_object_by_normalized_name(session, normalized_name: str):
    """Ищет объект в БД по нормализованному имени"""
    return session.query(Object).filter(Object.normalized_name == normalized_name).first()


# Алиас для обратной совместимости
get_object_by_name = get_object_by_normalized_name


def get_instructions_for_object(session, object_id: int):
    """Возвращает все инструкции для объекта"""
    return session.query(Instruction).filter(Instruction.object_id == object_id).all()


def find_object_and_instructions(session, item_name: str, parser_category: ParserCategory):
    """
    Ищет объект в БД по нормализованному имени и возвращает его с инструкциями.
    """
    obj = get_object_by_normalized_name(session, item_name)

    if not obj:
        return None, []

    instructions = get_instructions_for_object(session, obj.id)
    return obj, instructions