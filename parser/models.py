from dataclasses import dataclass, field
from typing import List, Dict, Optional
from enum import Enum


class Category(Enum):
    """Категории из исходного чек-листа"""
    SURFACE = "Поверхности"
    HOUSEHOLD_APPLIANCES = "Бытовая техника"
    THERMAL_EQUIPMENT = "Тепловое оборудование"
    PACKAGING_EQUIPMENT = "Упаковочное оборудование"
    TECH_EQUIPMENT = "Технологическое оборудование"
    INVENTORY = "Инвентарь, посуда и т.д."
    CLEANING_EQUIPMENT = "Моечный, уборочный инвентарь и оборудование"
    DISHWASHING_EQUIPMENT = "Посудомоечное оборудование"
    REFRIGERATION_EQUIPMENT = "Холодильное оборудование"
    DOSING_EQUIPMENT = "Дозирующее оборудование"
    PLUMBING = "Сантехническое оборудование"
    FURNITURE = "Мебель"
    OFFICE_EQUIPMENT = "Офисная техника"
    SANITARY_POST = "Санитарный пост"
    PPE = "Многоразовые резиновые СИЗ"
    CONTACT_SURFACES = "Контактные поверхности"   # <-- новая категория
    OTHER = "Прочее"

    @classmethod
    def _missing_(cls, value):
        return cls.OTHER


@dataclass
class ChecklistItem:
    """Элемент чек-листа"""
    name: str
    category: Category = Category.OTHER
    subcategory: Optional[str] = None
    checked: bool = False
    markers: List[str] = field(default_factory=list)


@dataclass
class ChecklistData:
    """Данные чек-листа"""
    file_path: str
    room_name: Optional[str] = None
    enterprise: Optional[str] = None
    room_category: Optional[str] = None

    # Общие моющие средства (из раздела "Дополнительная информация")
    cleaning_product: Optional[str] = None
    cleaning_concentration: Optional[str] = None
    cleaning_method_text: Optional[str] = None

    # Общее дезинфицирующее средство
    disinfection_product: Optional[str] = None
    disinfection_concentration: Optional[str] = None
    disinfection_method_text: Optional[str] = None

    # Специализированные моющие средства
    # Для пола и трапов
    floor_cleaning_product: Optional[str] = None
    floor_cleaning_concentration: Optional[str] = None
    floor_cleaning_method_text: Optional[str] = None

    # Для теплового оборудования
    thermal_cleaning_product: Optional[str] = None
    thermal_cleaning_concentration: Optional[str] = None
    thermal_cleaning_method_text: Optional[str] = None

    # Для стеклянных/зеркальных поверхностей (зеркала, окна, мониторы)
    glass_cleaning_product: Optional[str] = None
    glass_cleaning_concentration: Optional[str] = None
    glass_cleaning_method_text: Optional[str] = None

    # Цветовое кодирование инвентаря (из выпадающего списка)
    inventory_color: Optional[str] = None

    items: List[ChecklistItem] = field(default_factory=list)

    def get_checked_items(self) -> List[ChecklistItem]:
        """Возвращает только отмеченные элементы"""
        return [item for item in self.items if item.checked]

    def group_checked_by_category(self) -> Dict[Category, List[ChecklistItem]]:
        """Группирует отмеченные элементы по категориям"""
        result = {}
        for item in self.get_checked_items():
            if item.category not in result:
                result[item.category] = []
            result[item.category].append(item)
        return result