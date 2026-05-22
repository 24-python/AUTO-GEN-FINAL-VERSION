#!/usr/bin/env python3
"""
Миграция конфигурационных словарей из кода в БД.
Заполняет таблицы object_properties и object_groups данными из старых словарей.
Запускается один раз после обновления моделей.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from db.database import SessionLocal
from db.models import Object, ObjectProperty, ObjectGroup

# Переносимые словари (актуальные на момент миграции)
SUPPORT_MAINTENANCE_OBJECTS = {
    "прибор кисл теста", "бисквиторезки", "блендеры",
    "вакуумные роторные шприцы", "водяные бани", "депозитор волюметрический",
    "дозатор для геля (пульверизатор)", "дозаторы для жидкостей", "дробилки",
    "измельчители", "картофелечистки", "машина для резки конд изделий",
    "металлодетектор", "миксеры планетарные", "минифилы (дозаторы крема)",
    "овощерезки", "овощечистки", "отсадочные машины", "пневматические распылители",
    "прессы для теста", "просеиватели мука", "просеиватели сахар",
    "протирочные машины", "распылители для желе и сиропов",
    "рентгеновские системы контроля", "слайсера", "солодоварки",
    "тарталетницы", "термощупы", "тестоделители", "тестомесы",
    "тестоокруглители", "тестораскатки", "ультразвуковые нарезки",
    "ферментаторы", "шприц-дозатор начинки", "весы напольные",
    "весы настольные", "производственные столы д", "производственные столы н",
}

SPLIT_SURFACE_OBJECTS = {
    "холод шкафы", "камеры мороз", "камеры шок замор",
    "морозильный ларь", "холод столы", "пароконвектоматы",
    "ледогенератор", "ферментаторы", "солодоварки",
    "пмм купольная", "пмм туннельная", "таромоечная машина",
    "печи ротационные", "печи подовые", "багетницы",
    "аппарат для приготовления фрикаделек"
}

MULTI_METHOD_OBJECTS = {
    "вешала", "внутрицеховая тара (вёдра ящики)", "гастроёмкости", "дежи", "доски",
    "изотермические контейнеры (bigbox)", "инвентарь", "корзины для расстойки теста",
    "крючки", "листы для выпечки а", "листы для выпечки н", "листы от шпилек",
    "мусаты", "ножи", "оборотная тара", "отсадочные мешки", "передвижные ёмкости",
    "посуда", "расстоечные термочехлы", "секачи", "силапеновые коврики",
    "съёмные детали оборудования а", "съёмные детали оборудования н", "тележки",
    "тележки подкатные", "формы для выпечки а", "формы для выпечки н",
    "формы для выпечки с", "шампура", "шпильки", "ёмкости для перетаривания",
    "ёмкости для сыпучих продуктов"
}

FLOOR_OBJECTS = {"пол", "трапы"}
GLASS_OBJECTS = {"зеркала", "окна внешние", "окна внутрицеховые", "монитор"}
THERMAL_OBJECTS = {
    "плиты индук", "плиты элек", "плиты газ",
    "варочные котлы", "сковороды", "фритюры", "грили",
    "пароконвектоматы", "печи подовые", "печи ротационные",
    "расстоечные шкафы", "вафельницы"
}

GROUP_HEADERS = {
    frozenset({"весы напольные", "весы настольные"}): "Санитарная обработка средств измерений",
    frozenset({
        "прибор кисл теста", "бисквиторезки", "блендеры", "вакуумные роторные шприцы",
        "водяные бани", "депозитор волюметрический", "дозатор для геля (пульверизатор)",
        "дозаторы для жидкостей", "дробилки", "измельчители", "картофелечистки",
        "машина для резки конд изделий", "металлодетектор", "миксеры планетарные",
        "минифилы (дозаторы крема)", "овощерезки", "овощечистки", "отсадочные машины",
        "пневматические распылители", "прессы для теста", "просеиватели мука",
        "просеиватели сахар", "протирочные машины", "распылители для желе и сиропов",
        "рентгеновские системы контроля", "слайсера", "солодоварки", "тарталетницы",
        "термощупы", "тестоделители", "тестомесы", "тестоокруглители", "тестораскатки",
        "ультразвуковые нарезки", "ферментаторы", "шприц-дозатор начинки",
        "пмм купольная", "пмм туннельная", "таромоечная машина",
        "стиральные машины", "сушильные машины", "холодильная витрина", "ледогенератор",
        "морозильный ларь", "холод шкафы", "холод столы"
    }): "Санитарная обработка технологического оборудования",
    frozenset({
        "производственные столы д", "производственные столы н"}): "Санитарная обработка рабочих поверхностей",
    frozenset({
        "авд", "ведра", "ветошь", "мопы", "поломоечная машина", "пылесосы", "сгоны",
        "тележки уборочные", "щётки"
    }): "Санитарная обработка уборочного инвентаря и т.д.",
    frozenset({
        "вешала", "внутрицеховая тара (вёдра ящики)", "гастроёмкости", "дежи", "доски",
        "изотермические контейнеры (bigbox)", "инвентарь", "корзины для расстойки теста",
        "крючки", "листы для выпечки а", "листы для выпечки н", "листы от шпилек",
        "мусаты", "ножи", "оборотная тара", "отсадочные мешки", "передвижные ёмкости",
        "посуда", "расстоечные термочехлы", "секачи", "силапеновые коврики",
        "съёмные детали оборудования а", "съёмные детали оборудования н", "тележки",
        "тележки подкатные", "формы для выпечки а", "формы для выпечки н",
        "формы для выпечки с", "шампура", "шпильки", "ёмкости для перетаривания",
        "ёмкости для сыпучих продуктов"
    }): "Санитарная обработка производственного инвентаря, посуды",
}


def migrate():
    session = SessionLocal()
    try:
        all_objects = session.query(Object).all()
        obj_map = {obj.normalized_name: obj for obj in all_objects}

        # ---------- 1. Сбор объединённых свойств объектов ----------
        print("🔹 Сбор свойств объектов...")
        # Структура: { object_id: { 'is_split': bool, 'is_multi': bool, 'support': bool, 'special': str } }
        obj_props = {}

        for name in SUPPORT_MAINTENANCE_OBJECTS:
            obj = obj_map.get(name)
            if obj:
                props = obj_props.setdefault(obj.id, {
                    'is_split': False,
                    'is_multi_method': False,
                    'has_support_maintenance': False,
                    'special_product_type': None
                })
                props['has_support_maintenance'] = True

        for name in SPLIT_SURFACE_OBJECTS:
            obj = obj_map.get(name)
            if obj:
                props = obj_props.setdefault(obj.id, {
                    'is_split': False,
                    'is_multi_method': False,
                    'has_support_maintenance': False,
                    'special_product_type': None
                })
                props['is_split'] = True

        for name in MULTI_METHOD_OBJECTS:
            obj = obj_map.get(name)
            if obj:
                props = obj_props.setdefault(obj.id, {
                    'is_split': False,
                    'is_multi_method': False,
                    'has_support_maintenance': False,
                    'special_product_type': None
                })
                props['is_multi_method'] = True

        for name in FLOOR_OBJECTS:
            obj = obj_map.get(name)
            if obj:
                props = obj_props.setdefault(obj.id, {
                    'is_split': False,
                    'is_multi_method': False,
                    'has_support_maintenance': False,
                    'special_product_type': None
                })
                props['special_product_type'] = 'floor'

        for name in GLASS_OBJECTS:
            obj = obj_map.get(name)
            if obj:
                props = obj_props.setdefault(obj.id, {
                    'is_split': False,
                    'is_multi_method': False,
                    'has_support_maintenance': False,
                    'special_product_type': None
                })
                props['special_product_type'] = 'glass'

        for name in THERMAL_OBJECTS:
            obj = obj_map.get(name)
            if obj:
                props = obj_props.setdefault(obj.id, {
                    'is_split': False,
                    'is_multi_method': False,
                    'has_support_maintenance': False,
                    'special_product_type': None
                })
                props['special_product_type'] = 'thermal'

        print(f"   Уникальных объектов со свойствами: {len(obj_props)}")

        # ---------- 2. Применение свойств в БД (создание/обновление) ----------
        print("🔹 Сохранение свойств...")
        created = 0
        updated = 0
        for obj_id, props in obj_props.items():
            obj_prop = session.query(ObjectProperty).filter_by(object_id=obj_id).first()
            if obj_prop:
                obj_prop.is_split = props['is_split']
                obj_prop.is_multi_method = props['is_multi_method']
                obj_prop.has_support_maintenance = props['has_support_maintenance']
                obj_prop.special_product_type = props['special_product_type']
                updated += 1
            else:
                session.add(ObjectProperty(
                    object_id=obj_id,
                    is_split=props['is_split'],
                    is_multi_method=props['is_multi_method'],
                    has_support_maintenance=props['has_support_maintenance'],
                    special_product_type=props['special_product_type']
                ))
                created += 1
        print(f"   Создано: {created}, обновлено: {updated}")

        # ---------- 3. Группы объектов ----------
        print("🔹 Перенос групп объектов...")
        group_count = 0
        for group_set, header in GROUP_HEADERS.items():
            for name in group_set:
                obj = obj_map.get(name)
                if obj:
                    existing = session.query(ObjectGroup).filter_by(
                        group_name=header, object_id=obj.id).first()
                    if not existing:
                        session.add(ObjectGroup(group_name=header, object_id=obj.id))
                        group_count += 1
        print(f"   Создано {group_count} записей групп")

        session.commit()
        print("✅ Миграция успешно завершена!")

    except Exception as e:
        session.rollback()
        print(f"❌ Ошибка миграции: {e}")
        raise
    finally:
        session.close()


if __name__ == "__main__":
    print("🚀 Запуск миграции конфигурации в БД")
    migrate()