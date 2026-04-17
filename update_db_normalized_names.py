#!/usr/bin/env python3
"""
update_db_normalized_names.py

Обновляет normalized_name и display_name в БД для проблемных объектов.
Удаляет дубликаты и добавляет правильные записи.
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from db.database import SessionLocal
from db.models import Object, Category

# Список исправлений: {старый_normalized: (новый_normalized, новый_display, категория)}
FIXES = {
    # Уже есть в БД, но с другим normalized_name
    "автоматический клипсаторы": ("автоматическиеклипсаторы", "автоматические клипсаторы", "Упаковочное оборудование"),
    "демосистема (контактная поверхность)": ("демосистема(контактная поверхность)",
                                           "демосистема (контактная поверхность)", "Поверхности"),
    "диспенсер для полотенец": ("диспенсердержатель для полотенец", "диспенсер/держатель для полотенец",
                                        "Санитарный пост"),
    "корзины для расстойки теста": ("корзины длярасстойкитеста", "корзины для расстойки теста",
                                 "Инвентарь, посуда и т.д."),
    "машина для резки конд изделий": ("машина для резкиконд изделий", "машина для резки кондитерских изделий",
                                   "Технологическое оборудование"),
    "минифилы (дозаторы крема)": ("минифилы(дозаторы крема)", "минифилы (дозаторы крема)",
                                "Технологическое оборудование"),
    "расстоечные термочехлы": ("расстоечныетермочехлы", "расстоечные термочехлы", "Инвентарь, посуда и т.д."),
    "расстоечные шкафы": ("расстоечныешкафы", "расстоечные шкафы", "Тепловое оборудование"),
    "силапеновые коврики": ("силапеновыековрики", "силапеновые коврики", "Инвентарь, посуда и т.д."),
    "таромоечная машина": ("таромоечнаямашина", "таромоечная машина", "Посудомоечное оборудование"),
    "тележки подкатные": ("тележкиподкатные", "тележки подкатные", "Инвентарь, посуда и т.д."),
    "ёмкости для перетаривания": ("ёмкости дляперетаривания", "ёмкости для перетаривания", "Инвентарь, посуда и т.д."),
}

# Новые объекты, которых нет в БД
NEW_OBJECTS = {
    "дозатор для геля (пульверизатор)": ("дозатор для геля (пульверизатор)", "Технологическое оборудование"),
    "металлодетектор": ("металлодетектор", "Технологическое оборудование"),
    "отсадочные машины": ("отсадочные машины", "Технологическое оборудование"),
    "шприц-дозатор начинки": ("шприц-дозатор начинки", "Технологическое оборудование"),
    "ёлки для колец": ("ёлки для колец", "Поверхности"),
}


def main():
    session = SessionLocal()

    # Кэш категорий
    categories = {cat.name: cat.id for cat in session.query(Category).all()}

    print("=" * 80)
    print("🔧 ОБНОВЛЕНИЕ NORMALIZED_NAME В БД")
    print("=" * 80)

    updated_count = 0
    deleted_count = 0
    added_count = 0

    # 1. ИСПРАВЛЯЕМ СУЩЕСТВУЮЩИЕ ОБЪЕКТЫ
    print("\n📝 ИСПРАВЛЕНИЕ СУЩЕСТВУЮЩИХ ОБЪЕКТОВ:")
    print("-" * 60)

    for old_norm, (new_norm, new_display, cat_name) in FIXES.items():
        # Ищем объект по старому normalized_name
        obj = session.query(Object).filter(Object.normalized_name == old_norm).first()

        if obj:
            print(f"\n   Старое: {old_norm}")
            print(f"   Новое:  {new_norm}")
            print(f"   Display: {new_display}")
            print(f"   Категория: {cat_name}")

            # Проверяем, нет ли уже объекта с новым именем
            existing = session.query(Object).filter(Object.normalized_name == new_norm).first()
            if existing and existing.id != obj.id:
                # Удаляем дубликат
                print(f"   ⚠️ Найден дубликат, удаляем старый (id={obj.id})")
                session.delete(obj)
                deleted_count += 1
            else:
                # Обновляем
                obj.normalized_name = new_norm
                obj.display_name = new_display
                if cat_name in categories:
                    obj.category_id = categories[cat_name]
                updated_count += 1
                print(f"   ✅ Обновлено")
        else:
            print(f"\n   ❌ Не найден: {old_norm}")

    # 2. ДОБАВЛЯЕМ НОВЫЕ ОБЪЕКТЫ
    print("\n\n➕ ДОБАВЛЕНИЕ НОВЫХ ОБЪЕКТОВ:")
    print("-" * 60)

    for display_name, (norm_name, cat_name) in NEW_OBJECTS.items():
        # Проверяем, нет ли уже
        existing = session.query(Object).filter(
            (Object.normalized_name == norm_name) |
            (Object.display_name == display_name)
        ).first()

        if existing:
            print(f"\n   ⏭️ Уже существует: {norm_name} → {existing.display_name}")
            continue

        if cat_name not in categories:
            print(f"\n   ❌ Категория не найдена: {cat_name}")
            continue

        # Определяем base_name и modifier
        parts = norm_name.split()
        base_name = parts[0] if parts else norm_name
        modifier = ' '.join(parts[1:]) if len(parts) > 1 else None

        obj = Object(
            normalized_name=norm_name,
            display_name=display_name,
            base_name=base_name,
            modifier=modifier,
            category_id=categories[cat_name],
            sort_priority=0
        )
        session.add(obj)
        added_count += 1
        print(f"\n   ✅ Добавлен: {norm_name}")
        print(f"      Display: {display_name}")
        print(f"      Категория: {cat_name}")

    # Сохраняем изменения
    session.commit()

    # 3. ИТОГИ
    print("\n" + "=" * 80)
    print("📊 ИТОГИ ОБНОВЛЕНИЯ")
    print("=" * 80)
    print(f"   Обновлено объектов: {updated_count}")
    print(f"   Удалено дубликатов: {deleted_count}")
    print(f"   Добавлено новых: {added_count}")
    print(f"   ВСЕГО изменений: {updated_count + deleted_count + added_count}")

    # 4. ПРОВЕРКА ОСТАВШИХСЯ ПРОБЛЕМНЫХ ИМЁН
    print("\n" + "=" * 80)
    print("🔍 ПРОВЕРКА ОСТАВШИХСЯ ПРОБЛЕМНЫХ ИМЁН")
    print("=" * 80)

    problem_patterns = [
        ('%автоматически%', 'автоматическиеклипсаторы'),
        ('%диспенсер%полотенец%', 'диспенсердержатель для полотенец'),
        ('%дозатор%геля%', 'дозатор для геля (пульверизатор)'),
        ('%ёлки%', 'ёлки для колец'),
        ('%шприц%', 'шприц-дозатор начинки'),
    ]

    for pattern, expected in problem_patterns:
        objects = session.query(Object).filter(Object.normalized_name.like(pattern)).all()
        if objects:
            print(f"\n   Найдено по '{pattern}':")
            for obj in objects:
                status = "✅" if obj.normalized_name == expected else "⚠️"
                print(f"      {status} {obj.normalized_name} → {obj.display_name} [{obj.category.name}]")

    session.close()

    print("\n" + "=" * 80)
    print("✅ ОБНОВЛЕНИЕ ЗАВЕРШЕНО")
    print("=" * 80)


if __name__ == "__main__":
    main()