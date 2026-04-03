#!/usr/bin/env python3
"""
prepare_import_json.py

Скрипт для подготовки JSON-файла импорта в базу данных.
Читает:
- текстовый файл с объектами (полный перечень всех поверхностей из чек-листа.txt)
- Excel-файл с инструкциями (Объекты для базы данных.xlsx)

Создает:
- import_data.json (для импорта в БД)
"""

import json
import re
from pathlib import Path
from typing import Dict, List, Optional
from dataclasses import dataclass, field, asdict
from collections import defaultdict

# Для работы с Excel
try:
    import openpyxl
except ImportError:
    print("Установите openpyxl: pip install openpyxl")
    exit(1)


# ============================================================
# МОДЕЛИ ДАННЫХ
# ============================================================

@dataclass
class Instruction:
    """Инструкция по обработке объекта"""
    cleaning_method: str = ""  # мойка/дезинфекция/очистка
    instruction_number: str = ""  # № инструкции
    product_name: str = ""  # средство
    cleaning_technique: str = ""  # метод уборки
    concentration: str = ""  # концентрация
    temperature: str = ""  # температура
    exposure_time: str = ""  # время выдержки
    inventory: str = ""  # инвентарь
    frequency: str = ""  # периодичность
    executor: str = ""  # исполнитель
    control_method: str = ""  # метод контроля

    def to_dict(self) -> dict:
        """Преобразует в словарь, исключая пустые поля (опционально)"""
        result = {}
        for key, value in asdict(self).items():
            if value and value.strip():
                result[key] = value
        return result


@dataclass
class ImportObject:
    """Объект для импорта в БД"""
    name: str  # полное название из текстового файла
    base_name: str  # базовая часть (без модификатора)
    modifier: Optional[str]  # модификатор (П, ОК, Н, А, С, и т.д.)
    category: str  # категория
    sort_priority: int = 0  # приоритет сортировки внутри категории
    instructions: List[Instruction] = field(default_factory=list)

    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "base_name": self.base_name,
            "modifier": self.modifier,
            "category": self.category,
            "sort_priority": self.sort_priority,
            "instructions": [instr.to_dict() for instr in self.instructions]
        }


# ============================================================
# ПАРСЕР ТЕКСТОВОГО ФАЙЛА
# ============================================================

def parse_objects_from_txt(txt_path: str) -> List[ImportObject]:
    """
    Парсит текстовый файл с объектами.
    Формат:
        Категория:
        объект1
        объект2
        ...
    """
    objects = []
    current_category = None

    with open(txt_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    for line in lines:
        line = line.strip()
        if not line:
            continue

        # Проверка на заголовок категории (заканчивается на ":")
        if line.endswith(':'):
            current_category = line.rstrip(':')
            continue

        # Если нет активной категории, пропускаем
        if current_category is None:
            continue

        # Парсим объект
        obj = parse_object_line(line, current_category)
        if obj:
            objects.append(obj)

    return objects


def parse_object_line(line: str, category: str) -> Optional[ImportObject]:
    """
    Парсит строку объекта.
    Определяет base_name и modifier.

    Примеры:
        "потолок (высота _____м)" → base_name="потолок", modifier=None
        "потолок (высота _____м) П" → base_name="потолок", modifier="П"
        "плиты индук." → base_name="плиты", modifier="индук."
        "формы для выпечки С" → base_name="формы для выпечки", modifier="С"
    """
    name = line.strip()

    # Определяем base_name и modifier
    base_name = name
    modifier = None

    # Паттерны для извлечения модификатора
    patterns = [
        # "потолок (высота _____м) П" → modifier = "П"
        (r'^(.*?)\s+([А-Яа-я0-9]+(?:\.[а-я]+)?)$', 1, 2),
        # "плиты индук." → modifier = "индук."
        (r'^(.*?)\s+([а-я]+\.?)$', 1, 2),
        # "формы для выпечки С" → modifier = "С"
        (r'^(.*?)\s+([А-Я])$', 1, 2),
    ]

    for pattern, base_idx, mod_idx in patterns:
        match = re.match(pattern, name)
        if match:
            base_name = match.group(base_idx).strip()
            modifier = match.group(mod_idx).strip()
            break

    # Очистка base_name от лишних символов
    base_name = re.sub(r'\([^)]*\)', '', base_name).strip()
    base_name = re.sub(r'\s+', ' ', base_name)

    return ImportObject(
        name=name,
        base_name=base_name,
        modifier=modifier,
        category=category,
        sort_priority=0
    )


# ============================================================
# ПАРСЕР EXCEL-ФАЙЛА
# ============================================================

def parse_instructions_from_excel(xlsx_path: str) -> Dict[str, List[Instruction]]:
    """
    Парсит Excel-файл с инструкциями.
    Возвращает словарь: {название_объекта: [список_инструкций]}
    """
    instructions_map = defaultdict(list)

    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb.active

    # Получаем заголовки (первая строка)
    headers = []
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=col).value
        headers.append(value if value else "")

    # Определяем индексы колонок
    col_indices = {
        'object': 0,  # Объект обработки
        'method': 1,  # Способ обработки
        'instruction_no': 2,  # № инструкции
        'product': 3,  # Наименование средства
        'technique': 4,  # Метод уборки
        'concentration': 5,  # Концентрация
        'temperature': 6,  # Температура
        'exposure': 7,  # Время выдержки
        'inventory': 8,  # Инвентарь
        'frequency': 9,  # Периодичность
        'executor': 10,  # Исполнитель
        'control': 11,  # Метод контроля
    }

    # Проходим по строкам (начиная со 2-й)
    for row in range(2, ws.max_row + 1):
        object_name = ws.cell(row=row, column=col_indices['object'] + 1).value
        if not object_name or not str(object_name).strip():
            continue

        object_name = str(object_name).strip()

        # Пропускаем строки-заголовки категорий (если они попали)
        if object_name.endswith(':'):
            continue

        # Создаем инструкцию
        instruction = Instruction(
            cleaning_method=str(ws.cell(row=row, column=col_indices['method'] + 1).value or ""),
            instruction_number=str(ws.cell(row=row, column=col_indices['instruction_no'] + 1).value or ""),
            product_name=str(ws.cell(row=row, column=col_indices['product'] + 1).value or ""),
            cleaning_technique=str(ws.cell(row=row, column=col_indices['technique'] + 1).value or ""),
            concentration=str(ws.cell(row=row, column=col_indices['concentration'] + 1).value or ""),
            temperature=str(ws.cell(row=row, column=col_indices['temperature'] + 1).value or ""),
            exposure_time=str(ws.cell(row=row, column=col_indices['exposure'] + 1).value or ""),
            inventory=str(ws.cell(row=row, column=col_indices['inventory'] + 1).value or ""),
            frequency=str(ws.cell(row=row, column=col_indices['frequency'] + 1).value or ""),
            executor=str(ws.cell(row=row, column=col_indices['executor'] + 1).value or ""),
            control_method=str(ws.cell(row=row, column=col_indices['control'] + 1).value or "")
        )

        instructions_map[object_name].append(instruction)

    return instructions_map


# ============================================================
# МАППИНГ НАЗВАНИЙ (если Excel названы иначе)
# ============================================================

def create_name_mapping() -> Dict[str, str]:
    """
    Создает маппинг названий из Excel в названия из текстового файла.
    Если названия совпадают, маппинг не нужен.
    """
    mapping = {
        # Примеры (если нужно):
        # "потолок подвесной": "потолок (высота _____м) П",
        # "потолок окрашенный": "потолок (высота _____м) ОК",
    }
    return mapping


# ============================================================
# ГЕНЕРАЦИЯ JSON
# ============================================================

def generate_import_json(
        txt_path: str,
        xlsx_path: str,
        output_path: str,
        category_sort_order: Dict[str, int] = None
):
    """
    Генерирует JSON-файл для импорта в БД.
    """
    print("🔧 ПОДГОТОВКА JSON ДЛЯ ИМПОРТА")
    print("=" * 50)

    # 1. Парсим объекты из текстового файла
    print(f"\n📄 Чтение объектов из: {txt_path}")
    objects = parse_objects_from_txt(txt_path)
    print(f"   Найдено объектов: {len(objects)}")

    # 2. Парсим инструкции из Excel
    print(f"\n📊 Чтение инструкций из: {xlsx_path}")
    instructions_map = parse_instructions_from_excel(xlsx_path)
    print(f"   Найдено объектов с инструкциями: {len(instructions_map)}")
    total_instructions = sum(len(instrs) for instrs in instructions_map.values())
    print(f"   Всего инструкций: {total_instructions}")

    # 3. Маппинг названий (если нужно)
    name_mapping = create_name_mapping()

    # 4. Связываем инструкции с объектами
    print("\n🔗 Связывание инструкций с объектами...")
    matched = 0
    not_matched = []

    for obj in objects:
        # Проверяем прямое совпадение
        if obj.name in instructions_map:
            obj.instructions = instructions_map[obj.name]
            matched += 1
        # Проверяем маппинг
        elif obj.name in name_mapping and name_mapping[obj.name] in instructions_map:
            obj.instructions = instructions_map[name_mapping[obj.name]]
            matched += 1
        else:
            not_matched.append(obj.name)

    print(f"   Совпало: {matched}")
    print(f"   Не совпало (будут без инструкций): {len(not_matched)}")

    # 5. Категории с порядком сортировки
    if category_sort_order is None:
        # Порядок категорий из текстового файла
        categories_order = [
            "Поверхности",
            "Сантехническое оборудование",
            "Санитарный пост",
            "Мебель",
            "Офисная техника",
            "Многоразовые резиновые СИЗ",
            "Бытовая техника",
            "Инвентарь, посуда и т.д.",
            "Моечный, уборочный инвентарь и оборудование",
            "Посудомоечное оборудование",
            "Холодильное оборудование",
            "Дозирующее оборудование",
            "Тепловое оборудование",
            "Технологическое оборудование",
            "Упаковочное оборудование",
        ]
        category_sort_order = {name: idx for idx, name in enumerate(categories_order, 1)}

    categories = [
        {"name": name, "sort_order": order}
        for name, order in category_sort_order.items()
    ]

    # 6. Формируем итоговый JSON
    output_data = {
        "version": "1.0",
        "created_at": "2026-04-02",
        "categories": categories,
        "objects": [obj.to_dict() for obj in objects]
    }

    # 7. Сохраняем JSON
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(output_data, f, ensure_ascii=False, indent=2)

    print(f"\n✅ JSON файл создан: {output_path}")
    print(f"   Размер: {Path(output_path).stat().st_size / 1024:.1f} KB")

    # 8. Отчет по несовпавшим объектам
    if not_matched:
        print(f"\n⚠️ ВНИМАНИЕ: Для {len(not_matched)} объектов нет инструкций:")
        for name in not_matched[:20]:  # показываем первые 20
            print(f"   - {name}")
        if len(not_matched) > 20:
            print(f"   ... и еще {len(not_matched) - 20} объектов")

    return output_data


# ============================================================
# ТОЧКА ВХОДА
# ============================================================

def main():
    import argparse

    parser = argparse.ArgumentParser(description="Подготовка JSON для импорта в БД")
    parser.add_argument("--txt", default="полный перечень всех поверхностей из чек-листа.txt",
                        help="Путь к текстовому файлу с объектами")
    parser.add_argument("--xlsx", default="Объекты для базы данных.xlsx",
                        help="Путь к Excel-файлу с инструкциями")
    parser.add_argument("--output", "-o", default="import_data.json",
                        help="Выходной JSON файл")

    args = parser.parse_args()

    # Проверяем существование файлов
    if not Path(args.txt).exists():
        print(f"❌ Файл не найден: {args.txt}")
        return

    if not Path(args.xlsx).exists():
        print(f"❌ Файл не найден: {args.xlsx}")
        return

    generate_import_json(args.txt, args.xlsx, args.output)


if __name__ == "__main__":
    main()