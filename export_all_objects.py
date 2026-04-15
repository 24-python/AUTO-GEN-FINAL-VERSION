#!/usr/bin/env python3
"""
export_all_objects.py

Выводит ПОЛНЫЙ список всех объектов из чек-листа:
- Базовые объекты (с SDT)
- Объекты из Unicode-ячеек
- Все модификаторы

Формат вывода:
[статус] normalized_name | display_name | категория | модификаторы
"""

import zipfile
import re
import sys
import json
from pathlib import Path
from collections import defaultdict
from lxml import etree

sys.path.insert(0, str(Path(__file__).parent))

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}

# ============================================================
# КАТЕГОРИИ (15 шт)
# ============================================================
ALL_CATEGORIES = [
    'Поверхности',
    'Сантехническое оборудование',
    'Санитарный пост',
    'Мебель',
    'Офисная техника',
    'Многоразовые резиновые СИЗ',
    'Бытовая техника',
    'Инвентарь, посуда и т.д.',
    'Моечный, уборочный инвентарь и оборудование',
    'Посудомоечное оборудование',
    'Холодильное оборудование',
    'Дозирующее оборудование',
    'Тепловое оборудование',
    'Технологическое оборудование',
    'Упаковочное оборудование',
]


# ============================================================
# НОРМАЛИЗАЦИЯ ИМЁН
# ============================================================
def normalize_name(name: str) -> str:
    """Нормализует имя для поиска в БД"""
    if not name:
        return ""
    # Удаляем знаки препинания
    normalized = re.sub(r'[^\w\s\-]', '', name)
    # Нижний регистр
    normalized = normalized.lower()
    # Убираем лишние пробелы
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized.strip()


def get_display_name(name: str) -> str:
    """Возвращает красивое имя для вывода в техкарте"""
    # Убираем лишние пробелы и символы
    name = name.strip()
    # Заменяем множественные пробелы
    name = re.sub(r'\s+', ' ', name)
    return name


# ============================================================
# ФУНКЦИЯ ОПРЕДЕЛЕНИЯ СОСТОЯНИЯ ЧЕК-БОКСА
# ============================================================
def get_checkbox_state(sdt_element) -> bool:
    """Определяет состояние чек-бокса по символу внутри SDT"""
    texts = sdt_element.xpath('.//w:t', namespaces=NAMESPACES)
    for t in texts:
        if t.text:
            if '☒' in t.text:
                return True
            if '☐' in t.text:
                return False
    return False


# ============================================================
# ПАРСИНГ ЯЧЕЙКИ С SDT
# ============================================================
def parse_sdt_cell(cell, category: str) -> list:
    """Парсит ячейку с SDT, возвращает список объектов"""
    sdt_elements = cell.xpath('.//w:sdt', namespaces=NAMESPACES)
    if not sdt_elements:
        return []

    elements = []

    for para in cell.xpath('.//w:p', namespaces=NAMESPACES):
        for child in para.getchildren():
            tag = child.tag.split('}')[-1]

            if tag == 'sdt':
                elements.append({
                    'type': 'checkbox',
                    'checked': get_checkbox_state(child)
                })
            elif tag == 'r':
                texts = child.xpath('.//w:t', namespaces=NAMESPACES)
                for t in texts:
                    if t.text:
                        text = t.text.strip()
                        if text and text not in ['☒', '☐']:
                            elements.append({
                                'type': 'text',
                                'value': text
                            })

    if not elements:
        return []

    objects = []
    current_obj = None
    current_state = False
    text_parts = []

    for elem in elements:
        if elem['type'] == 'checkbox':
            if text_parts:
                full_text = clean_text(''.join(text_parts))
                if full_text:
                    if current_obj is None:
                        current_obj = {
                            'name': full_text,
                            'checked': current_state,
                            'modifiers': [],
                            'category': category
                        }
                        objects.append(current_obj)
                    else:
                        current_obj['modifiers'].append({
                            'name': full_text,
                            'checked': current_state
                        })
                text_parts = []
            current_state = elem['checked']
        elif elem['type'] == 'text':
            text_parts.append(elem['value'])

    # Последний текст
    if text_parts:
        full_text = clean_text(''.join(text_parts))
        if full_text:
            if current_obj is None:
                current_obj = {
                    'name': full_text,
                    'checked': current_state,
                    'modifiers': [],
                    'category': category
                }
                objects.append(current_obj)
            else:
                current_obj['modifiers'].append({
                    'name': full_text,
                    'checked': current_state
                })

    return objects


# ============================================================
# ПАРСИНГ ЯЧЕЙКИ С UNICODE
# ============================================================
def parse_unicode_cell(cell, category: str) -> list:
    """Парсит ячейку с Unicode-символами ☒/☐"""
    texts = cell.xpath('.//w:t', namespaces=NAMESPACES)
    full_text = ''.join([t.text for t in texts if t.text])

    if not full_text:
        return []

    objects = []
    parts = re.split(r'([☐☒])', full_text)

    for i in range(1, len(parts), 2):
        if i < len(parts):
            checkbox = parts[i]
            text = parts[i + 1].strip() if i + 1 < len(parts) else ""

            # Обрезаем до следующего чек-бокса
            next_cb = text.find('☒')
            if next_cb == -1:
                next_cb = text.find('☐')
            if next_cb != -1:
                text = text[:next_cb].strip()

            if text:
                objects.append({
                    'name': text,
                    'checked': (checkbox == '☒'),
                    'modifiers': [],
                    'category': category
                })

    return objects


def clean_text(text: str) -> str:
    """Очищает текст"""
    if not text:
        return ""
    text = ' '.join(text.split())
    return text.strip()


# ============================================================
# ОСНОВНАЯ ФУНКЦИЯ
# ============================================================
def export_all_objects(docx_path: str):
    """Экспортирует все объекты"""

    print(f"\n{'=' * 80}")
    print(f"📊 ПОЛНЫЙ ЭКСПОРТ ВСЕХ ОБЪЕКТОВ ИЗ ЧЕК-ЛИСТА")
    print(f"{'=' * 80}")
    print(f"📄 Файл: {docx_path}\n")

    all_objects = []

    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        with docx_zip.open('word/document.xml') as xml_file:
            root = etree.fromstring(xml_file.read())

            cells = root.xpath('.//w:tc', namespaces=NAMESPACES)

            current_category = "Без категории"

            for cell in cells:
                # Получаем текст ячейки
                texts = cell.xpath('.//w:t', namespaces=NAMESPACES)
                full_text = ''.join([t.text for t in texts if t.text])

                # Определяем категорию
                for cat in ALL_CATEGORIES:
                    if cat in full_text and ':' in full_text:
                        current_category = cat
                        break

                # Проверяем SDT
                sdt_elements = cell.xpath('.//w:sdt', namespaces=NAMESPACES)
                has_sdt = len(sdt_elements) > 0

                # Проверяем Unicode
                has_unicode = '☒' in full_text or '☐' in full_text

                if has_sdt:
                    objects = parse_sdt_cell(cell, current_category)
                    all_objects.extend(objects)
                elif has_unicode:
                    objects = parse_unicode_cell(cell, current_category)
                    all_objects.extend(objects)

    return all_objects


def print_formatted_list(objects: list):
    """Выводит отформатированный список объектов"""

    # Группируем по категориям
    by_category = defaultdict(list)
    for obj in objects:
        by_category[obj['category']].append(obj)

    total_objects = len(objects)
    total_checked = sum(1 for o in objects if o['checked'])
    total_modifiers = sum(len(o['modifiers']) for o in objects)
    total_checked_mods = sum(
        sum(1 for m in o['modifiers'] if m['checked'])
        for o in objects
    )

    print(f"\n{'=' * 80}")
    print(f"📊 ОБЩАЯ СТАТИСТИКА")
    print(f"{'=' * 80}")
    print(f"   Всего объектов: {total_objects}")
    print(f"   Отмечено объектов: {total_checked}")
    print(f"   Всего модификаторов: {total_modifiers}")
    print(f"   Отмечено модификаторов: {total_checked_mods}")

    print(f"\n{'=' * 80}")
    print(f"📋 ПОЛНЫЙ СПИСОК ОБЪЕКТОВ ПО КАТЕГОРИЯМ")
    print(f"{'=' * 80}")

    # Сортируем категории как в чек-листе
    category_order = [
        'Поверхности',
        'Сантехническое оборудование',
        'Санитарный пост',
        'Мебель',
        'Офисная техника',
        'Многоразовые резиновые СИЗ',
        'Бытовая техника',
        'Инвентарь, посуда и т.д.',
        'Моечный, уборочный инвентарь и оборудование',
        'Посудомоечное оборудование',
        'Холодильное оборудование',
        'Дозирующее оборудование',
        'Тепловое оборудование',
        'Технологическое оборудование',
        'Упаковочное оборудование',
    ]

    for cat in category_order:
        if cat in by_category:
            objects_in_cat = by_category[cat]
            checked_in_cat = sum(1 for o in objects_in_cat if o['checked'])

            print(f"\n{'─' * 80}")
            print(f"📁 {cat.upper()} ({len(objects_in_cat)} объектов, {checked_in_cat} отмечено)")
            print(f"{'─' * 80}")

            for obj in sorted(objects_in_cat, key=lambda x: x['name']):
                status = "✅" if obj['checked'] else "⬜"
                norm_name = normalize_name(obj['name'])
                disp_name = get_display_name(obj['name'])

                if obj['modifiers']:
                    mod_str = ", ".join([
                        f"{'✅' if m['checked'] else '⬜'}{m['name']}"
                        for m in obj['modifiers']
                    ])
                    print(f"{status} {obj['name']}")
                    print(f"   norm: {norm_name}")
                    print(f"   disp: {disp_name}")
                    print(f"   mods: {mod_str}")

                    # Выводим модификаторы отдельно
                    for m in obj['modifiers']:
                        m_status = "✅" if m['checked'] else "⬜"
                        full_name = f"{obj['name']} {m['name']}"
                        m_norm = normalize_name(full_name)
                        m_disp = get_display_name(full_name)
                        print(f"      └─ {m_status} {m['name']}")
                        print(f"         norm: {m_norm}")
                        print(f"         disp: {m_disp}")
                else:
                    print(f"{status} {obj['name']}")
                    print(f"   norm: {norm_name}")
                    print(f"   disp: {disp_name}")

    # Выводим сводку для БД
    print(f"\n{'=' * 80}")
    print(f"🗄️ СВОДКА ДЛЯ БАЗЫ ДАННЫХ")
    print(f"{'=' * 80}")

    unique_objects = set()
    for obj in objects:
        unique_objects.add(normalize_name(obj['name']))
        for m in obj['modifiers']:
            full_name = f"{obj['name']} {m['name']}"
            unique_objects.add(normalize_name(full_name))

    print(f"\n   Уникальных normalized_name: {len(unique_objects)}")

    # Список всех normalized_name
    print(f"\n📋 ВСЕ NORMALIZED_NAMES ({len(unique_objects)} шт.):")
    for name in sorted(unique_objects):
        print(f"   {name}")


def save_to_json(objects: list, output_path: str):
    """Сохраняет объекты в JSON"""
    data = []
    for obj in objects:
        obj_data = {
            'category': obj['category'],
            'name': obj['name'],
            'normalized_name': normalize_name(obj['name']),
            'display_name': get_display_name(obj['name']),
            'checked': obj['checked'],
            'modifiers': []
        }

        for m in obj['modifiers']:
            full_name = f"{obj['name']} {m['name']}"
            obj_data['modifiers'].append({
                'name': m['name'],
                'full_name': full_name,
                'normalized_name': normalize_name(full_name),
                'display_name': get_display_name(full_name),
                'checked': m['checked']
            })

        data.append(obj_data)

    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    print(f"\n✅ Сохранено в: {output_path}")


def main():
    import argparse
    parser_arg = argparse.ArgumentParser()
    parser_arg.add_argument("file", nargs="?", default="1.docx", help="Путь к файлу")
    parser_arg.add_argument("--json", help="Сохранить в JSON")
    args = parser_arg.parse_args()

    file_path = args.file
    if not Path(file_path).exists():
        print(f"❌ Файл не найден: {file_path}")
        return

    objects = export_all_objects(file_path)
    print_formatted_list(objects)

    if args.json:
        save_to_json(objects, args.json)

    print(f"\n{'=' * 80}")
    print(f"✅ ЭКСПОРТ ЗАВЕРШЁН")
    print(f"{'=' * 80}")


if __name__ == "__main__":
    main()