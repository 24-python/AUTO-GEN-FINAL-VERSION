#!/usr/bin/env python3
"""
export_all_objects_txt.py

Выводит ПОЛНЫЙ список всех объектов из чек-листа и сохраняет в TXT.
"""

import zipfile
import re
import sys
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

# Порядок категорий для вывода
CATEGORY_ORDER = [
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
    name = name.strip()
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
# ЭКСПОРТ В TXT
# ============================================================
def export_to_txt(objects: list, output_path: str):
    """Сохраняет объекты в TXT файл"""

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

    # Собираем все уникальные normalized_name
    unique_names = set()
    for obj in objects:
        unique_names.add(normalize_name(obj['name']))
        for m in obj['modifiers']:
            full_name = f"{obj['name']} {m['name']}"
            unique_names.add(normalize_name(full_name))

    with open(output_path, 'w', encoding='utf-8') as f:
        # Заголовок
        f.write("=" * 100 + "\n")
        f.write("ПОЛНЫЙ ЭКСПОРТ ОБЪЕКТОВ ИЗ ЧЕК-ЛИСТА\n")
        f.write("=" * 100 + "\n\n")

        # Общая статистика
        f.write("ОБЩАЯ СТАТИСТИКА\n")
        f.write("-" * 50 + "\n")
        f.write(f"Всего объектов:              {total_objects}\n")
        f.write(f"Отмечено объектов (✅):       {total_checked}\n")
        f.write(f"Не отмечено объектов (⬜):    {total_objects - total_checked}\n")
        f.write(f"Всего модификаторов:          {total_modifiers}\n")
        f.write(f"Отмечено модификаторов (✅):  {total_checked_mods}\n")
        f.write(f"Не отмечено модификаторов:    {total_modifiers - total_checked_mods}\n")
        f.write(f"Уникальных normalized_name:   {len(unique_names)}\n")
        f.write("\n")

        # Список по категориям
        f.write("=" * 100 + "\n")
        f.write("ПОЛНЫЙ СПИСОК ОБЪЕКТОВ ПО КАТЕГОРИЯМ\n")
        f.write("=" * 100 + "\n\n")

        for cat in CATEGORY_ORDER:
            if cat in by_category:
                objects_in_cat = by_category[cat]
                checked_in_cat = sum(1 for o in objects_in_cat if o['checked'])

                f.write(f"\n{'─' * 100}\n")
                f.write(f"📁 {cat.upper()} ({len(objects_in_cat)} объектов, {checked_in_cat} отмечено)\n")
                f.write(f"{'─' * 100}\n\n")

                for obj in sorted(objects_in_cat, key=lambda x: x['name']):
                    status = "✅" if obj['checked'] else "⬜"
                    norm_name = normalize_name(obj['name'])
                    disp_name = get_display_name(obj['name'])

                    f.write(f"{status} {obj['name']}\n")
                    f.write(f"       norm: {norm_name}\n")
                    f.write(f"       disp: {disp_name}\n")

                    if obj['modifiers']:
                        f.write(f"       модификаторы:\n")
                        for m in obj['modifiers']:
                            m_status = "✅" if m['checked'] else "⬜"
                            full_name = f"{obj['name']} {m['name']}"
                            m_norm = normalize_name(full_name)
                            m_disp = get_display_name(full_name)

                            f.write(f"          └─ {m_status} {m['name']}\n")
                            f.write(f"             norm: {m_norm}\n")
                            f.write(f"             disp: {m_disp}\n")
                    f.write("\n")

        # Список всех normalized_name
        f.write("\n" + "=" * 100 + "\n")
        f.write(f"ВСЕ NORMALIZED_NAMES ({len(unique_names)} шт.)\n")
        f.write("=" * 100 + "\n\n")

        for name in sorted(unique_names):
            f.write(f"{name}\n")

        # Сводка по категориям (таблица)
        f.write("\n" + "=" * 100 + "\n")
        f.write("СВОДКА ПО КАТЕГОРИЯМ (ТАБЛИЦА)\n")
        f.write("=" * 100 + "\n\n")

        f.write(f"{'Категория':<40} {'Объектов':<10} {'Отмечено':<10} {'%':<10}\n")
        f.write("-" * 70 + "\n")

        for cat in CATEGORY_ORDER:
            if cat in by_category:
                objects_in_cat = by_category[cat]
                total = len(objects_in_cat)
                checked = sum(1 for o in objects_in_cat if o['checked'])
                percent = round(checked / total * 100) if total > 0 else 0
                f.write(f"{cat:<40} {total:<10} {checked:<10} {percent}%\n")

        # Статистика по типам
        f.write("\n" + "=" * 100 + "\n")
        f.write("СТАТИСТИКА ПО ТИПАМ ОБЪЕКТОВ\n")
        f.write("=" * 100 + "\n\n")

        with_mods = sum(1 for o in objects if o['modifiers'])
        without_mods = total_objects - with_mods

        f.write(f"Объектов с модификаторами:    {with_mods}\n")
        f.write(f"Объектов без модификаторов:    {without_mods}\n")

        # Модификаторы по типам
        mod_counts = defaultdict(int)
        for obj in objects:
            for m in obj['modifiers']:
                mod_counts[m['name']] += 1

        if mod_counts:
            f.write(f"\nЧастота модификаторов:\n")
            for mod_name, count in sorted(mod_counts.items(), key=lambda x: -x[1]):
                f.write(f"   {mod_name}: {count}\n")

    print(f"\n✅ Сохранено в: {output_path}")


def export_all_objects(docx_path: str) -> list:
    """Экспортирует все объекты из документа"""

    print(f"\n{'=' * 80}")
    print(f"📊 ЭКСПОРТ ВСЕХ ОБЪЕКТОВ ИЗ ЧЕК-ЛИСТА")
    print(f"{'=' * 80}")
    print(f"📄 Файл: {docx_path}\n")

    all_objects = []

    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        with docx_zip.open('word/document.xml') as xml_file:
            root = etree.fromstring(xml_file.read())

            cells = root.xpath('.//w:tc', namespaces=NAMESPACES)
            print(f"📊 Всего ячеек: {len(cells)}")

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

    print(f"📊 Найдено объектов: {len(all_objects)}")

    return all_objects


def main():
    import argparse
    parser_arg = argparse.ArgumentParser()
    parser_arg.add_argument("file", nargs="?", default="1.docx", help="Путь к файлу")
    parser_arg.add_argument("--output", "-o", default="objects_export.txt", help="Выходной TXT файл")
    args = parser_arg.parse_args()

    file_path = args.file
    if not Path(file_path).exists():
        print(f"❌ Файл не найден: {file_path}")
        return

    objects = export_all_objects(file_path)
    export_to_txt(objects, args.output)

    print(f"\n{'=' * 80}")
    print(f"✅ ЭКСПОРТ ЗАВЕРШЁН")
    print(f"{'=' * 80}")


if __name__ == "__main__":
    main()