#!/usr/bin/env python3
"""
count_all_objects.py

Подсчёт ВСЕХ объектов в чек-листе (с SDT и без).
"""

import zipfile
import re
import sys
from pathlib import Path
from lxml import etree

sys.path.insert(0, str(Path(__file__).parent))

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}


def count_objects_in_cell_sdt(cell) -> int:
    """Подсчёт объектов в ячейке с SDT"""
    sdt_elements = cell.xpath('.//w:sdt', namespaces=NAMESPACES)
    if not sdt_elements:
        return 0

    # Каждый первый чек-бокс в ячейке = 1 объект
    # (модификаторы не считаем отдельными объектами)
    return 1


def count_objects_in_cell_unicode(cell) -> int:
    """Подсчёт объектов в ячейке с символами ☒/☐"""
    texts = cell.xpath('.//w:t', namespaces=NAMESPACES)
    full_text = ''.join([t.text for t in texts if t.text])

    if not full_text:
        return 0

    # Считаем количество чек-боксов
    checked_count = full_text.count('☒')
    unchecked_count = full_text.count('☐')
    total = checked_count + unchecked_count

    # Если есть чек-боксы, но нет SDT, каждый чек-бокс = отдельный объект
    # (в таких ячейках обычно один объект на чек-бокс)
    return total


def analyze_document(docx_path: str):
    """Анализирует документ и считает все объекты"""

    print(f"\n{'=' * 80}")
    print(f"🔧 ПОДСЧЁТ ВСЕХ ОБЪЕКТОВ В ЧЕК-ЛИСТЕ")
    print(f"{'=' * 80}")
    print(f"📄 Файл: {docx_path}\n")

    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        with docx_zip.open('word/document.xml') as xml_file:
            root = etree.fromstring(xml_file.read())

            cells = root.xpath('.//w:tc', namespaces=NAMESPACES)
            print(f"📊 Всего ячеек в документе: {len(cells)}")

            cells_with_sdt = 0
            cells_with_unicode_only = 0
            cells_with_both = 0
            empty_cells = 0

            total_objects_sdt = 0
            total_objects_unicode = 0

            # Детальный анализ по категориям
            categories = {}
            current_category = "Без категории"

            for idx, cell in enumerate(cells):
                # Проверяем, не является ли ячейка заголовком категории
                texts = cell.xpath('.//w:t', namespaces=NAMESPACES)
                full_text = ''.join([t.text for t in texts if t.text])

                # Определяем категорию по заголовкам
                if ':' in full_text and full_text.endswith(':'):
                    if 'Поверхности' in full_text:
                        current_category = 'Поверхности'
                    elif 'Бытовая техника' in full_text:
                        current_category = 'Бытовая техника'
                    elif 'Тепловое оборудование' in full_text:
                        current_category = 'Тепловое оборудование'
                    elif 'Упаковочное оборудование' in full_text:
                        current_category = 'Упаковочное оборудование'
                    elif 'Инвентарь' in full_text:
                        current_category = 'Инвентарь'
                    elif 'Технологическое оборудование' in full_text:
                        current_category = 'Технологическое оборудование'
                    elif 'Холодильное оборудование' in full_text:
                        current_category = 'Холодильное оборудование'
                    elif 'Санитарный пост' in full_text:
                        current_category = 'Санитарный пост'
                    elif 'Мебель' in full_text:
                        current_category = 'Мебель'
                    elif 'Офисная техника' in full_text:
                        current_category = 'Офисная техника'
                    elif 'Многоразовые резиновые СИЗ' in full_text:
                        current_category = 'Многоразовые резиновые СИЗ'
                    elif 'Моечный' in full_text:
                        current_category = 'Моечный инвентарь'
                    elif 'Посудомоечное' in full_text:
                        current_category = 'Посудомоечное оборудование'
                    elif 'Сантехническое' in full_text:
                        current_category = 'Сантехническое оборудование'
                    elif 'Дозирующее' in full_text:
                        current_category = 'Дозирующее оборудование'

                # Проверяем SDT
                sdt_elements = cell.xpath('.//w:sdt', namespaces=NAMESPACES)
                has_sdt = len(sdt_elements) > 0

                # Проверяем Unicode чек-боксы
                has_unicode = '☒' in full_text or '☐' in full_text

                if has_sdt and has_unicode:
                    cells_with_both += 1
                    obj_count = count_objects_in_cell_sdt(cell)
                    total_objects_sdt += obj_count

                    if obj_count > 0:
                        if current_category not in categories:
                            categories[current_category] = {'sdt': 0, 'unicode': 0}
                        categories[current_category]['sdt'] += obj_count

                elif has_sdt:
                    cells_with_sdt += 1
                    obj_count = count_objects_in_cell_sdt(cell)
                    total_objects_sdt += obj_count

                    if obj_count > 0:
                        if current_category not in categories:
                            categories[current_category] = {'sdt': 0, 'unicode': 0}
                        categories[current_category]['sdt'] += obj_count

                elif has_unicode:
                    cells_with_unicode_only += 1
                    obj_count = count_objects_in_cell_unicode(cell)
                    total_objects_unicode += obj_count

                    if obj_count > 0:
                        if current_category not in categories:
                            categories[current_category] = {'sdt': 0, 'unicode': 0}
                        categories[current_category]['unicode'] += obj_count
                else:
                    empty_cells += 1

    # Вывод статистики
    print(f"\n{'=' * 80}")
    print(f"📊 СТАТИСТИКА ПО ЯЧЕЙКАМ")
    print(f"{'=' * 80}")
    print(f"   Ячеек с SDT: {cells_with_sdt}")
    print(f"   Ячеек с Unicode-символами: {cells_with_unicode_only}")
    print(f"   Ячеек с обоими типами: {cells_with_both}")
    print(f"   Пустых/заголовочных ячеек: {empty_cells}")

    total_cells_with_objects = cells_with_sdt + cells_with_unicode_only + cells_with_both
    print(f"\n   ВСЕГО ячеек с объектами: {total_cells_with_objects}")

    print(f"\n{'=' * 80}")
    print(f"📊 ПОДСЧЁТ ОБЪЕКТОВ")
    print(f"{'=' * 80}")
    print(f"   Объектов через SDT: {total_objects_sdt}")
    print(f"   Объектов через Unicode: {total_objects_unicode}")
    print(f"\n   🎯 ВСЕГО ОБЪЕКТОВ В ЧЕК-ЛИСТЕ: {total_objects_sdt + total_objects_unicode}")

    # По категориям
    print(f"\n{'=' * 80}")
    print(f"📊 РАСПРЕДЕЛЕНИЕ ПО КАТЕГОРИЯМ")
    print(f"{'=' * 80}")

    grand_total = 0
    for cat, counts in sorted(categories.items()):
        cat_total = counts['sdt'] + counts['unicode']
        grand_total += cat_total
        print(f"\n   {cat}:")
        print(f"      SDT: {counts['sdt']}")
        print(f"      Unicode: {counts['unicode']}")
        print(f"      ВСЕГО: {cat_total}")

    print(f"\n   {'=' * 40}")
    print(f"   ОБЩИЙ ИТОГ: {grand_total} объектов")

    return grand_total


def main():
    import argparse
    parser_arg = argparse.ArgumentParser()
    parser_arg.add_argument("file", nargs="?", default="1.docx", help="Путь к файлу")
    args = parser_arg.parse_args()

    file_path = args.file
    if not Path(file_path).exists():
        print(f"❌ Файл не найден: {file_path}")
        return

    total = analyze_document(file_path)

    print(f"\n{'=' * 80}")
    print(f"✅ ПОДСЧЁТ ЗАВЕРШЁН")
    print(f"{'=' * 80}")
    print(f"\n📌 ВСЕГО ОБЪЕКТОВ В ЧЕК-ЛИСТЕ: {total}")


if __name__ == "__main__":
    main()