#!/usr/bin/env python3
"""
parser_complete.py

ПОЛНЫЙ ПАРСЕР - находит ВСЕ объекты (и с модификаторами, и без)
"""

import zipfile
import sys
from pathlib import Path
from lxml import etree

sys.path.insert(0, str(Path(__file__).parent))

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}


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


def parse_single_cell(cell) -> dict | None:
    """Парсит ОДНУ ячейку"""

    # Проверяем, есть ли SDT
    sdt_elements = cell.xpath('.//w:sdt', namespaces=NAMESPACES)
    if not sdt_elements:
        return None

    # Собираем все элементы в плоский список
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
        return None

    # Собираем объект
    current_item = None
    current_checkbox_state = False
    text_parts = []

    for elem in elements:
        if elem['type'] == 'checkbox':
            # Сохраняем накопленный текст ПЕРЕД новым чек-боксом
            if text_parts:
                full_text = clean_text(''.join(text_parts))
                if full_text:
                    if current_item is None:
                        current_item = {
                            'name': full_text,
                            'checked': current_checkbox_state,
                            'modifiers': []
                        }
                    else:
                        current_item['modifiers'].append({
                            'name': full_text,
                            'checked': current_checkbox_state
                        })
                text_parts = []

            current_checkbox_state = elem['checked']

        elif elem['type'] == 'text':
            text_parts.append(elem['value'])

    # СОХРАНЯЕМ ПОСЛЕДНИЙ ТЕКСТ (это важно для объектов без модификаторов!)
    if text_parts:
        full_text = clean_text(''.join(text_parts))
        if full_text:
            if current_item is None:
                # Объект без модификаторов
                current_item = {
                    'name': full_text,
                    'checked': current_checkbox_state,
                    'modifiers': []
                }
            else:
                # Последний модификатор
                current_item['modifiers'].append({
                    'name': full_text,
                    'checked': current_checkbox_state
                })

    return current_item


def clean_text(text: str) -> str:
    """Очищает текст"""
    if not text:
        return ""
    text = ' '.join(text.split())
    return text.strip()


def parse_document(docx_path: str) -> list:
    """Парсит документ"""
    all_items = []

    print(f"\n{'=' * 80}")
    print(f"🔧 ПОЛНЫЙ ПАРСЕР (находит ВСЕ объекты)")
    print(f"{'=' * 80}")
    print(f"📄 Файл: {docx_path}\n")

    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        with docx_zip.open('word/document.xml') as xml_file:
            root = etree.fromstring(xml_file.read())

            cells = root.xpath('.//w:tc', namespaces=NAMESPACES)
            print(f"📊 Всего ячеек: {len(cells)}")

            cells_with_sdt = 0

            for cell in cells:
                item = parse_single_cell(cell)
                if item:
                    cells_with_sdt += 1
                    all_items.append(item)

            print(f"📊 Ячеек с SDT: {cells_with_sdt}")
            print(f"📊 Всего объектов: {len(all_items)}")

    return all_items


def print_results(items: list):
    """Выводит результаты"""
    checked = [i for i in items if i['checked']]
    checked_mods = sum(1 for i in items for m in i['modifiers'] if m['checked'])

    print(f"\n{'=' * 80}")
    print(f"📊 СТАТИСТИКА")
    print(f"{'=' * 80}")
    print(f"   Всего объектов: {len(items)}")
    print(f"   Отмечено объектов: {len(checked)}")
    print(f"   Отмечено модификаторов: {checked_mods}")

    # Ключевые объекты
    print(f"\n{'=' * 80}")
    print(f"🔍 КЛЮЧЕВЫЕ ОБЪЕКТЫ")
    print(f"{'=' * 80}")

    targets = ['пол', 'лестницы', 'потолок', 'стены', 'весы', 'двери', 'окна']

    for key in targets:
        found = []
        for item in items:
            if key in item['name'].lower():
                found.append(item)

        if found:
            for item in found:
                status = "✅" if item['checked'] else "⬜"
                print(f"\n{status} {item['name']}")
                for mod in item['modifiers']:
                    mod_status = "✅" if mod['checked'] else "⬜"
                    print(f"   └─ {mod_status} {mod['name']}")
        else:
            print(f"\n❌ '{key}' не найден")

    # Показываем первые 30 объектов для проверки
    print(f"\n{'=' * 80}")
    print(f"📋 ПЕРВЫЕ 30 ОБЪЕКТОВ")
    print(f"{'=' * 80}")

    for item in items[:30]:
        status = "✅" if item['checked'] else "⬜"
        if item['modifiers']:
            mods = ", ".join([f"{'✅' if m['checked'] else '⬜'}{m['name']}" for m in item['modifiers']])
            print(f"{status} {item['name']} [{mods}]")
        else:
            print(f"{status} {item['name']}")


def main():
    import argparse
    parser_arg = argparse.ArgumentParser()
    parser_arg.add_argument("file", help="Путь к файлу")
    args = parser_arg.parse_args()

    file_path = args.file
    if not Path(file_path).exists():
        print(f"❌ Файл не найден: {file_path}")
        return

    items = parse_document(file_path)
    print_results(items)

    print(f"\n{'=' * 80}")
    print(f"✅ ПАРСИНГ ЗАВЕРШЁН")
    print(f"{'=' * 80}")


if __name__ == "__main__":
    main()