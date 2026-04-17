#!/usr/bin/env python3
"""
test_sdt_parser_debug.py

Версия с подробной отладкой каждой ячейки.
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


def debug_parse_cell(cell, cell_idx):
    """Детальный разбор одной ячейки"""
    print(f"\n{'=' * 60}")
    print(f"🔍 ЯЧЕЙКА #{cell_idx}")
    print(f"{'=' * 60}")

    children = cell.getchildren()
    print(f"Количество дочерних элементов: {len(children)}")

    for i, child in enumerate(children):
        tag = child.tag.split('}')[-1]
        print(f"\n  [{i}] Тег: <{tag}>")

        if tag == 'sdt':
            # Чек-бокс
            checkbox = child.xpath('.//w14:checkbox', namespaces=NAMESPACES)
            if checkbox:
                checked = checkbox[0].get('{http://schemas.microsoft.com/office/word/2010/wordml}val')
                print(f"      ✅ Чек-бокс, checked={checked}")
            else:
                print(f"      ⚠️ SDT без чек-бокса")

            # Ищем текст внутри sdt
            sdt_texts = child.xpath('.//w:t', namespaces=NAMESPACES)
            if sdt_texts:
                text = ''.join([t.text for t in sdt_texts if t.text])
                print(f"      Текст внутри SDT: '{text}'")

        elif tag == 'r':
            # Текстовый элемент
            texts = child.xpath('.//w:t', namespaces=NAMESPACES)
            if texts:
                text = ''.join([t.text for t in texts if t.text])
                print(f"      Текст: '{text}'")

        elif tag == 'p':
            # Параграф
            texts = child.xpath('.//w:t', namespaces=NAMESPACES)
            if texts:
                text = ''.join([t.text for t in texts if t.text])
                print(f"      Текст (в параграфе): '{text}'")

        else:
            # Другие элементы
            texts = child.xpath('.//w:t', namespaces=NAMESPACES)
            if texts:
                text = ''.join([t.text for t in texts if t.text])
                print(f"      Текст (в {tag}): '{text}'")

    # Полный текст ячейки
    all_texts = cell.xpath('.//w:t', namespaces=NAMESPACES)
    full_text = ''.join([t.text for t in all_texts if t.text])
    print(f"\n  📝 ПОЛНЫЙ ТЕКСТ ЯЧЕЙКИ: '{full_text}'")


def main():
    import argparse
    parser_arg = argparse.ArgumentParser()
    parser_arg.add_argument("file", nargs="?", default="1.docx", help="Путь к файлу")
    parser_arg.add_argument("--cells", type=int, default=5, help="Количество ячеек для отладки")
    args = parser_arg.parse_args()

    file_path = args.file
    if not Path(file_path).exists():
        print(f"❌ Файл не найден: {file_path}")
        print("\nДоступные .docx файлы в текущей папке:")
        for f in Path(".").glob("*.docx"):
            if not f.name.startswith("~"):
                print(f"   - {f.name}")
        return

    print(f"\n{'=' * 80}")
    print(f"🔧 ДЕТАЛЬНАЯ ОТЛАДКА ПАРСИНГА")
    print(f"{'=' * 80}")
    print(f"📄 Файл: {file_path}\n")

    with zipfile.ZipFile(file_path, 'r') as docx_zip:
        with docx_zip.open('word/document.xml') as xml_file:
            xml_content = xml_file.read()
            root = etree.fromstring(xml_content)

            cells = root.xpath('.//w:tc', namespaces=NAMESPACES)
            print(f"📊 Всего ячеек: {len(cells)}")

            # Находим ячейки с SDT
            sdt_cells = []
            for idx, cell in enumerate(cells):
                sdt_elements = cell.xpath('.//w:sdt', namespaces=NAMESPACES)
                if sdt_elements:
                    sdt_cells.append((idx, cell))

            print(f"📊 Ячеек с SDT: {len(sdt_cells)}")
            print(f"\n🔍 АНАЛИЗ ПЕРВЫХ {min(args.cells, len(sdt_cells))} ЯЧЕЕК С SDT:")

            for i, (idx, cell) in enumerate(sdt_cells[:args.cells]):
                debug_parse_cell(cell, idx)

    print(f"\n{'=' * 80}")
    print(f"✅ ОТЛАДКА ЗАВЕРШЕНА")
    print(f"{'=' * 80}")


if __name__ == "__main__":
    main()