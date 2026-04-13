# debug_svetilniki_parser.py
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from parser.xml_parser import parse_checklist
import zipfile
from lxml import etree

# Настройки
NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
}


def debug_cell_content(file_path, search_text="светильники"):
    """Анализирует содержимое ячейки с искомым текстом"""

    print(f"🔍 АНАЛИЗ ФАЙЛА: {file_path}")
    print("=" * 60)

    with zipfile.ZipFile(file_path, 'r') as docx_zip:
        with docx_zip.open('word/document.xml') as xml_file:
            xml_content = xml_file.read()
            root = etree.fromstring(xml_content)

            # Находим все ячейки
            cells = root.xpath('.//w:tc', namespaces=NAMESPACES)

            for idx, cell in enumerate(cells):
                # Получаем текст ячейки
                texts = cell.xpath('.//w:t/text()', namespaces=NAMESPACES)
                full_text = ''.join(texts)

                if search_text in full_text.lower():
                    print(f"\n📄 Ячейка {idx}:")
                    print(f"   Текст: {repr(full_text)}")
                    print(f"   Длина: {len(full_text)}")

                    # Позиции чек-боксов
                    for i, ch in enumerate(full_text):
                        if ch in ['☐', '☒']:
                            print(f"   Позиция {i}: '{ch}'")

                    # Структура XML ячейки
                    print(f"\n   XML структура:")
                    for child in cell.iter():
                        if child.tag in [f'{{{NAMESPACES["w"]}}}t',
                                         f'{{{NAMESPACES["w"]}}}tc',
                                         f'{{{NAMESPACES["w14"]}}}checkbox']:
                            tag = child.tag.split('}')[-1]
                            if child.text and child.text.strip():
                                print(f"      {tag}: {repr(child.text[:50])}")
                    break
            else:
                print(f"\n❌ Ячейка с '{search_text}' не найдена")


def run_parser(file_path):
    """Запускает парсер и показывает результат"""
    print(f"\n🔍 РЕЗУЛЬТАТ ПАРСЕРА")
    print("=" * 60)

    data = parse_checklist(file_path)

    print(f"\n✅ Найдено объектов: {len(data.get_checked_items())}")

    print(f"\n📋 ОБЪЕКТЫ СО 'СВЕТ'")
    for item in data.get_checked_items():
        if "свет" in item.name.lower():
            print(f"   {item.name} → категория: {item.category.value}")

    if not any("свет" in item.name.lower() for item in data.get_checked_items()):
        print("   ❌ Светильники не найдены")

    return data


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("file", help="Путь к .docx файлу")
    args = parser.parse_args()

    if not Path(args.file).exists():
        print(f"❌ Файл не найден: {args.file}")
        sys.exit(1)

    # Анализируем XML
    debug_cell_content(args.file)

    # Запускаем парсер
    run_parser(args.file)