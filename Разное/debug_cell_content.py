# debug_cell_content.py
import zipfile
from lxml import etree
from pathlib import Path

file_path = "Чек-лист проблемные.docx"

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}

with zipfile.ZipFile(file_path, 'r') as docx_zip:
    with docx_zip.open('word/document.xml') as xml_file:
        xml_content = xml_file.read()
        root = etree.fromstring(xml_content)

        # Ищем все ячейки
        cells = root.xpath('.//w:tc', namespaces=NAMESPACES)

        print("🔍 ПОИСК ЯЧЕЙКИ С 'камер'")
        print("=" * 50)

        for idx, cell in enumerate(cells):
            text = ''.join(cell.xpath('.//w:t/text()', namespaces=NAMESPACES))
            if 'камер' in text:
                print(f"\n📄 Ячейка {idx}:")
                print(f"   Текст: {repr(text)}")
                print(f"   Длина: {len(text)}")
                # Показываем символы
                for i, ch in enumerate(text):
                    if ch in ['☐', '☒']:
                        print(f"   Позиция {i}: '{ch}'")