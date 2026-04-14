# find_svetilniki_in_xml.py
import zipfile
from lxml import etree
from pathlib import Path

file_path = "тарамоечная.docx"

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
}

with zipfile.ZipFile(file_path, 'r') as docx_zip:
    with docx_zip.open('word/document.xml') as xml_file:
        xml_content = xml_file.read()
        root = etree.fromstring(xml_content)

        # Ищем все текстовые элементы
        text_elements = root.xpath('.//w:t', namespaces=NAMESPACES)

        print("🔍 ПОИСК 'светильники' В XML")
        print("=" * 50)

        found = False
        for elem in text_elements:
            if elem.text and "светильники" in elem.text.lower():
                found = True
                print(f"✅ Найдено: {repr(elem.text)}")
                # Показываем родительскую ячейку
                cell = elem.xpath('./ancestor::w:tc', namespaces=NAMESPACES)
                if cell:
                    all_text = ''.join(cell[0].xpath('.//w:t/text()', namespaces=NAMESPACES))
                    print(f"   Полный текст ячейки: {repr(all_text)}")

        if not found:
            print("❌ 'светильники' не найдены в XML")