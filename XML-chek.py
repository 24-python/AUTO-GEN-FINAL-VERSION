import zipfile
from lxml import etree
from pathlib import Path


def analyze_docx_structure(docx_path):
    """Анализирует структуру XML документа"""

    docx_path = Path(docx_path)
    print(f"\n🔍 Анализ файла: {docx_path}")

    with zipfile.ZipFile(docx_path, 'r') as docx_zip:
        # Проверяем наличие document.xml
        if 'word/document.xml' not in docx_zip.namelist():
            print("❌ Не найден document.xml")
            return

        with docx_zip.open('word/document.xml') as xml_file:
            xml_content = xml_file.read()
            root = etree.fromstring(xml_content)

            # Пространства имен
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
                  'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'}

            # 1. Общая статистика
            paragraphs = root.xpath('.//w:p', namespaces=ns)
            tables = root.xpath('.//w:tbl', namespaces=ns)
            cells = root.xpath('.//w:tc', namespaces=ns)
            texts = root.xpath('.//w:t', namespaces=ns)

            print(f"\n📊 Статистика документа:")
            print(f"  - Параграфов: {len(paragraphs)}")
            print(f"  - Таблиц: {len(tables)}")
            print(f"  - Ячеек таблиц: {len(cells)}")
            print(f"  - Текстовых элементов: {len(texts)}")

            # 2. Поиск чек-боксов
            checkboxes_old = root.xpath('.//w:checkBox', namespaces=ns)
            checkboxes_new = root.xpath('.//w14:checkbox', namespaces=ns)

            print(f"\n🔲 Чек-боксы:")
            print(f"  - Старый формат (w:checkBox): {len(checkboxes_old)}")
            print(f"  - Новый формат (w14:checkbox): {len(checkboxes_new)}")

            # 3. Анализ первых 5 ячеек
            print(f"\n📋 Первые 5 ячеек:")
            for i, cell in enumerate(cells[:5]):
                cell_texts = cell.xpath('.//w:t', namespaces=ns)
                full_text = ''.join([t.text for t in cell_texts if t.text])
                print(f"\n  Ячейка {i + 1}:")
                print(f"    Текст: {full_text[:100]}...")

                # Проверяем объединение ячеек
                grid_span = cell.xpath('.//w:gridSpan', namespaces=ns)
                if grid_span:
                    span_val = grid_span[0].get('{' + ns['w'] + '}val')
                    print(f"    Объединение колонок: {span_val}")

            # 4. Сохраняем полный XML для изучения
            output_file = docx_path.with_suffix('.xml')
            with open(output_file, 'wb') as f:
                f.write(etree.tostring(root, pretty_print=True, encoding='utf-8'))
            print(f"\n💾 Полный XML сохранён в: {output_file}")

            return root


# Использование
analyze_docx_structure('uploads/20260222_121922_111f85b2_Чек-лист.docx')