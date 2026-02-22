"""
Парсер с двумя стоп-сигналами: чек-бокс ИЛИ граница ячейки
"""

import zipfile
from pathlib import Path
from lxml import etree
import re

from parser.models import ChecklistData, ChecklistItem, Category


class DualStopParser:
    """Парсер, который останавливается на чек-боксе ИЛИ границе ячейки"""

    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
    }

    def parse(self, file_path: str) -> ChecklistData:
        file_path = Path(file_path)
        data = ChecklistData(file_path=str(file_path))

        print(f"\n🔍 Открываем архив: {file_path}")

        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            with docx_zip.open('word/document.xml') as xml_file:
                xml_content = xml_file.read()
                root = etree.fromstring(xml_content)

                # Находим все ячейки таблиц
                cells = root.xpath('.//w:tc', namespaces=self.NAMESPACES)
                print(f"📊 Найдено ячеек: {len(cells)}")

                for cell in cells:
                    self._process_cell(cell, data)

        print(f"\n✅ Найдено элементов: {len(data.items)}")
        return data

    def _process_cell(self, cell, data: ChecklistData):
        """Обрабатывает одну ячейку таблицы"""

        # Получаем все текстовые элементы в ячейке с их позициями
        text_elements = cell.xpath('.//w:t', namespaces=self.NAMESPACES)

        # Собираем полный текст ячейки для контекста
        full_text = ''.join([t.text for t in text_elements if t.text])

        # Если в ячейке есть ☒
        if '☒' in full_text:
            # Разбиваем по чек-боксам внутри ячейки
            parts = re.split(r'([☐☒])', full_text)

            i = 0
            while i < len(parts):
                if parts[i] == '☒':  # Нашли отмеченный
                    if i + 1 < len(parts):
                        # Берём текст до следующего чек-бокса или до конца ячейки
                        next_text = parts[i + 1]

                        # Ищем следующий чек-бокс в этом же тексте
                        next_checkbox_pos = -1
                        for j, char in enumerate(next_text):
                            if char in ['☐', '☒']:
                                next_checkbox_pos = j
                                break

                        if next_checkbox_pos != -1:
                            # Обрезаем до следующего чек-бокса
                            item_text = next_text[:next_checkbox_pos].strip()
                        else:
                            # Берём весь текст до конца ячейки
                            item_text = next_text.strip()

                        # Очищаем
                        item_text = re.sub(r'[^\w\s\-\(\)]', '', item_text)
                        item_text = re.sub(r'\s+', ' ', item_text).strip()

                        if item_text and len(item_text) > 1:
                            print(f"  ✅ {item_text}")
                            data.items.append(ChecklistItem(
                                name=item_text,
                                category=Category.OTHER,
                                checked=True
                            ))
                i += 1


def parse_checklist(file_path: str) -> ChecklistData:
    parser = DualStopParser()
    return parser.parse(file_path)