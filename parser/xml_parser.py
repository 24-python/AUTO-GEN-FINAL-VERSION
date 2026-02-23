"""
Парсер с определением категорий по ключевым словам
"""

import zipfile
from pathlib import Path
from lxml import etree
import re

from parser.models import ChecklistData, ChecklistItem, Category


class CategoryParser:
    """Парсер с определением категорий по ключевым словам"""

    NAMESPACES = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml'
    }

    # Полный словарь категорий по ключевым словам
    KEYWORD_CATEGORIES = [
        # Бытовая техника
        ('чайники', Category.HOUSEHOLD_APPLIANCES),
        ('термопоты', Category.HOUSEHOLD_APPLIANCES),
        ('кулера', Category.HOUSEHOLD_APPLIANCES),
        ('микроволновки', Category.HOUSEHOLD_APPLIANCES),
        ('холодильники', Category.HOUSEHOLD_APPLIANCES),
        ('стиральные машины', Category.HOUSEHOLD_APPLIANCES),
        ('сушильные машины', Category.HOUSEHOLD_APPLIANCES),

        # Упаковочное оборудование
        ('вакуумные упаковщики', Category.PACKAGING_EQUIPMENT),
        ('запайщики', Category.PACKAGING_EQUIPMENT),
        ('этикетировщики', Category.PACKAGING_EQUIPMENT),
        ('термоусадочные туннели', Category.PACKAGING_EQUIPMENT),
        ('автоматический клипсаторы', Category.PACKAGING_EQUIPMENT),
        ('flow pack', Category.PACKAGING_EQUIPMENT),

        # Тепловое оборудование
        ('плиты', Category.THERMAL_EQUIPMENT),
        ('варочные котлы', Category.THERMAL_EQUIPMENT),
        ('сковороды', Category.THERMAL_EQUIPMENT),
        ('фритюры', Category.THERMAL_EQUIPMENT),
        ('грили', Category.THERMAL_EQUIPMENT),
        ('пароконвектоматы', Category.THERMAL_EQUIPMENT),
        ('печи подовые', Category.THERMAL_EQUIPMENT),
        ('печи ротационные', Category.THERMAL_EQUIPMENT),

        # Инвентарь, посуда
        ('доски', Category.INVENTORY),
        ('посуда', Category.INVENTORY),
        ('ножи', Category.INVENTORY),
        ('мусаты', Category.INVENTORY),
        ('секачи', Category.INVENTORY),
        ('шампура', Category.INVENTORY),
        ('крючки', Category.INVENTORY),
        ('вешала', Category.INVENTORY),
        ('гастроёмкости', Category.INVENTORY),
        ('формы для выпечки', Category.INVENTORY),
        ('листы для выпечки', Category.INVENTORY),
        ('съёмные детали оборудования', Category.INVENTORY),
        ('силапеновые коврики', Category.INVENTORY),
        ('внутрицеховая тара', Category.INVENTORY),
        ('дежи', Category.INVENTORY),
        ('ёмкости для перетаривания', Category.INVENTORY),
        ('корзины для расстойки теста', Category.INVENTORY),
        ('отсадочные мешки', Category.INVENTORY),

        # Поверхности
        ('потолок', Category.SURFACE),
        ('светильники', Category.SURFACE),
        ('вытяжные зонты', Category.SURFACE),
        ('вентиляционные трубы', Category.SURFACE),
        ('диффузоры', Category.SURFACE),
        ('кондиционер', Category.SURFACE),
        ('сплит-системы', Category.SURFACE),
        ('инсектицидные лампы', Category.SURFACE),
        ('стены', Category.SURFACE),
        ('бактерицидные лампы', Category.SURFACE),
        ('стерилизатор', Category.SURFACE),
        ('металлоконструкции', Category.SURFACE),
        ('кабель-каналы', Category.SURFACE),
        ('водопроводные трубы', Category.SURFACE),
        ('выключатели', Category.SURFACE),
        ('розетки', Category.SURFACE),
        ('дверные ручки', Category.SURFACE),
        ('окна внешние', Category.SURFACE),
        ('окна внутрицеховые', Category.SURFACE),
        ('электрощиты', Category.SURFACE),
        ('пожарный щит', Category.SURFACE),
        ('отопительные приборы', Category.SURFACE),
        ('бойлера', Category.SURFACE),
        ('двери', Category.SURFACE),
        ('ворота/роллета', Category.SURFACE),
        ('завесы ПВХ', Category.SURFACE),
        ('завесы тепловые', Category.SURFACE),
        ('демосистема', Category.SURFACE),
        ('лестницы', Category.SURFACE),
        ('пол', Category.SURFACE),

        # Моечный инвентарь
        ('мопы', Category.CLEANING_EQUIPMENT),
        ('ветошь', Category.CLEANING_EQUIPMENT),
        ('щётки', Category.CLEANING_EQUIPMENT),
        ('сгоны', Category.CLEANING_EQUIPMENT),
        ('ведра', Category.CLEANING_EQUIPMENT),
        ('тележки', Category.CLEANING_EQUIPMENT),
        ('АВД', Category.CLEANING_EQUIPMENT),
        ('пылесосы', Category.CLEANING_EQUIPMENT),
        ('поломоечная машина', Category.CLEANING_EQUIPMENT),

        # Технологическое оборудование
        ('производственные столы', Category.TECH_EQUIPMENT),
        ('весы', Category.TECH_EQUIPMENT),
        ('блендеры', Category.TECH_EQUIPMENT),
        ('миксеры планетарные', Category.TECH_EQUIPMENT),
        ('бисквиторезки', Category.TECH_EQUIPMENT),
        ('тестомесы', Category.TECH_EQUIPMENT),
        ('тестоделители', Category.TECH_EQUIPMENT),
        ('депозитор волюметрический', Category.TECH_EQUIPMENT),
        ('тестоокруглители', Category.TECH_EQUIPMENT),
        ('тестораскатки', Category.TECH_EQUIPMENT),
        ('прессы для теста', Category.TECH_EQUIPMENT),
        ('тарталетницы', Category.TECH_EQUIPMENT),
        ('пневматические распылители', Category.TECH_EQUIPMENT),
        ('термощупы', Category.TECH_EQUIPMENT),
        ('дробилки', Category.TECH_EQUIPMENT),
        ('машина для резки', Category.TECH_EQUIPMENT),
        ('овощерезки', Category.TECH_EQUIPMENT),
        ('овощечистки', Category.TECH_EQUIPMENT),
        ('измельчители', Category.TECH_EQUIPMENT),
        ('слайсера', Category.TECH_EQUIPMENT),
        ('протирочные машины', Category.TECH_EQUIPMENT),
        ('картофелечистки', Category.TECH_EQUIPMENT),
        ('металлодетектор', Category.TECH_EQUIPMENT),
        ('рентгеновские системы контроля', Category.TECH_EQUIPMENT),
        ('вакуумные роторные шприцы', Category.TECH_EQUIPMENT),

        # Сантехника
        ('раковина', Category.PLUMBING),
        ('смеситель', Category.PLUMBING),
        ('унитаз', Category.PLUMBING),
        ('писсуар', Category.PLUMBING),
        ('душевая кабина', Category.PLUMBING),
        ('душевые секции', Category.PLUMBING),
        ('технологическая мойка', Category.PLUMBING),

        # Посудомоечное оборудование
        ('ПММ', Category.DISHWASHING_EQUIPMENT),
        ('купольная', Category.DISHWASHING_EQUIPMENT),
        ('туннельная', Category.DISHWASHING_EQUIPMENT),
        ('таромоечная машина', Category.DISHWASHING_EQUIPMENT),

        # Холодильное оборудование
        ('камеры', Category.REFRIGERATION_EQUIPMENT),
        ('морозильный ларь', Category.REFRIGERATION_EQUIPMENT),
        ('холод. столы', Category.REFRIGERATION_EQUIPMENT),
        ('холод. шкафы', Category.REFRIGERATION_EQUIPMENT),
        ('ледогенератор', Category.REFRIGERATION_EQUIPMENT),

        # Мебель
        ('столы', Category.FURNITURE),
        ('стулья', Category.FURNITURE),
        ('полки', Category.FURNITURE),
        ('стеллажи', Category.FURNITURE),
        ('шкафы', Category.FURNITURE),
        ('вешалки', Category.FURNITURE),
        ('скамейки', Category.FURNITURE),
        ('зеркала', Category.FURNITURE),

        # Офисная техника
        ('компьютеры', Category.OFFICE_EQUIPMENT),
        ('мониторы', Category.OFFICE_EQUIPMENT),
        ('принтера', Category.OFFICE_EQUIPMENT),
        ('телефоны', Category.OFFICE_EQUIPMENT),

        # СИЗ
        ('фартуки', Category.PPE),
        ('сапоги', Category.PPE),
        ('перчатки', Category.PPE),
        ('нарукавники', Category.PPE),
    ]

    def __init__(self):
        self.current_category = Category.OTHER

    def parse(self, file_path: str) -> ChecklistData:
        file_path = Path(file_path)
        data = ChecklistData(file_path=str(file_path))
        self.current_category = Category.OTHER

        print(f"\n🔍 Открываем архив: {file_path}")

        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            with docx_zip.open('word/document.xml') as xml_file:
                xml_content = xml_file.read()
                root = etree.fromstring(xml_content)

                # Парсим информацию об участке и предприятии из первой строки
                self._parse_header_from_first_row(root, data)

                # Находим все ячейки
                cells = root.xpath('.//w:tc', namespaces=self.NAMESPACES)
                print(f"📊 Найдено ячеек: {len(cells)}")

                for cell in cells:
                    self._process_cell(cell, data)

        self._print_statistics(data)
        print(f"\n✅ Найдено элементов: {len(data.items)}")

        return data

    def _parse_header_from_first_row(self, root, data: ChecklistData):
        """Парсит предприятие и участок из первой строки"""
        first_row = root.xpath('.//w:tr[1]', namespaces=self.NAMESPACES)

        if first_row:
            cells = first_row[0].xpath('.//w:tc', namespaces=self.NAMESPACES)

            if len(cells) >= 2:
                enterprise_text = self._get_cell_text(cells[0])
                if enterprise_text and not data.enterprise:
                    data.enterprise = enterprise_text.strip()
                    print(f"🏭 Найдено предприятие: {data.enterprise}")

                if len(cells) >= 2:
                    room_text = self._get_cell_text(cells[1])
                    if room_text and not data.room_name:
                        data.room_name = room_text.strip()
                        print(f"🏢 Найден участок: {data.room_name}")

    def _get_cell_text(self, cell) -> str:
        """Извлекает текст из ячейки"""
        texts = []
        for text_elem in cell.xpath('.//w:t', namespaces=self.NAMESPACES):
            if text_elem.text:
                texts.append(text_elem.text)
        return ' '.join(texts)

    def _get_category_by_keywords(self, item_text: str) -> Category:
        """Определяет категорию по ключевым словам"""
        item_lower = item_text.lower()
        for keyword, category in self.KEYWORD_CATEGORIES:
            if keyword in item_lower:
                return category
        return Category.OTHER

    def _process_cell(self, cell, data: ChecklistData):
        """Обрабатывает одну ячейку"""

        text_elements = cell.xpath('.//w:t', namespaces=self.NAMESPACES)
        full_text = ''.join([t.text for t in text_elements if t.text])

        # Если есть ☒
        if '☒' in full_text:
            # Разбиваем по чек-боксам
            parts = re.split(r'([☐☒])', full_text)

            i = 0
            while i < len(parts):
                if parts[i] == '☒':
                    if i + 1 < len(parts):
                        next_text = parts[i + 1]

                        # Ищем следующий чек-бокс
                        next_checkbox_pos = -1
                        for j, char in enumerate(next_text):
                            if char in ['☐', '☒']:
                                next_checkbox_pos = j
                                break

                        if next_checkbox_pos != -1:
                            item_text = next_text[:next_checkbox_pos].strip()
                        else:
                            item_text = next_text.strip()

                        # Очищаем
                        item_text = re.sub(r'[^\w\s\-\(\)]', '', item_text)
                        item_text = re.sub(r'\s+', ' ', item_text).strip()

                        # Убираем одиночные буквы в конце
                        item_text = re.sub(r'\s+[А-Я]$', '', item_text)

                        if item_text and len(item_text) > 1:
                            # Определяем категорию по ключевым словам
                            category = self._get_category_by_keywords(item_text)

                            print(f"  [{category.value}] ✅ {item_text}")
                            data.items.append(ChecklistItem(
                                name=item_text,
                                category=category,
                                checked=True
                            ))
                i += 1

    def _print_statistics(self, data: ChecklistData):
        """Выводит статистику по категориям"""
        print("\n📊 Статистика по категориям:")

        grouped = data.group_checked_by_category()
        for category, items in grouped.items():
            print(f"  {category.value}: {len(items)} элементов")
            for item in items[:3]:
                print(f"    • {item.name}")
            if len(items) > 3:
                print(f"    ... и ещё {len(items) - 3}")


def parse_checklist(file_path: str) -> ChecklistData:
    parser = CategoryParser()
    return parser.parse(file_path)