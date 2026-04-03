"""
Генератор технологических карт из шаблона .docx
"""

from docx import Document
from docx.shared import Pt
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from pathlib import Path
from typing import List, Tuple

from parser.models import ChecklistData, ChecklistItem
from db.models import Instruction
from db.database import SessionLocal
from .mapper import find_object_and_instructions


class TechCardGenerator:
    """Генератор технологических карт"""

    # Путь к папке с шаблонами техкарт
    TEMPLATES_DIR = Path("tech_card_templates")
    DEFAULT_TEMPLATE = TEMPLATES_DIR / "шаблон.docx"

    def __init__(self, template_path: str = None):
        if template_path:
            self.template_path = Path(template_path)
        else:
            self.template_path = self.DEFAULT_TEMPLATE

    def generate(self, checklist_data: ChecklistData, output_path: str) -> str:
        """
        Генерирует технологическую карту из данных чек-листа

        Args:
            checklist_data: данные из парсера
            output_path: путь для сохранения результата

        Returns:
            путь к сохраненному файлу
        """
        # Загружаем шаблон
        if not self.template_path.exists():
            raise FileNotFoundError(f"Шаблон не найден: {self.template_path}")

        doc = Document(self.template_path)

        # 1. Заполняем шапку
        self._fill_header(doc, checklist_data)

        # 2. Получаем инструкции для отмеченных объектов
        objects_with_instructions = self._fetch_instructions(checklist_data)

        # 3. Заполняем таблицу
        self._fill_table(doc, objects_with_instructions)

        # 4. Сохраняем
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        doc.save(str(output_file))

        return str(output_file)

    def _fill_header(self, doc: Document, data: ChecklistData):
        """Заполняет заголовок и поле помещения"""
        for paragraph in doc.paragraphs:
            if "Технологическая карта санитарной обработки №" in paragraph.text:
                # Оставляем как есть или можно добавить номер
                pass
            if "Помещение:" in paragraph.text:
                paragraph.text = f"Помещение: {data.room_name or '______________'}"

    def _fetch_instructions(self, data: ChecklistData) -> List[Tuple[ChecklistItem, List[Instruction]]]:
        """
        Для каждого отмеченного объекта из чек-листа ищет инструкции в БД
        """
        session = SessionLocal()
        result = []

        for item in data.get_checked_items():
            obj, instructions = find_object_and_instructions(session, item.name, item.category)
            if obj:
                result.append((item, instructions))
            else:
                # Объект не найден в БД — добавляем с пустыми инструкциями
                result.append((item, []))

        session.close()
        return result

    def _fill_table(self, doc: Document, objects_with_instructions: List[Tuple[ChecklistItem, List[Instruction]]]):
        """
        Заполняет таблицу техкарты

        Логика:
        - Для каждого объекта может быть несколько инструкций
        - Первый столбец объединяется для всех инструкций одного объекта
        - Остальные столбцы заполняются для каждой инструкции отдельно
        """
        # Находим таблицу (первая таблица в документе)
        if not doc.tables:
            return
        table = doc.tables[0]

        # Удаляем пустую строку-образец (последняя строка)
        if len(table.rows) > 1:
            self._delete_row(table, len(table.rows) - 1)

        current_row = 1  # Начинаем после заголовков (индекс 0 - заголовки)

        for item, instructions in objects_with_instructions:
            if not instructions:
                # Нет инструкций — добавляем одну строку с пустыми данными
                self._add_empty_row(table, current_row, item.name)
                current_row += 1
                continue

            # Есть инструкции — добавляем строки для каждой
            start_row = current_row
            for idx, instr in enumerate(instructions):
                self._add_instruction_row(table, current_row, item.name if idx == 0 else None, instr)
                current_row += 1
            end_row = current_row - 1

            # Объединяем первый столбец для всех строк этого объекта
            if start_row != end_row:
                self._merge_cells_vertical(table, 0, start_row, end_row)

    def _add_instruction_row(self, table, row_index: int, object_name: str | None, instruction: Instruction):
        """Добавляет строку с данными инструкции"""
        if row_index >= len(table.rows):
            table.add_row()

        row = table.rows[row_index]

        # Первый столбец: название объекта (только для первой строки объекта)
        if object_name:
            row.cells[0].text = object_name
        else:
            row.cells[0].text = ""

        # Остальные столбцы: данные из инструкции
        self._set_cell_text(row.cells[1], instruction.cleaning_method or "")
        self._set_cell_text(row.cells[2], instruction.instruction_number or "")
        self._set_cell_text(row.cells[3], instruction.product_name or "")
        self._set_cell_text(row.cells[4], instruction.cleaning_technique or "")
        self._set_cell_text(row.cells[5], instruction.concentration or "")
        self._set_cell_text(row.cells[6], instruction.temperature or "")
        self._set_cell_text(row.cells[7], instruction.exposure_time or "")
        self._set_cell_text(row.cells[8], instruction.inventory or "")
        self._set_cell_text(row.cells[9], instruction.frequency or "")
        self._set_cell_text(row.cells[10], instruction.executor or "")
        self._set_cell_text(row.cells[11], instruction.control_method or "")

    def _add_empty_row(self, table, row_index: int, object_name: str):
        """Добавляет строку с пустыми данными (для объектов без инструкций)"""
        if row_index >= len(table.rows):
            table.add_row()

        row = table.rows[row_index]
        row.cells[0].text = object_name

        for col in range(1, 12):
            row.cells[col].text = "_____________"

    def _set_cell_text(self, cell, text: str):
        """Устанавливает текст в ячейку с сохранением форматирования"""
        cell.text = text
        # Сохраняем базовое форматирование (шрифт, размер)
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.size = Pt(10)

    def _delete_row(self, table, row_idx: int):
        """Удаляет строку из таблицы"""
        tbl = table._tbl
        tr = table.rows[row_idx]._tr
        tbl.remove(tr)

    def _merge_cells_vertical(self, table, col_idx: int, start_row: int, end_row: int):
        """
        Объединяет ячейки в столбце col_idx от start_row до end_row
        """
        if start_row >= end_row:
            return

        # Получаем первую ячейку для объединения
        top_cell = table.cell(start_row, col_idx)

        # Формируем строку с диапазоном для объединения
        tc = top_cell._tc
        tcPr = tc.get_or_add_tcPr()

        # Создаем элемент vMerge для объединения
        vMerge = OxmlElement('w:vMerge')
        vMerge.set(qn('w:val'), 'restart')
        tcPr.append(vMerge)

        # Для остальных ячеек
        for row_idx in range(start_row + 1, end_row + 1):
            cell = table.cell(row_idx, col_idx)
            tc = cell._tc
            tcPr = tc.get_or_add_tcPr()
            vMerge = OxmlElement('w:vMerge')
            tcPr.append(vMerge)