from docx import Document
from docx.shared import Pt
from pathlib import Path
from parser.models import ChecklistData
from db.models import Instruction
from db.database import SessionLocal
from generator.mapper import find_object_and_instructions
from collections import defaultdict


class TechCardGenerator:
    TEMPLATES_DIR = Path("tech_card_templates")
    DEFAULT_TEMPLATE = TEMPLATES_DIR / "шаблон.docx"

    def __init__(self, template_path: str = None):
        self.template_path = Path(template_path) if template_path else self.DEFAULT_TEMPLATE

    def _set_cell_font(self, cell, text: str, font_name: str = 'Arial', size_pt: int = 9, bold: bool = False):
        """Устанавливает текст и шрифт в ячейке"""
        cell.text = text
        for paragraph in cell.paragraphs:
            for run in paragraph.runs:
                run.font.name = font_name
                run.font.size = Pt(size_pt)
                run.font.bold = bold

    def _apply_font_to_row(self, row, font_name: str = 'Arial', size_pt: int = 9, bold: bool = False):
        """Применяет шрифт ко всем ячейкам строки"""
        for col in range(len(row.cells)):
            for paragraph in row.cells[col].paragraphs:
                for run in paragraph.runs:
                    run.font.name = font_name
                    run.font.size = Pt(size_pt)
                    run.font.bold = bold

    def _merge_cells_horizontal(self, row, start_col: int, end_col: int):
        """Объединяет ячейки в строке по горизонтали"""
        if start_col >= end_col:
            return
        start_cell = row.cells[start_col]
        for col in range(start_col + 1, end_col + 1):
            start_cell.merge(row.cells[col])

    def generate(self, checklist_data: ChecklistData, output_path: str) -> str:
        doc = Document(self.template_path)

        # Заполняем помещение
        for para in doc.paragraphs:
            if "Помещение:" in para.text:
                para.text = f"Помещение: {checklist_data.room_name or '______________'}"
                for run in para.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(9)
                    run.font.bold = True
                break

        # Группируем объекты по категориям
        session = SessionLocal()
        category_items = defaultdict(list)

        for item in checklist_data.get_checked_items():
            obj, instructions = find_object_and_instructions(session, item.name, item.category)
            if instructions:
                for instr in instructions:
                    category_items[item.category.value].append((item.name, instr))
            else:
                category_items[item.category.value].append((item.name, None))
        session.close()

        # Заполняем таблицу
        table = doc.tables[0]
        start_row = 6  # после 5 объединенных строк + заголовки

        # Удаляем старые строки данных
        while len(table.rows) > start_row:
            tbl = table._tbl
            tbl.remove(table.rows[start_row]._tr)

        current_row = start_row

        # Проходим по категориям
        for category_name, items in category_items.items():
            # Добавляем строку с названием категории (BOLD, объединенная)
            if current_row >= len(table.rows):
                table.add_row()
            category_row = table.rows[current_row]

            # Объединяем все 12 колонок
            self._merge_cells_horizontal(category_row, 0, 11)

            # Устанавливаем текст категории с ЖИРНЫМ шрифтом Arial 9pt
            self._set_cell_font(category_row.cells[0], category_name, font_name='Arial', size_pt=9, bold=True)
            current_row += 1

            # Добавляем строки для каждого объекта в категории
            for obj_name, instr in items:
                if current_row >= len(table.rows):
                    table.add_row()
                row = table.rows[current_row]

                # Заполняем ячейки
                row.cells[0].text = obj_name
                if instr:
                    row.cells[1].text = instr.cleaning_method or ""
                    row.cells[2].text = instr.instruction_number or ""
                    row.cells[3].text = instr.product_name or ""
                    row.cells[4].text = instr.cleaning_technique or ""
                    row.cells[5].text = instr.concentration or ""
                    row.cells[6].text = instr.temperature or ""
                    row.cells[7].text = instr.exposure_time or ""
                    row.cells[8].text = instr.inventory or ""
                    row.cells[9].text = instr.frequency or ""
                    row.cells[10].text = instr.executor or ""
                    row.cells[11].text = instr.control_method or ""
                else:
                    for col in range(1, 12):
                        row.cells[col].text = "_____________"

                # Принудительно применяем Arial 8pt (без bold) ко всей строке объекта
                self._apply_font_to_row(row, font_name='Arial', size_pt=8, bold=False)

                current_row += 1

        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)

        # Удаляем старый файл, если существует
        if output_file.exists():
            output_file.unlink()

        doc.save(str(output_file))
        return str(output_file)