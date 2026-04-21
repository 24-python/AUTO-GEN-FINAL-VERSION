"""
Простой генератор техкарт на чистом python-docx (без lxml).
Медленнее, но гарантированно создаёт рабочий файл.
"""

from docx import Document
from pathlib import Path
from parser.models import ChecklistData
from db.models import Instruction, Category as DBCategory, Object as DBObject
from db.database import SessionLocal
from collections import defaultdict
import time


class SimpleFastGenerator:
    TEMPLATES_DIR = Path("tech_card_templates")
    DEFAULT_TEMPLATE = TEMPLATES_DIR / "шаблон.docx"

    def __init__(self, template_path: str = None):
        self.template_path = Path(template_path) if template_path else self.DEFAULT_TEMPLATE

    def _get_category_priority(self, session) -> dict:
        categories = session.query(DBCategory).order_by(DBCategory.sort_order).all()
        return {cat.name: cat.sort_order for cat in categories}

    def generate(self, checklist_data: ChecklistData, output_path: str) -> str:
        print("      📝 Генерация через python-docx (простая версия)...")
        total_start = time.time()

        # 1. Загрузка шаблона
        doc = Document(self.template_path)
        main_table = doc.tables[1]

        # 2. Помещение
        room_cell = main_table.cell(1, 0)
        if "Помещение:" in room_cell.text:
            room_cell.text = f"Помещение: {checklist_data.room_name or '______________'}"

        # 3. Загрузка данных из БД
        session = SessionLocal()
        category_priority = self._get_category_priority(session)

        checked_items = checklist_data.get_checked_items()
        all_names = [item.name for item in checked_items]

        db_objects = session.query(DBObject).filter(DBObject.normalized_name.in_(all_names)).all()
        objects_dict = {obj.normalized_name: obj for obj in db_objects}

        object_ids = [obj.id for obj in db_objects]
        db_instructions = session.query(Instruction).filter(Instruction.object_id.in_(object_ids)).all()

        instructions_dict = defaultdict(list)
        for instr in db_instructions:
            instructions_dict[instr.object_id].append(instr)

        # 4. Группировка
        category_items = defaultdict(list)
        category_order = {}

        for item in checked_items:
            obj = objects_dict.get(item.name)
            cat_name = item.category.value
            priority = category_priority.get(cat_name, 999)
            category_order[cat_name] = priority

            display_name = obj.display_name if obj else item.name

            if obj and obj.id in instructions_dict:
                for instr in instructions_dict[obj.id]:
                    category_items[cat_name].append((display_name, instr))
            else:
                category_items[cat_name].append((display_name, None))

        session.close()

        # 5. Сортировка
        sorted_categories = sorted(category_items.keys(), key=lambda x: category_order.get(x, 999))
        for cat_name in sorted_categories:
            category_items[cat_name].sort(key=lambda x: x[0])

        # 6. Очистка таблицы (удаляем все строки после заголовков)
        start_row = 7  # после строки с заголовками колонок
        while len(main_table.rows) > start_row:
            main_table._tbl.remove(main_table.rows[start_row]._tr)

        # 7. Заполнение
        for category_name in sorted_categories:
            items = category_items[category_name]
            if not items:
                continue

            # Строка категории
            row = main_table.add_row()
            row.cells[0].merge(row.cells[11])
            row.cells[0].text = category_name

            # Объекты
            for obj_name, instr in items:
                row = main_table.add_row()
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

        # 8. Сохранение
        output_file = Path(output_path)
        output_file.parent.mkdir(parents=True, exist_ok=True)
        if output_file.exists():
            output_file.unlink()
        doc.save(str(output_file))

        size_kb = output_file.stat().st_size / 1024
        print(f"      📁 Размер файла: {size_kb:.1f} KB")
        print(f"      ✅ ГОТОВО! Время: {time.time() - total_start:.2f}с")
        return str(output_file)