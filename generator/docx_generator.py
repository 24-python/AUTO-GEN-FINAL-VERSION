# generator/docx_generator.py

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt
from pathlib import Path
from parser.models import ChecklistData
from db.models import (Instruction, Category as DBCategory, Object as DBObject,
                       RoomCategory, Product, ObjectProperty, ObjectGroup,
                       CleaningMethodOrder)
from db.database import SessionLocal
from collections import defaultdict
import re
import copy


class TechCardGenerator:
    TEMPLATES_DIR = Path("tech_card_templates")
    DEFAULT_TEMPLATE = TEMPLATES_DIR / "шаблон.docx"

    # Цвета для инвентаря
    INVENTORY_COLORS = {
        "черный": "969696",
        "красный": "FFCCCC",
        "желтый": "FFFFCC",
        "зеленый": "99FF99",
        "синий": "99CCFF",
        "голубой": "CCECFF",
    }

    # Порядок уровней обработки для вывода
    LEVEL_ORDER = {"основная": 1, "поддерживающая": 2, "генеральная": 3}

    # Заглушки, которые не должны подставляться из чек-листа
    SKIP_PLACEHOLDERS = {
        "Выберите элемент.", "Не выбрано", "Выберите элемент",
        "", None
    }

    def __init__(self, template_path: str = None):
        self.template_path = Path(template_path) if template_path else self.DEFAULT_TEMPLATE
        # Кэши конфигурации, загружаемые из БД
        self._config_loaded = False

    def _load_config(self, session):
        """Загружает конфигурацию из БД (CleaningMethodOrder, ObjectProperty, ObjectGroup)"""
        if self._config_loaded:
            return

        # 1. Порядок методов очистки
        methods = session.query(CleaningMethodOrder).order_by(CleaningMethodOrder.sort_order).all()
        self.cleaning_method_order = {m.method_name: m.sort_order for m in methods}

        # 2. Свойства объектов
        props = session.query(ObjectProperty).join(DBObject).all()
        self.split_objects = set()
        self.multi_method_objects = set()
        self.support_objects = set()
        self.floor_objects = set()
        self.glass_objects = set()
        self.thermal_objects = set()

        for prop in props:
            name = prop.object.normalized_name
            if prop.is_split:
                self.split_objects.add(name)
            if prop.is_multi_method:
                self.multi_method_objects.add(name)
            if prop.has_support_maintenance:
                self.support_objects.add(name)
            if prop.special_product_type == 'floor':
                self.floor_objects.add(name)
            elif prop.special_product_type == 'glass':
                self.glass_objects.add(name)
            elif prop.special_product_type == 'thermal':
                self.thermal_objects.add(name)

        # ===== ОТЛАДКА: выводим загруженные множества =====
        print(f"[DEBUG] multi_method_objects: {sorted(self.multi_method_objects)}")

        # 3. Группы объектов
        groups = session.query(ObjectGroup).all()
        self.group_headers = {}  # frozenset -> header
        temp_groups = defaultdict(set)
        for g in groups:
            if g.object:
                temp_groups[g.group_name].add(g.object.normalized_name)
        for header, obj_set in temp_groups.items():
            self.group_headers[frozenset(obj_set)] = header

        self._config_loaded = True

    # ... (все остальные методы без изменений, кроме generate) ...

    def generate(self, checklist_data: ChecklistData, output_path: str, mode: int = 1,
                 enterprise_products_path: str = None) -> str:
        doc = Document(self.template_path)
        main_table = doc.tables[1]

        # Заполняем помещение
        room_cell = main_table.cell(1, 0)
        if "Помещение:" in room_cell.text:
            room_cell.text = f"Помещение: {checklist_data.room_name or '___________'}"
            for para in room_cell.paragraphs:
                for run in para.runs:
                    run.font.name = 'Arial'
                    run.font.size = Pt(9)
                    run.font.bold = True

        # === ОБРАБОТКА 4-Й СТРОКИ (предупреждение об удалении сырья) ===
        warning_row = main_table.rows[3]
        if checklist_data.room_category not in ["производственное", "складское"]:
            self._clear_cell_text(warning_row.cells[0])

        session = SessionLocal()
        # Загрузка конфигурации из БД
        self._load_config(session)

        category_priority = self._get_category_priority(session)
        product_colors = self._load_product_colors(session)

        room_category_id = None
        if hasattr(checklist_data, 'room_category') and checklist_data.room_category:
            room_category_id = self._get_room_category_id(session, checklist_data.room_category)
            if room_category_id:
                print(f"🔍 Категория помещения: {checklist_data.room_category} (id={room_category_id})")

        if enterprise_products_path and Path(enterprise_products_path).exists():
            print(f"📦 Парсинг средств предприятия: {Path(enterprise_products_path).name}")

        checked_items = checklist_data.get_checked_items()
        all_names = [item.name for item in checked_items]

        db_objects = session.query(DBObject).filter(DBObject.normalized_name.in_(all_names)).all()
        objects_dict = {obj.normalized_name: obj for obj in db_objects}

        object_ids = [obj.id for obj in db_objects]
        db_instructions = session.query(Instruction).filter(Instruction.object_id.in_(object_ids)).all()

        instructions_dict = defaultdict(list)
        for instr in db_instructions:
            instructions_dict[instr.object_id].append(instr)

        category_object_instructions = defaultdict(list)
        all_object_instructions = []
        category_order = {}
        unmatched_objects = []

        for item in checked_items:
            obj = objects_dict.get(item.name)
            cat_name = item.category.value
            priority = category_priority.get(cat_name, 999)
            category_order[cat_name] = priority

            display_name = obj.display_name if obj else item.name
            sort_priority = obj.sort_priority if obj else 999
            normalized_name = item.name

            # ===== ОТЛАДКА: проверка для каждого объекта =====
            is_multi = normalized_name in self.multi_method_objects
            print(f"[DEBUG] Object: {normalized_name}, in multi_method_objects: {is_multi}")

            if obj and obj.id in instructions_dict:
                all_instrs = instructions_dict[obj.id]

                if normalized_name in self.split_objects:
                    instructions = self._select_split_instructions(all_instrs, room_category_id)
                    if instructions:
                        category_object_instructions[cat_name].append(('split', display_name, instructions, normalized_name))
                        all_object_instructions.append((display_name, instructions, sort_priority, normalized_name))
                    else:
                        unmatched_objects.append(display_name)
                elif normalized_name in self.support_objects:
                    by_method = defaultdict(list)
                    for instr in all_instrs:
                        method = instr.cleaning_method or ""
                        by_method[method].append(instr)
                    instructions = []
                    for method, instrs in by_method.items():
                        for maint_level in ["основная", "поддерживающая"]:
                            best = None
                            for instr in instrs:
                                if (instr.maintenance_type or "").lower() == maint_level and instr.room_category_id == room_category_id:
                                    best = instr
                                    break
                            if not best:
                                for instr in instrs:
                                    if (instr.maintenance_type or "").lower() == maint_level and instr.room_category_id is None:
                                        best = instr
                                        break
                            if best:
                                instructions.append(best)
                    instructions.sort(
                        key=lambda x: (
                            self.LEVEL_ORDER.get((x.maintenance_type or "").lower(), 99),
                            self.cleaning_method_order.get((x.cleaning_method or "").lower().strip(), 99)
                        )
                    )
                    if instructions:
                        category_object_instructions[cat_name].append(('normal', display_name, instructions, normalized_name))
                        all_object_instructions.append((display_name, instructions, sort_priority, normalized_name))
                    else:
                        unmatched_objects.append(display_name)
                elif normalized_name in self.multi_method_objects:
                    instructions = self._select_all_instructions_for_room(all_instrs, room_category_id)
                    if instructions:
                        category_object_instructions[cat_name].append(('normal', display_name, instructions, normalized_name))
                        all_object_instructions.append((display_name, instructions, sort_priority, normalized_name))
                    else:
                        unmatched_objects.append(display_name)
                else:
                    instructions = self._select_instructions_for_room(all_instrs, room_category_id)
                    if instructions:
                        category_object_instructions[cat_name].append(('normal', display_name, instructions, normalized_name))
                        all_object_instructions.append((display_name, instructions, sort_priority, normalized_name))
                    else:
                        unmatched_objects.append(display_name)
            elif obj:
                unmatched_objects.append(display_name)
            else:
                unmatched_objects.append(item.name)

        session.close()

        # ... (остальная часть generate без изменений) ...