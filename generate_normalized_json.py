#!/usr/bin/env python3
"""
ГЕНЕРАТОР НОРМАЛИЗОВАННОГО JSON
"""

import json
import re
from pathlib import Path
from openpyxl import load_workbook
from collections import defaultdict

MATCH_RULES = {
    "потолок (высота _____м) П": ["потолок подвесной", "потолок (подвесной)"],
    "потолок (высота _____м) ОК": ["потолок окрашенный", "потолок (окрашенный)"],
    "светильники": ["светильники (наружные поверхности)"],
    "вытяжные зонты Н": ["вытяжные зонты"],
    "вытяжные зонты А": ["вытяжные зонты"],
    "вентиляционные трубы": ["вентиляционные трубы (наружные поверхности)"],
    "диффузоры": ["вентиляционные диффузоры"],
    "кондиционер": ["кондиционер (наружные поверхности)"],
    "сплит-системы": ["кондиционер (наружные поверхности)"],
    "инсектицидные лампы": ["инсектицидные лампы"],
    "стены (выше 2 м)": ["стены выше 2м (краска)", "стены выше 2м"],
    "стены (до 2 м)": ["стены до 2м (плитка)", "стены до 2м (окрашенные)", "стены на всю высоту"],
    "бактерицидные лампы": ["бактерицидные рециркуляторы"],
    "кабель-каналы": ["кабель-каналы"],
    "водопроводные трубы": ["водопроводные трубы"],
    "выключатели": ["выключатели"],
    "розетки": ["розетки"],
    "дверные ручки": ["дверные ручки"],
    "окна внешние": ["окна (внешние поверхности: стеклопакеты, рамы, откосы, отлив)"],
    "окна внутрицеховые": ["окна (внутренние поверхности: стеклопакеты, внутренние рамы, ручки, подоконники, откосы)"],
    "пожарный щит": ["пожарные щиты (наружные поверхности)"],
    "отопительные приборы": ["отопительные приборы и пространство за ними"],
    "бойлера": ["бойлер (наружные поверхности)", "бойлера (наружные surfaces)"],
    "двери": ["двери"],
    "завесы ПВХ": ["завесы ПВХ полосовые"],
    "пол": ["пол"],
    "мусорные корзины": ["мусорные корзины"],
    "контейнеры для отходов": ["контейнеры для сбора отходов"],
    "трапы": ["канализационные трапы"],
    "стоки": ["стоки"],
    "пластиковые паллеты": ["пластиковые паллеты"],
    "пластиковые подкаты": ["пластиковые подкаты", "подкаты"],
    "держатели для ножей": ["держатели для ножей"],
    "держатели для инвентаря": ["держатели для уборочного инвентаря"],
    "раковина": ["сантехническое оборудование (раковины, унитазы, поддон душевой)",
                 "санитарный пост (раковина, смеситель)", "технологическая мойка (раковина, смеситель)",
                 "технологическая раковина (раковина, смеситель)"],
    "смеситель": ["санитарный пост (раковина, смеситель)", "технологическая мойка (раковина, смеситель)",
                  "технологическая раковина (раковина, смеситель)"],
    "унитаз": ["сантехническое оборудование (раковины, унитазы, поддон душевой)"],
    "душевая кабина": ["душевые кабины (наружные и внутренние поверхности)"],
    "душевые секции": ["сантехническое оборудование (раковины, унитазы, поддон душевой)"],
    "технологическая мойка": ["технологическая мойка (раковина, смеситель)",
                              "технологическая раковина (раковина, смеситель)"],
    "диспенсер для полотенец": ["диспенсеры для полотенец"],
    "диспенсер для т/б": ["держатели для туалетной бумаги"],
    "дозаторы для мыла/антисептика": ["дозаторы для мыла/антисептика"],
    "столы": ["столы (все поверхности)", "столы (столешница – контактная поверхность)", "стеллажи, столы"],
    "стулья": ["стулья"],
    "полки": ["полки", "полки навесные"],
    "стеллажи": ["стеллажи, столы", "стеллажи"],
    "шкафы": ["шкафчики"],
    "вешалки": ["вешалки"],
    "скамейки": ["скамейки"],
    "зеркала": ["зеркальные поверхности"],
    "чайники": ["чайник"],
    "микроволновки": ["микроволновая печь"],
    "холодильники": ["холодильник"],
    "стиральные машины": ["стиральная машина"],
    "сушильные машины": ["сушильная машина"],
    "доски": ["доски пластиковые"],
    "посуда": ["посуда"],
    "инвентарь": ["инвентарь"],
    "ножи": ["ножи"],
    "гастроёмкости": ["гастроёмкости"],
    "листы для выпечки Н": ["листы от шпилек из нержавеющей стали"],
    "листы для выпечки А": ["листы алюминиевые"],
    "съёмные детали оборудования Н": ["съёмные деталей оборудования", "съёмные детали оборудования"],
    "съёмные детали оборудования А": ["съёмные деталей оборудования", "съёмные детали оборудования"],
    "внутрицеховая тара (вёдра, ящики)": ["ящики"],
    "передвижные ёмкости": ["ёмкости на колёсах"],
    "шпильки": ["шпильки"],
    "листы от шпилек": ["листы от шпилек из нержавеющей стали"],
    "тележки подкатные": ["тележки подкатные"],
    "мопы": ["многоразовый протирочный материал (мопы)"],
    "ведра": ["крупногабаритный уборочный инвентарь (тележки, вёдра, АВД, пылесос)"],
    "АВД": ["крупногабаритный уборочный инвентарь (тележки, вёдра, АВД, пылесос)"],
    "пылесосы": ["крупногабаритный уборочный инвентарь (тележки, вёдра, АВД, пылесос)"],
    "поломоечная машина": ["крупногабаритный уборочный инвентарь (тележки, вёдра, АВД, пылесос)"],
    "холод. шкафы": ["Внешние поверхности холодильное оборудование (холодильные шкафы)",
                     "Внутренние поверхности холодильное оборудование (холодильные шкафы)"],
    "плиты индук.": ["плиты"],
    "плиты элек.": ["плиты"],
    "плиты газ.": ["плиты"],
    "производственные столы Н": ["рабочие поверхности (столы металлические)", "рабочие поверхности (столы)"],
    "производственные столы Д": ["рабочие поверхности (столы металлические)", "рабочие поверхности (столы)"],
    "весы напольные": ["весы напольные"],
    "весы настольные": ["весы настольные"],
    "просеиватели мука": ["просеиватели"],
    "просеиватели сахар": ["просеиватели"],
    "вакуумные упаковщики": ["упаковочное оборудование (вакуумный упаковщик)"],
}


def parse_objects_from_text(text_path):
    objects = []
    current_category = None

    with open(text_path, 'r', encoding='utf-8') as f:
        lines = f.readlines()

    for line in lines:
        line = line.strip()
        if not line:
            continue

        if line == "просеиватели: мука сахар":
            continue

        if ':' in line:
            current_category = line.replace(':', '').strip()
            continue

        if not current_category:
            continue

        line = re.sub(r'\s+', ' ', line).strip()

        name = line
        base_name = line
        modifier = None

        if line.endswith(' П'):
            name = line
            base_name = line[:-2].strip()
            modifier = 'П'
        elif line.endswith(' ОК'):
            name = line
            base_name = line[:-3].strip()
            modifier = 'ОК'
        elif line.endswith(' Н'):
            name = line
            base_name = line[:-2].strip()
            modifier = 'Н'
        elif line.endswith(' А'):
            name = line
            base_name = line[:-2].strip()
            modifier = 'А'
        elif ' (выше 2 м)' in line:
            name = line
            base_name = 'стены'
            modifier = 'выше 2 м'
        elif ' (до 2 м)' in line:
            name = line
            base_name = 'стены'
            modifier = 'до 2 м'
        elif ' (внешние)' in line:
            name = line
            base_name = 'окна'
            modifier = 'внешние'
        elif ' (внутрицеховые)' in line:
            name = line
            base_name = 'окна'
            modifier = 'внутрицеховые'
        elif ' (контактная поверхность)' in line:
            name = line
            base_name = 'демосистема'
            modifier = 'контактная поверхность'

        if 'потолок' in line and ('П' in line or 'ОК' in line):
            name = line
            base_name = 'потолок'
            modifier = 'П' if 'П' in line else 'ОК'

        if 'вытяжные зонты' in line and ('Н' in line or 'А' in line):
            name = line
            base_name = 'вытяжные зонты'
            modifier = 'Н' if 'Н' in line else 'А'

        objects.append({
            "name": name,
            "base_name": base_name,
            "modifier": modifier,
            "category": current_category,
            "sort_priority": 0
        })

    return objects


def parse_excel_instructions(excel_path):
    instructions = []

    wb = load_workbook(excel_path, data_only=True)
    ws = wb.active

    current_category = None

    for row in ws.iter_rows(min_row=1, values_only=True):
        if not any(row):
            continue

        object_name = str(row[0]).strip() if row[0] else ""

        if object_name.endswith(':'):
            current_category = object_name.replace(':', '').strip()
            continue

        if not object_name or object_name == "Объект обработки" or not current_category:
            continue

        instruction = {
            "object_name": object_name,
            "category": current_category,
            "cleaning_method": str(row[1]).strip() if len(row) > 1 and row[1] else "",
            "instruction_number": str(row[2]).strip() if len(row) > 2 and row[2] else "",
            "product_name": str(row[3]).strip() if len(row) > 3 and row[3] else "",
            "cleaning_technique": str(row[4]).strip() if len(row) > 4 and row[4] else "",
            "concentration": str(row[5]).strip() if len(row) > 5 and row[5] else "",
            "temperature": str(row[6]).strip() if len(row) > 6 and row[6] else "",
            "exposure_time": str(row[7]).strip() if len(row) > 7 and row[7] else "",
            "inventory": str(row[8]).strip() if len(row) > 8 and row[8] else "",
            "frequency": str(row[9]).strip() if len(row) > 9 and row[9] else "",
            "executor": str(row[10]).strip() if len(row) > 10 and row[10] else "",
            "control_method": str(row[11]).strip() if len(row) > 11 and row[11] else ""
        }

        instructions.append(instruction)

    return instructions


def get_instruction_key(instr):
    """Уникальный ключ для сравнения (включает cleaning_method, чтобы не потерять разные методы)"""
    return (
        instr["cleaning_method"],
        instr["product_name"],
        instr["concentration"],
        instr["temperature"],
        instr["exposure_time"],
        instr["frequency"]
    )


def is_empty_instruction(instr):
    fields = ["cleaning_method", "product_name", "concentration", "temperature", "exposure_time", "frequency",
              "executor", "control_method"]
    return not any(instr.get(f, "").strip() for f in fields)


def match_instructions_to_objects(objects, excel_instructions):
    excel_by_category = defaultdict(list)
    for instr in excel_instructions:
        excel_by_category[instr["category"]].append(instr)

    all_instructions = []

    for obj in objects:
        obj_name = obj["name"]
        obj_category = obj["category"]

        found = False
        rules = MATCH_RULES.get(obj_name, [])

        if rules:
            for rule in rules:
                for instr in excel_by_category.get(obj_category, []):
                    if rule in instr["object_name"]:
                        new_instr = instr.copy()
                        new_instr["object_name"] = obj_name
                        all_instructions.append(new_instr)
                        found = True

        if obj_name == "тележки":
            for cat in ["Инвентарь, посуда и т.д.", "Моечный, уборочный инвентарь и оборудование"]:
                for instr in excel_by_category.get(cat, []):
                    if "тележки" in instr["object_name"].lower():
                        new_instr = instr.copy()
                        new_instr["object_name"] = obj_name
                        new_instr["category"] = obj_category
                        all_instructions.append(new_instr)
                        found = True

        if not found:
            all_instructions.append({
                "object_name": obj_name,
                "category": obj_category,
                "cleaning_method": "",
                "instruction_number": "",
                "product_name": "",
                "cleaning_technique": "",
                "concentration": "",
                "temperature": "",
                "exposure_time": "",
                "inventory": "",
                "frequency": "",
                "executor": "",
                "control_method": ""
            })

    # Удаляем только полные дубликаты (все поля совпадают)
    unique = []
    seen = set()
    duplicate_count = 0

    for instr in all_instructions:
        key = get_instruction_key(instr)
        if key not in seen:
            seen.add(key)
            unique.append(instr)
        else:
            duplicate_count += 1

    print(f"   Удалено полных дубликатов: {duplicate_count}")

    # Группируем по объектам
    grouped = defaultdict(list)
    for instr in unique:
        grouped[instr["object_name"]].append(instr)

    # Для каждого объекта оставляем все непустые и одну пустую (если нет непустых)
    final = []
    empty_removed = 0

    for obj_name, instructions in grouped.items():
        non_empty = [i for i in instructions if not is_empty_instruction(i)]
        empty = [i for i in instructions if is_empty_instruction(i)]

        if non_empty:
            final.extend(non_empty)
            empty_removed += len(empty)
        else:
            if empty:
                final.append(empty[0])
                empty_removed += len(empty) - 1

    print(f"   Удалено пустых дублей: {empty_removed}")

    return final, duplicate_count + empty_removed


def main():
    print("🔧 ГЕНЕРАТОР НОРМАЛИЗОВАННОГО JSON")
    print("=" * 60)

    text_file = Path("полный перечень всех поверхностей из чек-листа.txt")
    excel_file = Path("Объекты с определенной обработкой.xlsx")
    output_file = Path("normalized_data.json")

    if not text_file.exists():
        print(f"❌ Нет файла: {text_file}")
        return
    if not excel_file.exists():
        print(f"❌ Нет файла: {excel_file}")
        return

    print("\n📖 Чтение текстового файла...")
    objects = parse_objects_from_text(text_file)
    print(f"   ✅ Найдено объектов: {len(objects)}")

    print("\n📖 Чтение Excel...")
    excel_instructions = parse_excel_instructions(excel_file)
    print(f"   ✅ Найдено инструкций в Excel: {len(excel_instructions)}")

    print("\n🔄 Сопоставление и очистка...")
    final_instructions, removed = match_instructions_to_objects(objects, excel_instructions)

    categories = sorted(set(obj["category"] for obj in objects))

    non_empty_count = 0
    for obj in objects:
        if any(i["object_name"] == obj["name"] and not is_empty_instruction(i) for i in final_instructions):
            non_empty_count += 1

    result = {
        "statistics": {
            "total_objects": len(objects),
            "objects_with_instructions": non_empty_count,
            "objects_without_instructions": len(objects) - non_empty_count,
            "total_instructions": len(final_instructions),
            "removed_duplicates": removed
        },
        "categories": [{"name": cat, "sort_order": i + 1} for i, cat in enumerate(categories)],
        "objects": objects,
        "instructions": final_instructions
    }

    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(result, f, ensure_ascii=False, indent=2)

    print(f"\n✅ JSON сохранен: {output_file}")
    print(f"\n📊 СТАТИСТИКА:")
    print(f"   - Всего объектов: {result['statistics']['total_objects']}")
    print(f"   - Объектов с инструкциями: {result['statistics']['objects_with_instructions']}")
    print(f"   - Объектов без инструкций: {result['statistics']['objects_without_instructions']}")
    print(f"   - Всего инструкций: {result['statistics']['total_instructions']}")
    print(f"   - Удалено дубликатов: {result['statistics']['removed_duplicates']}")


if __name__ == "__main__":
    main()