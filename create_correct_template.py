from docx import Document
from docx.shared import Cm, Pt, Inches
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.enum.section import WD_ORIENTATION
from pathlib import Path


def set_cell_font(cell, text, font_name='Arial', font_size=Pt(9), bold=False, alignment=None):
    """Универсальная установка текста и шрифта в ячейке"""
    cell.paragraphs[0].clear()
    run = cell.paragraphs[0].add_run(text)
    run.font.name = font_name
    run.font.size = font_size
    run.font.bold = bold
    if alignment:
        cell.paragraphs[0].alignment = alignment


def create_template():
    doc = Document()

    # Альбомная ориентация
    section = doc.sections[0]
    section.orientation = WD_ORIENTATION.LANDSCAPE
    section.page_width = Inches(11.69)
    section.page_height = Inches(8.27)
    section.top_margin = Cm(1.5)
    section.bottom_margin = Cm(1.5)
    section.left_margin = Cm(1.5)
    section.right_margin = Cm(1.5)

    # Таблица: 6 строк, 12 колонок
    table = doc.add_table(rows=6, cols=12)
    table.style = 'Table Grid'

    # === ОБЪЕДИНЕНИЕ ПЕРВЫХ 5 СТРОК ===
    for row_idx in range(5):
        start_cell = table.cell(row_idx, 0)
        for col_idx in range(1, 12):
            start_cell.merge(table.cell(row_idx, col_idx))

    # Заполнение объединенных строк (Arial Bold 9pt, левое выравнивание)
    texts = [
        "Технологическая карта санитарной обработки № ________",
        "Помещение: _________________________",
        "",
        "Санитарная обработка помещений осуществляется после полного удаления из них сырьевых компонентов и материалов.",
        ""
    ]
    for row_idx, text in enumerate(texts):
        set_cell_font(table.cell(row_idx, 0), text, bold=True, alignment=WD_PARAGRAPH_ALIGNMENT.LEFT)

    # === СТРОКА 5: ЗАГОЛОВКИ КОЛОНОК (Arial Bold 9pt) ===
    headers = [
        "Объект обработки", "Способ обработки", "№ инструкции",
        "Наименование средства", "Метод уборки",
        "Концентрация средства, %, метод разведения",
        "Температура раствора, °C", "Время выдержки средства, мин",
        "Инвентарь", "Периодичность", "Исполнитель",
        "Метод контроля/периодичность"
    ]
    for col, header in enumerate(headers):
        set_cell_font(table.cell(5, col), header, bold=True)

    # Ширина колонок
    widths = [Cm(3.0), Cm(2.0), Cm(1.5), Cm(2.5), Cm(2.0), Cm(2.5),
              Cm(1.8), Cm(1.8), Cm(1.8), Cm(2.0), Cm(2.0), Cm(2.5)]
    for col, width in enumerate(widths):
        table.columns[col].width = width

    # Добавляем пустую строку для данных (Arial Bold 9pt)
    row = table.add_row()
    for col in range(12):
        set_cell_font(row.cells[col], "")

    # Сохраняем
    output_path = Path("tech_card_templates/шаблон.docx")
    output_path.parent.mkdir(exist_ok=True)
    doc.save(str(output_path))
    print(f"✅ Шаблон создан: {output_path}")
    print("   - Альбомная ориентация")
    print("   - Первые 5 строк объединены, Arial Bold 9pt, выравнивание по левому краю")
    print("   - Заголовки: Arial Bold 9pt")
    print("   - Строки данных: Arial Bold 9pt")


if __name__ == "__main__":
    import os

    old = Path("tech_card_templates/шаблон.docx")
    if old.exists():
        os.remove(old)
    create_template()