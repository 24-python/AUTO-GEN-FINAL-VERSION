"""
Вспомогательные функции
"""

import re
from typing import List, Optional, Tuple


def extract_text_from_cell(cell) -> str:
    """Извлекает текст из ячейки таблицы"""
    texts = []
    for paragraph in cell.paragraphs:
        if paragraph.text.strip():
            texts.append(paragraph.text.strip())
    return ' '.join(texts)


def merge_adjacent_cells(row) -> List[str]:
    """Объединяет текст из ячеек строки"""
    result = []
    for cell in row.cells:
        text = extract_text_from_cell(cell)
        if text:
            result.append(text)
    return result