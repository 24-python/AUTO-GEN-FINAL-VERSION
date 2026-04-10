"""
Модуль для генерации технологических карт
"""

from generator.docx_generator import TechCardGenerator
from generator.mapper import (
    CategoryMapper,
    find_object_and_instructions,
    get_instructions_for_object,
    get_object_by_normalized_name
)