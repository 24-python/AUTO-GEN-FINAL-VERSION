# debug_parser_cameras.py
import re

def normalize_name(name: str) -> str:
    normalized = re.sub(r'[^\w\s\-\(\)]', '', name)
    normalized = normalized.lower()
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized.strip()

# То, что парсер может выдать из чек-листа
test_cases = [
    "камеры холд.",
    "камеры холд",
    "камеры холод.",
    "камеры холод",
    "камеры мороз.",
    "камеры мороз",
    "камеры шок. замор.",
    "камеры шок замор",
]

print("🔍 ЧТО ПАРСЕР ОТПРАВИТ В БД")
print("=" * 40)

for test in test_cases:
    result = normalize_name(test)
    print(f"  '{test}' → '{result}'")