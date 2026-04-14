# debug_why_in_other.py
from parser.xml_parser import parse_checklist
from db.database import SessionLocal
from db.models import Object
import re


def normalize_name(name: str) -> str:
    """Текущая функция нормализации из парсера"""
    normalized = re.sub(r'[^\w\s\-\(\)]', '', name)
    normalized = normalized.lower()
    normalized = re.sub(r'\s+', ' ', normalized)
    return normalized.strip()


# Парсим чек-лист
data = parse_checklist("тарамоечная.docx")

print("\n🔍 ОБЪЕКТЫ В 'ПРОЧЕЕ'")
print("=" * 50)

session = SessionLocal()

for item in data.get_checked_items():
    if item.category.value == "Прочее":
        normalized = normalize_name(item.name)
        print(f"\n📌 {item.name}")
        print(f"   нормализованное: {normalized}")

        # Ищем в БД
        obj = session.query(Object).filter(Object.normalized_name == normalized).first()
        if obj:
            print(f"   ✅ НАЙДЕН в БД: {obj.display_name}")
        else:
            print(f"   ❌ НЕ НАЙДЕН в БД")
            # Поиск похожих
            similar = session.query(Object).filter(Object.normalized_name.like(f"%{normalized[:5]}%")).all()
            if similar:
                print(f"   📋 Похожие в БД:")
                for s in similar:
                    print(f"      - {s.normalized_name} → {s.display_name}")

session.close()