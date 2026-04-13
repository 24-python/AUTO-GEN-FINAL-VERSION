#!/usr/bin/env python3
"""
test_generator.py - Пакетная генерация технологических карт
Поддерживает .docx и .zip
"""

import sys
import zipfile
import tempfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from parser.xml_parser import parse_checklist
from generator.docx_generator import TechCardGenerator


def extract_zip(archive_path: Path, extract_to: Path) -> list:
    """Распаковывает ZIP и возвращает список .docx файлов"""
    docx_files = []
    try:
        with zipfile.ZipFile(archive_path, 'r') as zf:
            zf.extractall(extract_to)
        for docx in extract_to.glob("**/*.docx"):
            if not docx.name.startswith("~"):
                docx_files.append(docx)
    except Exception as e:
        print(f"   ❌ Ошибка распаковки {archive_path.name}: {e}")
    return docx_files


def find_files(input_dir: Path) -> list:
    """Находит все .docx и .zip файлы в папке"""
    files = []
    for f in input_dir.iterdir():
        if f.is_file() and f.suffix.lower() in {'.docx', '.zip'}:
            if not f.name.startswith("~") and "tech_card" not in f.name.lower():
                files.append(f)
    return files


def process_docx(file_path: Path, generator: TechCardGenerator, output_dir: Path) -> dict:
    """Обрабатывает один .docx файл"""
    result = {"file": file_path.name, "status": "error", "output": None, "error": None}
    try:
        checklist_data = parse_checklist(str(file_path))
        room_name = checklist_data.room_name or file_path.stem
        safe_name = "".join(c for c in room_name if c.isalnum() or c in (' ', '-', '_')).strip()
        output_path = output_dir / f"{safe_name}_tech_card.docx"
        generator.generate(checklist_data, str(output_path))
        result["status"] = "success"
        result["output"] = str(output_path)
    except Exception as e:
        result["error"] = str(e)
    return result


def process_zip(zip_path: Path, generator: TechCardGenerator, output_dir: Path) -> list:
    """Обрабатывает ZIP архив"""
    results = []
    with tempfile.TemporaryDirectory() as tmpdir:
        temp_path = Path(tmpdir)
        docx_files = extract_zip(zip_path, temp_path)
        if not docx_files:
            results.append({"file": zip_path.name, "status": "error", "error": "В архиве не найдено .docx файлов"})
            return results
        for docx_file in docx_files:
            results.append(process_docx(docx_file, generator, output_dir))
    return results


def print_report(results: list):
    success = sum(1 for r in results if r["status"] == "success")
    error = len(results) - success
    print("\n" + "=" * 60)
    print("📊 ОТЧЁТ О ГЕНЕРАЦИИ")
    print("=" * 60)
    for r in results:
        if r["status"] == "success":
            print(f"✅ {r['file']} → {r['output']}")
        else:
            print(f"❌ {r['file']} → {r.get('error', 'Ошибка')}")
    print("-" * 60)
    print(f"✅ Успешно: {success}")
    print(f"❌ Ошибок: {error}")
    print("=" * 60)


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Пакетная генерация технологических карт")
    parser.add_argument("--input", "-i", default="checklists", help="Папка с чек-листами")
    parser.add_argument("--output", "-o", default="tech_cards", help="Папка для сохранения")
    args = parser.parse_args()

    input_dir = Path(args.input)
    if not input_dir.exists():
        print(f"❌ Папка не найдена: {input_dir}")
        return

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    print("🔧 ПАКЕТНАЯ ГЕНЕРАЦИЯ ТЕХНОЛОГИЧЕСКИХ КАРТ")
    print("=" * 60)
    print(f"📂 Входная папка: {input_dir}")
    print(f"📂 Выходная папка: {output_dir}")

    files = find_files(input_dir)
    if not files:
        print("\n❌ Не найдено файлов (.docx, .zip)")
        return

    print(f"\n📋 Найдено файлов: {len(files)}")
    for f in files:
        size = f.stat().st_size / 1024
        print(f"   - {f.name} ({size:.1f} KB)")

    generator = TechCardGenerator()
    all_results = []

    for file_path in files:
        if file_path.suffix.lower() == '.docx':
            print(f"\n📄 Обработка: {file_path.name}")
            all_results.append(process_docx(file_path, generator, output_dir))
        elif file_path.suffix.lower() == '.zip':
            print(f"\n📦 Обработка ZIP: {file_path.name}")
            all_results.extend(process_zip(file_path, generator, output_dir))

    print_report(all_results)


if __name__ == "__main__":
    main()