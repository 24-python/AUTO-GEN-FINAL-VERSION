#!/usr/bin/env python3
"""
test_generator_new.py - Тестирование БЫСТРОГО генератора технологических карт
Поддерживает .docx и .zip
"""

import sys
import zipfile
import tempfile
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from parser.xml_parser import parse_checklist
from Разное.docx_generator_new import FastTechCardGenerator


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


def process_docx(file_path: Path, generator: FastTechCardGenerator, output_dir: Path) -> dict:
    """Обрабатывает один .docx файл"""
    result = {"file": file_path.name, "status": "error", "output": None, "error": None}
    try:
        print(f"   🔍 Парсинг чек-листа...")
        parse_start = time.time()
        checklist_data = parse_checklist(str(file_path))
        parse_time = time.time() - parse_start

        checked_count = len(checklist_data.get_checked_items())
        print(f"   ✅ Найдено отмеченных объектов: {checked_count} (за {parse_time:.2f} сек)")

        room_name = checklist_data.room_name or file_path.stem
        safe_name = "".join(c for c in room_name if c.isalnum() or c in (' ', '-', '_')).strip()
        output_path = output_dir / f"{safe_name}_tech_card.docx"

        print(f"   ⚡ Генерация техкарты (БЫСТРЫЙ генератор)...")
        gen_start = time.time()
        generator.generate(checklist_data, str(output_path))
        gen_time = time.time() - gen_start

        result["status"] = "success"
        result["output"] = str(output_path)
        result["objects"] = checked_count
        result["parse_time"] = parse_time
        result["gen_time"] = gen_time
        result["total_time"] = parse_time + gen_time

        print(
            f"   ✅ Готово! (парсинг: {parse_time:.2f}с, генерация: {gen_time:.2f}с, всего: {parse_time + gen_time:.2f}с)")
    except Exception as e:
        result["error"] = str(e)
        import traceback
        traceback.print_exc()
    return result


def process_zip(zip_path: Path, generator: FastTechCardGenerator, output_dir: Path) -> list:
    """Обрабатывает ZIP архив"""
    results = []
    with tempfile.TemporaryDirectory() as tmpdir:
        temp_path = Path(tmpdir)
        docx_files = extract_zip(zip_path, temp_path)
        if not docx_files:
            results.append({
                "file": zip_path.name,
                "status": "error",
                "error": "В архиве не найдено .docx файлов"
            })
            return results
        for docx_file in docx_files:
            results.append(process_docx(docx_file, generator, output_dir))
    return results


def print_report(results: list):
    """Выводит отчёт о генерации"""
    success = sum(1 for r in results if r["status"] == "success")
    error = len(results) - success

    total_objects = sum(r.get("objects", 0) for r in results if r["status"] == "success")
    total_time = sum(r.get("total_time", 0) for r in results if r["status"] == "success")

    print("\n" + "=" * 70)
    print("📊 ОТЧЁТ О ГЕНЕРАЦИИ (БЫСТРЫЙ ГЕНЕРАТОР)")
    print("=" * 70)

    for r in results:
        if r["status"] == "success":
            print(f"✅ {r['file']} ({r.get('objects', 0)} об.)")
            print(f"   → {r['output']}")
            print(
                f"   ⏱️ Парсинг: {r.get('parse_time', 0):.2f}с | Генерация: {r.get('gen_time', 0):.2f}с | Всего: {r.get('total_time', 0):.2f}с")
        else:
            print(f"❌ {r['file']} → {r.get('error', 'Ошибка')}")

    print("-" * 70)
    print(f"✅ Успешно обработано файлов: {success}")
    print(f"❌ Ошибок: {error}")
    print(f"📋 Всего объектов: {total_objects}")
    print(f"⏱️ Общее время: {total_time:.2f} сек")
    print("=" * 70)


def process_single_file(file_path: str, output_dir: str = "tech_cards_new"):
    """Обрабатывает один файл (для тестирования)"""
    input_path = Path(file_path)
    if not input_path.exists():
        print(f"❌ Файл не найден: {file_path}")
        return None

    output_path = Path(output_dir)
    output_path.mkdir(parents=True, exist_ok=True)

    generator = FastTechCardGenerator()

    if input_path.suffix.lower() == '.docx':
        return process_docx(input_path, generator, output_path)
    elif input_path.suffix.lower() == '.zip':
        results = process_zip(input_path, generator, output_path)
        return results[0] if results else None
    else:
        print(f"❌ Неподдерживаемый формат: {input_path.suffix}")
        return None


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Тестирование БЫСТРОГО генератора технологических карт")
    parser.add_argument("--input", "-i", default="checklists", help="Папка с чек-листами")
    parser.add_argument("--output", "-o", default="tech_cards_new", help="Папка для сохранения")
    parser.add_argument("--single", "-s", help="Обработать один файл (для тестирования)")
    args = parser.parse_args()

    print("⚡ БЫСТРЫЙ ГЕНЕРАТОР ТЕХНОЛОГИЧЕСКИХ КАРТ")
    print("=" * 60)

    # Режим обработки одного файла
    if args.single:
        print("🔧 ТЕСТОВАЯ ГЕНЕРАЦИЯ (ОДИН ФАЙЛ)")
        print("=" * 60)
        result = process_single_file(args.single, args.output)
        if result:
            print_report([result])
        return

    # Режим пакетной обработки
    input_dir = Path(args.input)
    if not input_dir.exists():
        print(f"❌ Папка не найдена: {input_dir}")
        return

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

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

    generator = FastTechCardGenerator()
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