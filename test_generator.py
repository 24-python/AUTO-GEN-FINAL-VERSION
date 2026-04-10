#!/usr/bin/env python3
"""
test_generator.py - Пакетная генерация технологических карт
Поддерживает .docx, .zip, .rar, .7z
"""

import sys
import os
import zipfile
import tempfile
import shutil
from pathlib import Path

# Добавляем корень проекта в путь
sys.path.insert(0, str(Path(__file__).parent))

from parser.xml_parser import parse_checklist
from generator.docx_generator import TechCardGenerator

# Попытка импорта для работы с архивами
try:
    import rarfile

    RAR_SUPPORT = True
except ImportError:
    RAR_SUPPORT = False

try:
    import py7zr

    SEVENZ_SUPPORT = True
except ImportError:
    SEVENZ_SUPPORT = False


def extract_archive(archive_path: Path, extract_to: Path) -> list:
    """
    Распаковывает архив и возвращает список путей к .docx файлам
    """
    docx_files = []
    ext = archive_path.suffix.lower()

    try:
        if ext == '.zip':
            with zipfile.ZipFile(archive_path, 'r') as zf:
                zf.extractall(extract_to)
        elif ext == '.rar':
            if not RAR_SUPPORT:
                print(f"   ⚠️ RAR не поддерживается. Установите: pip install rarfile")
                return []
            with rarfile.RarFile(archive_path, 'r') as rf:
                rf.extractall(extract_to)
        elif ext == '.7z':
            if not SEVENZ_SUPPORT:
                print(f"   ⚠️ 7Z не поддерживается. Установите: pip install py7zr")
                return []
            with py7zr.SevenZipFile(archive_path, 'r') as szf:
                szf.extractall(extract_to)
        else:
            return []

        # Ищем все .docx в распакованной папке
        for docx in extract_to.glob("**/*.docx"):
            if not docx.name.startswith("~"):
                docx_files.append(docx)

    except Exception as e:
        print(f"   ❌ Ошибка распаковки {archive_path.name}: {e}")

    return docx_files


def find_checklists(input_dir: Path) -> list:
    """
    Ищет все .docx файлы и архивы в папке
    """
    checklist_items = []

    # Поддерживаемые расширения
    extensions = {'.docx', '.zip', '.rar', '.7z'}

    for f in input_dir.iterdir():
        if f.is_file() and f.suffix.lower() in extensions:
            if not f.name.startswith("~") and "tech_card" not in f.name.lower():
                checklist_items.append(f)

    return checklist_items


def process_single_file(file_path: Path, generator: TechCardGenerator, output_dir: Path) -> dict:
    """
    Обрабатывает один .docx файл
    """
    result = {
        "file": file_path.name,
        "status": "error",
        "output": None,
        "error": None
    }

    try:
        # Парсинг
        checklist_data = parse_checklist(str(file_path))

        # Формируем имя выходного файла
        room_name = checklist_data.room_name or file_path.stem
        safe_name = "".join(c for c in room_name if c.isalnum() or c in (' ', '-', '_')).strip()
        output_name = f"{safe_name}_tech_card.docx"
        output_path = output_dir / output_name

        # Генерация
        generator.generate(checklist_data, str(output_path))

        result["status"] = "success"
        result["output"] = str(output_path)
        result["room_name"] = checklist_data.room_name
        result["enterprise"] = checklist_data.enterprise

    except Exception as e:
        result["error"] = str(e)
        import traceback
        traceback.print_exc()

    return result


def process_archive(archive_path: Path, generator: TechCardGenerator, output_dir: Path) -> list:
    """
    Обрабатывает архив: распаковывает и генерирует техкарты для всех .docx
    """
    results = []

    with tempfile.TemporaryDirectory() as tmpdir:
        temp_path = Path(tmpdir)

        print(f"\n📦 Распаковка: {archive_path.name}")
        docx_files = extract_archive(archive_path, temp_path)

        if not docx_files:
            results.append({
                "file": archive_path.name,
                "status": "error",
                "error": "В архиве не найдено .docx файлов"
            })
            return results

        print(f"   Найдено .docx файлов: {len(docx_files)}")

        for docx_file in docx_files:
            print(f"\n   📄 Обработка: {docx_file.name}")
            result = process_single_file(docx_file, generator, output_dir)
            results.append(result)

    return results


def print_report(results: list):
    """
    Выводит отчёт о генерации
    """
    print("\n" + "=" * 60)
    print("📊 ОТЧЁТ О ГЕНЕРАЦИИ")
    print("=" * 60)

    success_count = sum(1 for r in results if r["status"] == "success")
    error_count = len(results) - success_count

    for r in results:
        if r["status"] == "success":
            print(f"✅ {r['file']} → {r['output']}")
            if r.get("room_name"):
                print(f"   Помещение: {r['room_name']}")
        else:
            print(f"❌ {r['file']} → Ошибка: {r.get('error', 'Неизвестная ошибка')}")

    print("-" * 60)
    print(f"✅ Успешно: {success_count}")
    print(f"❌ Ошибок: {error_count}")
    print("=" * 60)


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Пакетная генерация технологических карт")
    parser.add_argument("--input", "-i", default="checklists",
                        help="Папка с чек-листами (по умолчанию: checklists)")
    parser.add_argument("--output", "-o", default="tech_cards",
                        help="Папка для сохранения техкарт (по умолчанию: tech_cards)")
    parser.add_argument("--single", "-s", help="Обработать один файл (чек-лист или архив)")
    args = parser.parse_args()

    print("🔧 ПАКЕТНАЯ ГЕНЕРАЦИЯ ТЕХНОЛОГИЧЕСКИХ КАРТ")
    print("=" * 60)

    # Проверка поддержки архивов
    if not RAR_SUPPORT:
        print("⚠️ RAR файлы не поддерживаются. Установите: pip install rarfile")
    if not SEVENZ_SUPPORT:
        print("⚠️ 7Z файлы не поддерживаются. Установите: pip install py7zr")

    # Создаём генератор
    generator = TechCardGenerator()

    # Режим одного файла
    if args.single:
        single_path = Path(args.single)
        if not single_path.exists():
            print(f"❌ Файл не найден: {single_path}")
            return

        output_dir = Path(args.output)
        output_dir.mkdir(parents=True, exist_ok=True)

        if single_path.suffix.lower() in ['.zip', '.rar', '.7z']:
            results = process_archive(single_path, generator, output_dir)
        elif single_path.suffix.lower() == '.docx':
            result = process_single_file(single_path, generator, output_dir)
            results = [result]
        else:
            print(f"❌ Неподдерживаемый формат: {single_path.suffix}")
            return

        print_report(results)
        return

    # Режим папки
    input_dir = Path(args.input)
    if not input_dir.exists():
        print(f"❌ Папка не найдена: {input_dir}")
        print(f"📌 Создайте папку 'checklists' и положите туда чек-листы")
        return

    output_dir = Path(args.output)
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"📂 Входная папка: {input_dir}")
    print(f"📂 Выходная папка: {output_dir}")

    # Поиск файлов
    files = find_checklists(input_dir)

    if not files:
        print("\n❌ Не найдено файлов (.docx, .zip, .rar, .7z)")
        print("\n📌 Что делать:")
        print("   1. Положите файлы в папку 'checklists'")
        print("   2. Или укажите другую папку: --input <путь>")
        return

    print(f"\n📋 Найдено файлов: {len(files)}")
    for f in files:
        size_kb = f.stat().st_size / 1024
        print(f"   - {f.name} ({size_kb:.1f} KB)")

    print("\n" + "-" * 40)
    print("🚀 НАЧАЛО ОБРАБОТКИ")
    print("-" * 40)

    all_results = []

    for file_path in files:
        if file_path.suffix.lower() in ['.zip', '.rar', '.7z']:
            results = process_archive(file_path, generator, output_dir)
            all_results.extend(results)
        elif file_path.suffix.lower() == '.docx':
            result = process_single_file(file_path, generator, output_dir)
            all_results.append(result)

    print_report(all_results)


if __name__ == "__main__":
    main()