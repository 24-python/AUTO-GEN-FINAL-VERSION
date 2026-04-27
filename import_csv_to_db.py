#!/usr/bin/env python3
"""
import_csv_to_db.py

Импорт данных из CSV в базу данных с выбором режима работы.
Подготовлен для интеграции в веб-версию.

Поддерживает:
- Импорт категорий помещений (room_categories)
- Импорт инструкций с привязкой к категории помещения (room_category_id)
"""

import csv
import sys
from pathlib import Path
from datetime import datetime
from typing import Dict, List, Tuple, Optional, Callable

sys.path.insert(0, str(Path(__file__).parent))

from db.database import SessionLocal, engine
from db.models import Base, Category, Object, Instruction, RoomCategory


class ImportStats:
    """Статистика импорта"""

    def __init__(self):
        self.categories_created = 0
        self.categories_updated = 0
        self.room_categories_created = 0
        self.room_categories_updated = 0
        self.objects_created = 0
        self.objects_updated = 0
        self.objects_skipped = 0
        self.instructions_created = 0
        self.instructions_updated = 0
        self.instructions_skipped = 0
        self.errors = []
        self.start_time = datetime.now()
        self.end_time = None

    def finish(self):
        self.end_time = datetime.now()

    def duration_seconds(self) -> float:
        if self.end_time:
            return (self.end_time - self.start_time).total_seconds()
        return 0

    def to_dict(self) -> dict:
        return {
            'categories_created': self.categories_created,
            'categories_updated': self.categories_updated,
            'room_categories_created': self.room_categories_created,
            'room_categories_updated': self.room_categories_updated,
            'objects_created': self.objects_created,
            'objects_updated': self.objects_updated,
            'objects_skipped': self.objects_skipped,
            'instructions_created': self.instructions_created,
            'instructions_updated': self.instructions_updated,
            'instructions_skipped': self.instructions_skipped,
            'errors': len(self.errors),
            'error_details': self.errors[:10],
            'duration_seconds': self.duration_seconds(),
        }

    def print_report(self):
        print("\n" + "=" * 80)
        print("📊 СТАТИСТИКА ИМПОРТА")
        print("=" * 80)
        print(f"   Категорий объектов создано:    {self.categories_created}")
        print(f"   Категорий объектов обновлено:  {self.categories_updated}")
        print(f"   Категорий помещений создано:   {self.room_categories_created}")
        print(f"   Категорий помещений обновлено: {self.room_categories_updated}")
        print(f"   Объектов создано:              {self.objects_created}")
        print(f"   Объектов обновлено:            {self.objects_updated}")
        print(f"   Объектов пропущено:            {self.objects_skipped}")
        print(f"   Инструкций создано:            {self.instructions_created}")
        print(f"   Инструкций обновлено:          {self.instructions_updated}")
        print(f"   Инструкций пропущено:          {self.instructions_skipped}")
        print(f"   Ошибок:                        {len(self.errors)}")
        print(f"   Время выполнения:              {self.duration_seconds():.2f} сек")
        print("=" * 80)

        if self.errors:
            print("\n⚠️ ОШИБКИ:")
            for err in self.errors[:10]:
                print(f"   - {err}")
            if len(self.errors) > 10:
                print(f"   ... и ещё {len(self.errors) - 10}")


class CSVImporter:
    """Импортер CSV в БД"""

    # Режимы импорта
    MODE_ADD_NEW = 1      # Добавление новых, пропуск дубликатов
    MODE_CLEAR_DATA = 2   # Очистка данных перед импортом
    MODE_RECREATE_DB = 3  # Полное пересоздание БД
    MODE_EXIT = 4         # Выход без импорта

    MODE_NAMES = {
        MODE_ADD_NEW: "Добавление новых (пропуск дубликатов)",
        MODE_CLEAR_DATA: "Очистка данных перед импортом",
        MODE_RECREATE_DB: "Полное пересоздание БД",
        MODE_EXIT: "Выход без импорта",
    }

    def __init__(self, csv_path: str, mode: int = MODE_ADD_NEW, progress_callback: Callable = None):
        """
        Инициализация импортера.

        Args:
            csv_path: Путь к CSV файлу
            mode: Режим импорта (1, 2, 3, 4)
            progress_callback: Функция для обновления прогресса (для веб-версии)
        """
        self.csv_path = Path(csv_path)
        self.mode = mode
        self.progress_callback = progress_callback
        self.stats = ImportStats()
        self.session = None
        self.room_category_cache = {}  # Кэш {name: id} для категорий помещений

    def _log_progress(self, message: str, percent: int = None):
        """Логирование прогресса"""
        print(message)
        if self.progress_callback:
            self.progress_callback(message, percent)

    def _recreate_database(self):
        """Полное пересоздание БД"""
        self._log_progress("🔄 Удаление старых таблиц...", 5)
        Base.metadata.drop_all(bind=engine)

        self._log_progress("🔄 Создание новых таблиц...", 10)
        Base.metadata.create_all(bind=engine)

        self._log_progress("✅ БД пересоздана", 15)

    def _clear_data(self):
        """Очистка данных из таблиц"""
        self._log_progress("🔄 Очистка инструкций...", 5)
        self.session.query(Instruction).delete()

        self._log_progress("🔄 Очистка объектов...", 10)
        self.session.query(Object).delete()

        self._log_progress("🔄 Очистка категорий помещений...", 12)
        self.session.query(RoomCategory).delete()

        self._log_progress("🔄 Очистка категорий объектов...", 15)
        self.session.query(Category).delete()

        self.session.commit()
        self._log_progress("✅ Данные очищены", 20)

    def _read_csv(self) -> List[Dict]:
        """Чтение CSV файла"""
        self._log_progress("📖 Чтение CSV файла...", 25)

        rows = []
        with open(self.csv_path, 'r', encoding='utf-8-sig') as f:
            reader = csv.DictReader(f, delimiter=';')
            for row in reader:
                rows.append(row)

        self._log_progress(f"✅ Прочитано {len(rows)} строк", 30)
        return rows

    def _import_categories(self, rows: List[Dict]) -> Dict[str, int]:
        """Импорт категорий объектов"""
        self._log_progress("📁 Импорт категорий объектов...", 35)

        # Собираем уникальные категории
        unique_categories = {}
        for row in rows:
            cat_name = row.get('category_name', '').strip()
            if cat_name:
                unique_categories[cat_name] = unique_categories.get(cat_name, 0) + 1

        category_ids = {}

        for cat_name in unique_categories:
            existing = self.session.query(Category).filter(Category.name == cat_name).first()

            if existing:
                category_ids[cat_name] = existing.id
                self.stats.categories_updated += 1
            else:
                cat = Category(name=cat_name, sort_order=len(category_ids) + 1)
                self.session.add(cat)
                self.session.flush()
                category_ids[cat_name] = cat.id
                self.stats.categories_created += 1

        self.session.commit()
        self._log_progress(f"✅ Категорий объектов: {len(category_ids)} (новых: {self.stats.categories_created})", 40)
        return category_ids

    def _import_room_categories(self, rows: List[Dict]) -> Dict[str, int]:
        """Импорт категорий помещений из колонки room_category_name"""
        self._log_progress("📁 Импорт категорий помещений...", 42)

        unique_room_cats = set()
        for row in rows:
            rc_name = row.get('room_category_name', '').strip()
            if rc_name:
                unique_room_cats.add(rc_name)

        # Добавляем стандартные, если их нет
        standard_cats = ["Производственное", "Складское", "Инженерное", "Вспомогательное"]
        for sc in standard_cats:
            unique_room_cats.add(sc)

        room_cat_ids = {}

        for rc_name in sorted(unique_room_cats):
            existing = self.session.query(RoomCategory).filter(RoomCategory.name == rc_name).first()

            if existing:
                room_cat_ids[rc_name] = existing.id
                self.stats.room_categories_updated += 1
            else:
                rc = RoomCategory(name=rc_name)
                self.session.add(rc)
                self.session.flush()
                room_cat_ids[rc_name] = rc.id
                self.stats.room_categories_created += 1

        self.session.commit()
        self._log_progress(
            f"✅ Категорий помещений: {len(room_cat_ids)} (новых: {self.stats.room_categories_created})", 45
        )
        return room_cat_ids

    def _get_room_category_id(self, room_cat_name: str, room_cat_ids: Dict[str, int]) -> Optional[int]:
        """Получает ID категории помещения по названию"""
        if not room_cat_name:
            return None

        # Проверяем кэш
        if room_cat_name in self.room_category_cache:
            return self.room_category_cache[room_cat_name]

        # Ищем в переданном словаре
        if room_cat_name in room_cat_ids:
            rc_id = room_cat_ids[room_cat_name]
            self.room_category_cache[room_cat_name] = rc_id
            return rc_id

        # Ищем в БД
        rc = self.session.query(RoomCategory).filter(RoomCategory.name == room_cat_name).first()
        if rc:
            self.room_category_cache[room_cat_name] = rc.id
            return rc.id

        return None

    def _import_objects_and_instructions(self, rows: List[Dict], category_ids: Dict[str, int],
                                          room_cat_ids: Dict[str, int]):
        """Импорт объектов и инструкций"""
        total_rows = len(rows)

        # Группируем строки по object_id (если есть) или по normalized_name
        object_groups = {}
        for row in rows:
            obj_id = row.get('object_id', '').strip()
            norm_name = row.get('normalized_name', '').strip()

            if not norm_name:
                continue

            key = obj_id if obj_id else norm_name
            if key not in object_groups:
                object_groups[key] = {
                    'object_data': row,
                    'instructions': []
                }

            # Добавляем инструкцию, если есть данные
            if row.get('cleaning_method', '').strip() or row.get('product_name', '').strip():
                object_groups[key]['instructions'].append(row)

        processed = 0
        total_groups = len(object_groups)

        for key, group in object_groups.items():
            processed += 1

            if processed % 50 == 0:
                percent = 45 + int((processed / total_groups) * 45)
                self._log_progress(f"📦 Обработка объектов: {processed}/{total_groups}", percent)

            row = group['object_data']
            norm_name = row.get('normalized_name', '').strip()
            display_name = row.get('display_name', '').strip()
            base_name = row.get('base_name', '').strip()
            modifier = row.get('modifier', '').strip() or None
            cat_name = row.get('category_name', '').strip()
            sort_priority = int(row.get('sort_priority', 0) or 0)

            if not norm_name or not display_name:
                continue

            category_id = category_ids.get(cat_name)
            if not category_id:
                self.stats.errors.append(f"Категория не найдена: {cat_name} для {norm_name}")
                continue

            # Ищем существующий объект
            existing = self.session.query(Object).filter(Object.normalized_name == norm_name).first()

            if existing:
                if self.mode == self.MODE_ADD_NEW:
                    obj = existing
                    self.stats.objects_skipped += 1
                else:
                    existing.display_name = display_name
                    existing.base_name = base_name
                    existing.modifier = modifier
                    existing.category_id = category_id
                    existing.sort_priority = sort_priority
                    obj = existing
                    self.stats.objects_updated += 1
            else:
                obj = Object(
                    normalized_name=norm_name,
                    display_name=display_name,
                    base_name=base_name or display_name,
                    modifier=modifier,
                    category_id=category_id,
                    sort_priority=sort_priority
                )
                self.session.add(obj)
                self.session.flush()
                self.stats.objects_created += 1

            # Импорт инструкций
            for instr_row in group['instructions']:
                cleaning_method = instr_row.get('cleaning_method', '').strip()
                product_name = instr_row.get('product_name', '').strip()

                if not cleaning_method and not product_name:
                    continue

                # Получаем room_category_id из CSV
                room_cat_name = instr_row.get('room_category_name', '').strip()
                room_category_id = self._get_room_category_id(room_cat_name, room_cat_ids)

                # Ищем существующую инструкцию (с учётом room_category_id)
                existing_instr = self.session.query(Instruction).filter(
                    Instruction.object_id == obj.id,
                    Instruction.cleaning_method == cleaning_method,
                    Instruction.product_name == product_name,
                    Instruction.room_category_id == room_category_id
                ).first()

                if existing_instr:
                    if self.mode == self.MODE_ADD_NEW:
                        self.stats.instructions_skipped += 1
                    else:
                        existing_instr.room_category_id = room_category_id
                        existing_instr.instruction_number = instr_row.get('instruction_number', '').strip() or None
                        existing_instr.cleaning_technique = instr_row.get('cleaning_technique', '').strip() or None
                        existing_instr.concentration = instr_row.get('concentration', '').strip() or None
                        existing_instr.temperature = instr_row.get('temperature', '').strip() or None
                        existing_instr.exposure_time = instr_row.get('exposure_time', '').strip() or None
                        existing_instr.inventory = instr_row.get('inventory', '').strip() or None
                        existing_instr.frequency = instr_row.get('frequency', '').strip() or None
                        existing_instr.executor = instr_row.get('executor', '').strip() or None
                        existing_instr.control_method = instr_row.get('control_method', '').strip() or None
                        self.stats.instructions_updated += 1
                else:
                    instr = Instruction(
                        object_id=obj.id,
                        room_category_id=room_category_id,
                        cleaning_method=cleaning_method or None,
                        instruction_number=instr_row.get('instruction_number', '').strip() or None,
                        product_name=product_name or None,
                        cleaning_technique=instr_row.get('cleaning_technique', '').strip() or None,
                        concentration=instr_row.get('concentration', '').strip() or None,
                        temperature=instr_row.get('temperature', '').strip() or None,
                        exposure_time=instr_row.get('exposure_time', '').strip() or None,
                        inventory=instr_row.get('inventory', '').strip() or None,
                        frequency=instr_row.get('frequency', '').strip() or None,
                        executor=instr_row.get('executor', '').strip() or None,
                        control_method=instr_row.get('control_method', '').strip() or None
                    )
                    self.session.add(instr)
                    self.stats.instructions_created += 1

            if processed % 100 == 0:
                self.session.commit()

        self.session.commit()
        self._log_progress(
            f"✅ Объектов: создано {self.stats.objects_created}, обновлено {self.stats.objects_updated}", 90
        )
        self._log_progress(
            f"✅ Инструкций: создано {self.stats.instructions_created}, обновлено {self.stats.instructions_updated}", 95
        )

    def run(self) -> ImportStats:
        """Запуск импорта"""

        if not self.csv_path.exists():
            raise FileNotFoundError(f"CSV файл не найден: {self.csv_path}")

        self._log_progress("🚀 Начало импорта...", 0)
        self._log_progress(f"📂 Файл: {self.csv_path}", 0)
        self._log_progress(f"🔧 Режим: {self.MODE_NAMES.get(self.mode, 'Неизвестно')}", 0)

        try:
            self.session = SessionLocal()

            if self.mode == self.MODE_EXIT:
                self._log_progress("👋 Выход без импорта", 100)
                self.stats.finish()
                return self.stats

            if self.mode == self.MODE_RECREATE_DB:
                self._recreate_database()
            elif self.mode == self.MODE_CLEAR_DATA:
                self._clear_data()

            rows = self._read_csv()

            if not rows:
                self._log_progress("⚠️ CSV файл пуст", 100)
                self.stats.finish()
                return self.stats

            # Импорт категорий
            category_ids = self._import_categories(rows)
            room_cat_ids = self._import_room_categories(rows)

            # Импорт объектов и инструкций
            self._import_objects_and_instructions(rows, category_ids, room_cat_ids)

            self._log_progress("🎉 Импорт завершён!", 100)

        except Exception as e:
            self.session.rollback()
            self.stats.errors.append(str(e))
            self._log_progress(f"❌ Ошибка: {e}", 100)
            raise
        finally:
            self.session.close()
            self.stats.finish()

        return self.stats


def select_mode() -> int:
    """Интерактивный выбор режима"""
    print("\n" + "=" * 60)
    print("📥 ИМПОРТ CSV В БАЗУ ДАННЫХ")
    print("=" * 60)
    print("\nВыберите режим импорта:")
    print(f"  1. {CSVImporter.MODE_NAMES[CSVImporter.MODE_ADD_NEW]}")
    print(f"  2. {CSVImporter.MODE_NAMES[CSVImporter.MODE_CLEAR_DATA]}")
    print(f"  3. {CSVImporter.MODE_NAMES[CSVImporter.MODE_RECREATE_DB]}")
    print(f"  4. {CSVImporter.MODE_NAMES[CSVImporter.MODE_EXIT]}")

    while True:
        try:
            choice = input("\n🔢 Ваш выбор (1-4): ").strip()
            mode = int(choice)
            if mode in [1, 2, 3, 4]:
                return mode
            print("❌ Введите число от 1 до 4")
        except ValueError:
            print("❌ Введите корректное число")
        except KeyboardInterrupt:
            print("\n👋 Отмена")
            return CSVImporter.MODE_EXIT


def main():
    import argparse

    parser = argparse.ArgumentParser(description="Импорт CSV в базу данных")
    parser.add_argument("file", nargs="?", help="Путь к CSV файлу")
    parser.add_argument("--mode", "-m", type=int, choices=[1, 2, 3, 4],
                        help="Режим импорта (1-добавление, 2-очистка, 3-пересоздание, 4-выход)")
    parser.add_argument("--yes", "-y", action="store_true", help="Пропустить подтверждение")
    args = parser.parse_args()

    if args.file:
        csv_path = args.file
    else:
        csv_files = list(Path(".").glob("*.csv"))
        if not csv_files:
            print("❌ CSV файлы не найдены в текущей папке")
            return
        elif len(csv_files) == 1:
            csv_path = str(csv_files[0])
            print(f"📁 Найден файл: {csv_path}")
        else:
            print("\n📁 Найдено несколько CSV файлов:")
            for i, f in enumerate(csv_files, 1):
                print(f"   {i}. {f.name}")
            choice = input("\n🔢 Выберите номер файла: ").strip()
            try:
                csv_path = str(csv_files[int(choice) - 1])
            except (ValueError, IndexError):
                print("❌ Неверный выбор")
                return

    if not Path(csv_path).exists():
        print(f"❌ Файл не найден: {csv_path}")
        return

    if args.mode:
        mode = args.mode
    else:
        mode = select_mode()

    if mode == CSVImporter.MODE_EXIT:
        print("👋 Выход")
        return

    if mode in [CSVImporter.MODE_CLEAR_DATA, CSVImporter.MODE_RECREATE_DB] and not args.yes:
        print(f"\n⚠️ ВНИМАНИЕ! Выбран режим: {CSVImporter.MODE_NAMES[mode]}")
        print("   Все существующие данные будут УДАЛЕНЫ!")
        confirm = input("\n   Продолжить? (yes/no): ").strip().lower()
        if confirm not in ['yes', 'y', 'да', 'д']:
            print("👋 Отмена")
            return

    try:
        importer = CSVImporter(csv_path, mode)
        stats = importer.run()
        stats.print_report()
    except Exception as e:
        print(f"\n❌ Ошибка импорта: {e}")
        import traceback
        traceback.print_exc()


# ============================================================
# ФУНКЦИИ ДЛЯ ИНТЕГРАЦИИ В ВЕБ-ВЕРСИЮ
# ============================================================

def import_csv_web(csv_path: str, mode: int, progress_callback: Callable = None) -> dict:
    """
    Функция для вызова из веб-приложения.

    Args:
        csv_path: Путь к CSV файлу
        mode: Режим импорта (1, 2, 3)
        progress_callback: Функция(message, percent) для обновления прогресса

    Returns:
        dict: Статистика импорта
    """
    importer = CSVImporter(csv_path, mode, progress_callback)
    stats = importer.run()
    return stats.to_dict()


if __name__ == "__main__":
    main()