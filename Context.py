#!/usr/bin/env python3
"""
export_project.py - Экспорт выбранных файлов проекта в один текстовый файл
С графическим интерфейсом для выбора файлов и папок
Поддерживает: .py, .txt, .json, .md, .sql, .html, .css, .js, .csv, .xlsx, .docx, .db
"""

import sys
from pathlib import Path
from datetime import datetime

# Проверка наличия tkinter
try:
    import tkinter as tk
    from tkinter import filedialog, messagebox, ttk
except ImportError:
    print("=" * 60)
    print("❌ ОШИБКА: Tkinter не установлен!")
    print("=" * 60)
    print("\nTkinter необходим для работы графического интерфейса.")
    print("\n🔧 Установка Tkinter:")
    print("\n   Windows: Tkinter устанавливается вместе с Python")
    print("            Если не работает, переустановите Python с опцией 'tcl/tk'")
    print("\n   Linux (Ubuntu/Debian):")
    print("        sudo apt-get install python3-tk")
    print("\n   Linux (Fedora/CentOS/RHEL):")
    print("        sudo dnf install python3-tkinter")
    print("\n   macOS:")
    print("        brew install python-tk")
    print("        или переустановите Python через официальный установщик")
    print("\n" + "=" * 60)
    sys.exit(1)


class ProjectExporter:
    def __init__(self, root):
        self.root = root
        self.root.title("Экспорт проекта")
        self.root.geometry("800x800")

        # Текущая директория проекта (по умолчанию - папка скрипта)
        self.project_root = Path(__file__).parent.resolve()

        # Расширения для экспорта
        self.extensions = {
            '.py', '.txt', '.json', '.md', '.sql', '.html', '.css', '.js', '.csv'
        }

        # Список выбранных файлов
        self.selected_files = []

        self.setup_ui()
        self.bind_shortcuts()

        # Заполняем дерево текущей папкой
        self.populate_tree()

    def bind_shortcuts(self):
        """Привязка клавиатурных сокращений"""
        self.root.bind('<Delete>', lambda e: self.remove_selected_from_list())
        self.root.bind('<Control-a>', lambda e: self.select_all())
        self.root.bind('<Control-A>', lambda e: self.select_all())

    def setup_ui(self):
        # Информация о проекте
        info_frame = ttk.LabelFrame(self.root, text="Информация", padding=10)
        info_frame.pack(fill=tk.X, padx=10, pady=5)

        # Строка с путём и кнопкой выбора
        path_frame = ttk.Frame(info_frame)
        path_frame.pack(fill=tk.X, pady=2)
        self.path_label = ttk.Label(path_frame, text=f"Папка проекта: {self.project_root}")
        self.path_label.pack(side=tk.LEFT, anchor=tk.W)
        ttk.Button(path_frame, text="📂 Выбрать папку проекта",
                   command=self.choose_project_folder).pack(side=tk.RIGHT, padx=5)

        ttk.Label(info_frame, text=f"Поддерживаемые расширения: {', '.join(sorted(self.extensions))}").pack(anchor=tk.W)
        ttk.Label(info_frame, text="💡 Советы: Ctrl+A - выделить всё, Delete - удалить выбранные").pack(anchor=tk.W)

        # Панель с деревом проекта
        tree_frame = ttk.LabelFrame(self.root, text="Структура проекта (выберите папки/файлы)", padding=10)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Создаём дерево с возможностью множественного выбора
        self.tree = ttk.Treeview(tree_frame, selectmode=tk.EXTENDED, show="tree")
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Скроллбар для дерева
        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.config(yscrollcommand=scrollbar.set)

        # Кнопки для работы с выбранным
        buttons_frame = ttk.Frame(self.root)
        buttons_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(buttons_frame, text="➕ Добавить выбранные папки/файлы", command=self.add_selected_from_tree).pack(
            side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="📁 Добавить все файлы из выбранных папок",
                   command=self.add_all_from_selected_dirs).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons_frame, text="🗑️ Очистить список", command=self.clear_all).pack(side=tk.LEFT, padx=5)

        # Список выбранных файлов
        selected_frame = ttk.LabelFrame(self.root, text="Выбранные файлы для экспорта", padding=10)
        selected_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Создаём список с возможностью множественного выбора
        self.files_listbox = tk.Listbox(selected_frame, selectmode=tk.EXTENDED)
        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        listbox_scrollbar = ttk.Scrollbar(selected_frame, orient=tk.VERTICAL, command=self.files_listbox.yview)
        listbox_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.files_listbox.config(yscrollcommand=listbox_scrollbar.set)

        # Кнопки для списка файлов
        list_buttons_frame = ttk.Frame(self.root)
        list_buttons_frame.pack(fill=tk.X, padx=10, pady=5)

        ttk.Button(list_buttons_frame, text="🗑️ Удалить выбранные из списка",
                   command=self.remove_selected_from_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(list_buttons_frame, text="🔍 Показать выбранные в дереве", command=self.show_selected_in_tree).pack(
            side=tk.LEFT, padx=5)

        # Кнопки экспорта и выхода
        bottom_frame = ttk.Frame(self.root)
        bottom_frame.pack(fill=tk.X, padx=10, pady=10)

        ttk.Button(bottom_frame, text="🚀 Экспортировать", command=self.export_project).pack(side=tk.RIGHT, padx=5)
        ttk.Button(bottom_frame, text="❌ Выход", command=self.root.quit).pack(side=tk.RIGHT, padx=5)

    def choose_project_folder(self):
        """Открывает диалог выбора папки и обновляет дерево"""
        folder = filedialog.askdirectory(
            title="Выберите корневую папку проекта",
            initialdir=str(self.project_root)
        )
        if folder:
            self.project_root = Path(folder).resolve()
            self.path_label.config(text=f"Папка проекта: {self.project_root}")
            # Очищаем список выбранных файлов, так как пути изменились
            self.clear_all()
            # Перестраиваем дерево
            self.populate_tree()
            messagebox.showinfo("Папка выбрана", f"Теперь корневая папка: {self.project_root}")

    def populate_tree(self, parent="", path=None):
        """Заполняет дерево проектом (рекурсивно)"""
        if path is None:
            path = self.project_root
            self.tree.delete(*self.tree.get_children())
            if not path.exists():
                messagebox.showerror("Ошибка", f"Папка {path} не существует!")
                return
            node = self.tree.insert(parent, "end", text=path.name, open=True, tags=("dir", str(path)))
            self.populate_tree(node, path)
        else:
            try:
                items = sorted([item for item in path.iterdir() if not item.name.startswith('.')])
                for item in items:
                    if item.is_dir():
                        node = self.tree.insert(parent, "end", text=item.name, tags=("dir", str(item)))
                        self.populate_tree(node, item)
                    else:
                        if item.suffix in self.extensions:
                            self.tree.insert(parent, "end", text=item.name, tags=("file", str(item)))
            except PermissionError:
                pass

    def add_selected_from_tree(self):
        """Добавляет выбранные в дереве папки/файлы в список экспорта"""
        selected = self.tree.selection()
        added = 0
        for item in selected:
            tags = self.tree.item(item, "tags")
            if tags:
                path = Path(tags[1])
                if path.is_file():
                    if path not in self.selected_files:
                        self.selected_files.append(path)
                        self.files_listbox.insert(tk.END, str(path.relative_to(self.project_root)))
                        added += 1
                elif path.is_dir():
                    # Добавляем все поддерживаемые файлы из папки
                    for ext in self.extensions:
                        for f in path.rglob(f"*{ext}"):
                            if f not in self.selected_files:
                                self.selected_files.append(f)
                                self.files_listbox.insert(tk.END, str(f.relative_to(self.project_root)))
                                added += 1
        messagebox.showinfo("Добавлено", f"Добавлено {added} файлов")

    def add_all_from_selected_dirs(self):
        """Добавляет все поддерживаемые файлы из выбранных папок"""
        selected = self.tree.selection()
        added = 0
        for item in selected:
            tags = self.tree.item(item, "tags")
            if tags:
                path = Path(tags[1])
                if path.is_dir():
                    for ext in self.extensions:
                        for f in path.rglob(f"*{ext}"):
                            if f not in self.selected_files:
                                self.selected_files.append(f)
                                self.files_listbox.insert(tk.END, str(f.relative_to(self.project_root)))
                                added += 1
        messagebox.showinfo("Добавлено", f"Добавлено {added} файлов")

    def remove_selected_from_list(self):
        """Удаляет выбранные элементы из списка экспорта"""
        selected = self.files_listbox.curselection()
        for i in reversed(selected):
            self.files_listbox.delete(i)
            del self.selected_files[i]

    def select_all(self):
        """Выделяет все элементы в списке"""
        self.files_listbox.select_set(0, tk.END)

    def show_selected_in_tree(self):
        """Подсвечивает выбранные файлы в дереве"""
        # Очищаем предыдущие подсветки
        for item in self.tree.get_children():
            self.tree.item(item, tags=self.tree.item(item, "tags"))

        # Подсвечиваем выбранные файлы
        for file_path in self.selected_files:
            self.highlight_file_in_tree(file_path)

    def highlight_file_in_tree(self, file_path):
        """Рекурсивно ищет и подсвечивает файл в дереве"""
        def find_in_tree(parent=""):
            for item in self.tree.get_children(parent):
                tags = self.tree.item(item, "tags")
                if tags and len(tags) > 1:
                    item_path = Path(tags[1])
                    if item_path == file_path:
                        self.tree.selection_add(item)
                        self.tree.see(item)
                        return True
                if self.tree.get_children(item):
                    if find_in_tree(item):
                        return True
            return False

        find_in_tree()

    def clear_all(self):
        """Очистка всего списка"""
        self.files_listbox.delete(0, tk.END)
        self.selected_files.clear()

    def get_project_structure(self, root_dir: Path, selected_paths: list, prefix: str = "",
                              is_last: bool = True) -> str:
        """Формирует дерево проекта, отмечая выбранные файлы"""
        lines = []
        if prefix == "":
            lines.append(f"{root_dir.name}/")

        items = sorted([item for item in root_dir.iterdir() if not item.name.startswith('.')])

        for i, item in enumerate(items):
            is_last_item = (i == len(items) - 1)
            connector = "└── " if is_last_item else "├── "

            # Проверяем, выбран ли файл или папка
            is_selected = False
            if item.is_file():
                is_selected = item in selected_paths
            else:
                for f in selected_paths:
                    if str(f).startswith(str(item)):
                        is_selected = True
                        break

            marker = "✅ " if is_selected else "   "

            if item.is_dir():
                lines.append(f"{prefix}{connector}{marker}{item.name}/")
                extension = "    " if is_last_item else "│   "
                lines.append(self.get_project_structure(item, selected_paths, prefix + extension, is_last_item))
            else:
                if item.suffix in self.extensions:
                    size = item.stat().st_size / 1024
                    lines.append(f"{prefix}{connector}{marker}{item.name} ({size:.1f} KB)")

        return "\n".join(lines)

    def export_project(self):
        """Экспорт выбранных файлов"""
        if not self.selected_files:
            messagebox.showwarning("Внимание", "Не выбрано ни одного файла для экспорта")
            return

        # Выбор места сохранения
        output_path = filedialog.asksaveasfilename(
            title="Сохранить экспорт как",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )

        if not output_path:
            return

        try:
            with open(output_path, 'w', encoding='utf-8') as out:
                # Заголовок
                out.write("=" * 80 + "\n")
                out.write(f"ЭКСПОРТ ПРОЕКТА: {self.project_root.name}\n")
                out.write(f"Дата экспорта: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                out.write("=" * 80 + "\n\n")

                # Список экспортируемых файлов
                out.write("📋 СПИСОК ЭКСПОРТИРУЕМЫХ ФАЙЛОВ\n")
                out.write("-" * 80 + "\n")
                for file_path in sorted(self.selected_files):
                    rel_path = file_path.relative_to(self.project_root)
                    out.write(f"   {rel_path}\n")
                out.write("\n\n")

                # Структура проекта с отметками выбранных файлов
                out.write("📁 СТРУКТУРА ПРОЕКТА (✅ - выбранные файлы)\n")
                out.write("-" * 80 + "\n")
                out.write(self.get_project_structure(self.project_root, self.selected_files))
                out.write("\n\n")

                # Содержимое выбранных файлов
                out.write("📄 СОДЕРЖИМОЕ ВЫБРАННЫХ ФАЙЛОВ\n")
                out.write("-" * 80 + "\n")

                for file_path in sorted(self.selected_files):
                    try:
                        rel_path = file_path.relative_to(self.project_root)
                        out.write(f"\n{'=' * 80}\n")
                        out.write(f"ФАЙЛ: {rel_path}\n")
                        out.write(f"{'=' * 80}\n\n")

                        with open(file_path, 'r', encoding='utf-8') as f:
                            content = f.read()
                            out.write(content)
                            if not content.endswith('\n'):
                                out.write('\n')
                    except Exception as e:
                        out.write(f"[ОШИБКА ЧТЕНИЯ: {e}]\n")

            messagebox.showinfo("Успех", f"Экспорт завершён!\nФайл: {output_path}")

        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при экспорте: {e}")


def main():
    root = tk.Tk()
    app = ProjectExporter(root)
    root.mainloop()


if __name__ == "__main__":
    main()