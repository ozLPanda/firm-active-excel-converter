from __future__ import annotations

import threading
from dataclasses import replace
from pathlib import Path
from tkinter import BooleanVar, Button, Canvas, Checkbutton, Entry, Frame, Label, Menu, Radiobutton, Scrollbar, StringVar, TclError, Tk, Toplevel, filedialog, messagebox

from converter import (
    GroupConfig,
    SheetIssue,
    SourceProduct,
    convert,
    default_aliases_path,
    default_group_cache_path,
    default_schema_path,
    detect_source_groups,
    money,
    valid_product_code,
)


APP_TITLE = "Satu.kz Excel Converter"


class ConverterApp:
    def __init__(self) -> None:
        self.root = Tk()
        self.root.title(APP_TITLE)
        self.root.geometry("760x325")
        self.root.minsize(680, 325)

        self.source_path = StringVar()
        self.output_path = StringVar()
        self.status = StringVar(value="Выберите исходный прайс для конвертации.")
        self.progress_text = StringVar(value="")
        self.progress_value = 0.0
        self.image_progress_text = StringVar(value="")
        self.image_progress_value = 0.0
        self.allow_multiple_images = BooleanVar(value=False)
        self.apply_without_code_to_all: bool | None = None

        self._bind_global_entry_shortcuts()
        self._build_ui()
        self._center_window(self.root)

    def _bind_global_entry_shortcuts(self) -> None:
        self.root.bind_class("Entry", "<Control-KeyPress>", self._handle_entry_control_key)
        self.root.bind_class("Entry", "<Shift-Insert>", self._paste_into_entry)
        self.root.bind_class("Entry", "<Button-3>", self._show_entry_context_menu)

    def _center_window(self, window) -> None:
        window.update_idletasks()
        width = window.winfo_width()
        height = window.winfo_height()
        if width <= 1 or height <= 1:
            geometry = window.geometry().split("+", 1)[0]
            if "x" in geometry:
                try:
                    width_text, height_text = geometry.split("x", 1)
                    width = int(width_text)
                    height = int(height_text)
                except ValueError:
                    width = max(window.winfo_reqwidth(), 1)
                    height = max(window.winfo_reqheight(), 1)
            else:
                width = max(window.winfo_reqwidth(), 1)
                height = max(window.winfo_reqheight(), 1)
        x = max(0, (window.winfo_screenwidth() - width) // 2)
        y = max(0, (window.winfo_screenheight() - height) // 2)
        window.geometry(f"{width}x{height}+{x}+{y}")

    def _bind_entry_shortcuts(self, entry: Entry) -> None:
        entry.bind("<Control-KeyPress>", self._handle_entry_control_key)
        entry.bind("<Shift-Insert>", self._paste_into_entry)
        entry.bind("<Button-3>", self._show_entry_context_menu)

    def _handle_entry_control_key(self, event) -> str | None:
        # Windows keycodes make shortcuts work even when the active keyboard layout is Russian.
        keycode = int(getattr(event, "keycode", 0) or 0)
        if keycode == 86:  # V
            return self._paste_into_entry(event)
        if keycode == 67:  # C
            return self._copy_from_entry(event)
        if keycode == 88:  # X
            return self._cut_from_entry(event)
        if keycode == 65:  # A
            entry = event.widget
            entry.select_range(0, "end")
            entry.icursor("end")
            return "break"
        return None

    def _paste_into_entry(self, event) -> str:
        entry = event.widget
        try:
            text = self.root.clipboard_get()
        except TclError:
            return "break"
        try:
            if entry.selection_present():
                entry.delete("sel.first", "sel.last")
        except TclError:
            pass
        entry.insert("insert", text)
        return "break"

    def _copy_from_entry(self, event) -> str:
        entry = event.widget
        try:
            text = entry.selection_get()
        except TclError:
            return "break"
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        return "break"

    def _cut_from_entry(self, event) -> str:
        entry = event.widget
        try:
            text = entry.selection_get()
        except TclError:
            return "break"
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        entry.delete("sel.first", "sel.last")
        return "break"

    def _show_entry_context_menu(self, event) -> str:
        entry = event.widget
        entry.focus_set()
        menu = Menu(entry, tearoff=0)
        menu.add_command(label="Вырезать", command=lambda: self._cut_from_entry_widget(entry))
        menu.add_command(label="Копировать", command=lambda: self._copy_from_entry_widget(entry))
        menu.add_command(label="Вставить", command=lambda: self._paste_into_entry_widget(entry))
        menu.add_separator()
        menu.add_command(label="Выделить всё", command=lambda: self._select_all_entry_text(entry))
        menu.tk_popup(event.x_root, event.y_root)
        return "break"

    def _paste_into_entry_widget(self, entry: Entry) -> None:
        try:
            text = self.root.clipboard_get()
        except TclError:
            return
        try:
            if entry.selection_present():
                entry.delete("sel.first", "sel.last")
        except TclError:
            pass
        entry.insert("insert", text)

    def _copy_from_entry_widget(self, entry: Entry) -> None:
        try:
            text = entry.selection_get()
        except TclError:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)

    def _cut_from_entry_widget(self, entry: Entry) -> None:
        try:
            text = entry.selection_get()
        except TclError:
            return
        self.root.clipboard_clear()
        self.root.clipboard_append(text)
        entry.delete("sel.first", "sel.last")

    def _select_all_entry_text(self, entry: Entry) -> None:
        entry.select_range(0, "end")
        entry.icursor("end")

    def _build_ui(self) -> None:
        container = Frame(self.root, padx=16, pady=14)
        container.pack(fill="both", expand=True)
        container.columnconfigure(1, weight=1)

        Label(container, text="Исходный прайс").grid(row=0, column=0, sticky="w", pady=6)
        Entry(container, textvariable=self.source_path).grid(row=0, column=1, sticky="ew", padx=8, pady=6)
        Button(container, text="Выбрать...", command=self.pick_source).grid(row=0, column=2, pady=6)

        Label(container, text="Финальный файл").grid(row=1, column=0, sticky="w", pady=6)
        Entry(container, textvariable=self.output_path).grid(row=1, column=1, sticky="ew", padx=8, pady=6)
        Button(container, text="Куда сохранить...", command=self.pick_output).grid(row=1, column=2, pady=6)

        Checkbutton(
            container,
            text="Разрешить несколько разных фото для одного артикула (_2, _3)",
            variable=self.allow_multiple_images,
            anchor="w",
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(10, 2))

        self.convert_button = Button(container, text="Конвертировать", command=self.start_convert, height=2)
        self.convert_button.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(10, 8))

        Label(container, textvariable=self.status, anchor="w").grid(row=4, column=0, columnspan=3, sticky="ew")
        self.progress_canvas = Canvas(container, height=14, highlightthickness=1, highlightbackground="#b8b8b8", bg="#f3f3f3")
        self.progress_fill = self.progress_canvas.create_rectangle(0, 0, 0, 14, fill="#4a90e2", width=0)
        self.progress_canvas.bind("<Configure>", lambda _event: self._draw_progress())
        self.progress_canvas.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(8, 2))
        Label(container, textvariable=self.progress_text, anchor="w").grid(row=6, column=0, columnspan=3, sticky="ew")
        self.image_progress_canvas = Canvas(container, height=14, highlightthickness=1, highlightbackground="#b8b8b8", bg="#f3f3f3")
        self.image_progress_fill = self.image_progress_canvas.create_rectangle(0, 0, 0, 14, fill="#2e9d57", width=0)
        self.image_progress_canvas.bind("<Configure>", lambda _event: self._draw_image_progress())
        self.image_progress_canvas.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(8, 2))
        Label(container, textvariable=self.image_progress_text, anchor="w").grid(row=8, column=0, columnspan=3, sticky="ew")

    def pick_source(self) -> None:
        path = filedialog.askopenfilename(
            title="Выберите исходный прайс",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.source_path.set(path)
            self._fill_default_output()

    def pick_output(self) -> None:
        initial_dir = Path(self.source_path.get()).parent if self.source_path.get() else Path.cwd()
        path = filedialog.asksaveasfilename(
            title="Куда сохранить результат",
            initialdir=initial_dir,
            initialfile="satu-products-converted.xlsx",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if path:
            self.output_path.set(path)

    def _fill_default_output(self) -> None:
        if self.output_path.get():
            return
        source = self.source_path.get()
        if not source:
            return
        output = Path(source).with_name("satu-products-converted.xlsx")
        self.output_path.set(str(output))

    def start_convert(self) -> None:
        if not self.source_path.get().strip():
            messagebox.showerror(APP_TITLE, "Выберите исходный прайс.")
            return
        if not self.output_path.get().strip():
            messagebox.showerror(APP_TITLE, "Укажите путь для финального файла.")
            return

        source = Path(self.source_path.get())
        output = Path(self.output_path.get())
        schema = default_schema_path()

        if not source.exists():
            messagebox.showerror(APP_TITLE, "Выберите существующий файл исходного прайса.")
            return
        if not schema.exists():
            messagebox.showerror(APP_TITLE, f"Не найден файл структуры satu.kz:\n\n{schema}")
            return
        if output.exists() and output.is_dir():
            messagebox.showerror(APP_TITLE, "Финальный путь должен быть файлом .xlsx, а не папкой.")
            return

        try:
            self.status.set("Анализ групп...")
            group_configs = detect_source_groups(source, default_aliases_path(), default_schema_path(), default_group_cache_path())
        except Exception as exc:
            self._show_error(exc)
            return

        edited_group_configs = self.ask_edit_groups(group_configs)
        if edited_group_configs is None:
            self.status.set("Конвертация отменена.")
            return

        self.apply_without_code_to_all = None
        self.convert_button.config(state="disabled")
        self._reset_progress()
        self.status.set("Конвертация выполняется...")
        allow_multiple_images = self.allow_multiple_images.get()
        thread = threading.Thread(target=self._convert_in_background, args=(source, output, edited_group_configs, allow_multiple_images), daemon=True)
        thread.start()

    def _convert_in_background(
        self,
        source: Path,
        output: Path,
        group_configs: list[GroupConfig],
        allow_multiple_images: bool,
    ) -> None:
        try:
            stats = convert(
                source,
                output,
                default_aliases_path(),
                default_schema_path(),
                self.confirm_sheet_issue,
                group_configs,
                default_group_cache_path(),
                self.confirm_duplicate_product,
                self.ask_edit_missing_prices,
                self._on_progress,
                self._on_image_progress,
                allow_multiple_images_per_product=allow_multiple_images,
            )
        except Exception as exc:
            self.root.after(0, self._show_error, exc)
            return
        self.root.after(
            0,
            self._show_success,
            output,
            stats.converted_count,
            stats.missing_group_ids,
            stats.skipped_without_code,
            stats.skipped_without_price,
            stats.skipped_invalid_code,
            stats.skipped_without_group_id,
            stats.exported_images_count,
            stats.images_output_dir,
        )

    def _reset_progress(self) -> None:
        self.progress_value = 0
        self.image_progress_value = 0
        self._draw_progress()
        self._draw_image_progress()
        self.progress_text.set("")
        self.image_progress_text.set("")

    def _set_progress(self, processed: int, total: int) -> None:
        if total <= 0:
            self.progress_value = 0
            self._draw_progress()
            self.progress_text.set("Товаров для обработки: 0")
            return
        self.progress_value = max(0, min(1, processed / total))
        self._draw_progress()
        self.progress_text.set(f"Обработано товаров: {processed} из {total}")

    def _on_progress(self, processed: int, total: int) -> None:
        self.root.after(0, self._set_progress, processed, total)

    def _set_image_progress(self, processed: int, total: int) -> None:
        if total < 0:
            total_sheets = abs(total)
            self.image_progress_value = 0 if total_sheets == 0 else max(0, min(1, processed / total_sheets))
            self._draw_image_progress()
            self.image_progress_text.set(f"Анализ листов с изображениями: {processed} из {total_sheets}")
            return
        if total <= 0:
            self.image_progress_value = 0
            self._draw_image_progress()
            self.image_progress_text.set("Изображений для обработки: 0")
            return
        self.image_progress_value = max(0, min(1, processed / total))
        self._draw_image_progress()
        self.image_progress_text.set(f"Обработано изображений: {processed} из {total}")

    def _on_image_progress(self, processed: int, total: int) -> None:
        self.root.after(0, self._set_image_progress, processed, total)

    def _draw_progress(self) -> None:
        width = max(0, self.progress_canvas.winfo_width())
        height = max(1, self.progress_canvas.winfo_height())
        self.progress_canvas.coords(self.progress_fill, 0, 0, width * self.progress_value, height)

    def _draw_image_progress(self) -> None:
        width = max(0, self.image_progress_canvas.winfo_width())
        height = max(1, self.image_progress_canvas.winfo_height())
        self.image_progress_canvas.coords(self.image_progress_fill, 0, 0, width * self.image_progress_value, height)

    def confirm_sheet_issue(self, issue: SheetIssue) -> bool:
        if issue.issue_type != "without_code":
            return False
        if self.apply_without_code_to_all is not None:
            return self.apply_without_code_to_all

        event = threading.Event()
        result: dict[str, bool] = {}

        def ask() -> None:
            include, apply_to_all = self.ask_include_without_code(issue)
            result["include"] = include
            if apply_to_all:
                self.apply_without_code_to_all = include
            event.set()

        self.root.after(0, ask)
        event.wait()
        return result["include"]

    def confirm_duplicate_product(self, code: str, products: list[SourceProduct]) -> SourceProduct:
        event = threading.Event()
        result: dict[str, SourceProduct] = {}

        def ask() -> None:
            result["product"] = self.ask_duplicate_product(code, products)
            event.set()

        self.root.after(0, ask)
        event.wait()
        return result["product"]

    def ask_duplicate_product(self, code: str, products: list[SourceProduct]) -> SourceProduct:
        dialog = Toplevel(self.root)
        dialog.title("Повтор кода товара")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("760x420")
        dialog.minsize(640, 320)

        selected = StringVar(value="0")
        frame = Frame(dialog, padx=16, pady=14)
        frame.pack(fill="both", expand=True)
        Label(
            frame,
            text=f"Код товара {code} найден в нескольких строках с разными названиями. Выберите запись, которую оставить.",
            justify="left",
            wraplength=700,
        ).pack(anchor="w", pady=(0, 10))

        list_frame = Frame(frame)
        list_frame.pack(fill="both", expand=True)
        for index, product in enumerate(products):
            text = f"{product.name}\nЛист: {product.source_sheet} | Цена: {product.retail_price}"
            Radiobutton(list_frame, text=text, variable=selected, value=str(index), justify="left", anchor="w").pack(fill="x", anchor="w", pady=4)

        buttons = Frame(frame)
        buttons.pack(anchor="e", pady=(10, 0))

        def choose() -> None:
            dialog.destroy()

        Button(buttons, text="Оставить выбранную", width=20, command=choose).pack(side="right")
        dialog.protocol("WM_DELETE_WINDOW", choose)
        self._center_window(dialog)
        self.root.wait_window(dialog)
        return products[int(selected.get())]

    def ask_edit_missing_prices(self, products: list[SourceProduct]) -> list[SourceProduct]:
        event = threading.Event()
        result: dict[str, list[SourceProduct]] = {"products": []}

        def ask() -> None:
            result["products"] = self.ask_edit_missing_prices_dialog(products)
            event.set()

        self.root.after(0, ask)
        event.wait()
        return result["products"]

    def ask_edit_missing_prices_dialog(self, products: list[SourceProduct]) -> list[SourceProduct]:
        dialog = Toplevel(self.root)
        dialog.title("Проверка кода и цены")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("980x560")
        dialog.minsize(760, 380)

        rows: list[tuple[SourceProduct, StringVar, StringVar]] = []
        result: dict[str, list[SourceProduct]] = {"products": []}

        container = Frame(dialog, padx=14, pady=12)
        container.pack(fill="both", expand=True)
        Label(
            container,
            text="Проверьте товары с пустой/нулевой ценой или невалидным кодом. В Excel попадут только строки с корректным кодом и ценой больше 0.",
            anchor="w",
            justify="left",
            wraplength=930,
        ).pack(fill="x", pady=(0, 10))

        header = Frame(container)
        header.pack(fill="x")
        Label(header, text="Лист", width=24, anchor="w").grid(row=0, column=0, sticky="ew", padx=(0, 8))
        Label(header, text="Код", width=18, anchor="w").grid(row=0, column=1, sticky="ew", padx=(0, 8))
        Label(header, text="Название", width=50, anchor="w").grid(row=0, column=2, sticky="ew", padx=(0, 8))
        Label(header, text="Цена", width=14, anchor="w").grid(row=0, column=3, sticky="ew")

        canvas = Canvas(container, highlightthickness=1, highlightbackground="#cfcfcf")
        scrollbar = Scrollbar(container, orient="vertical", command=canvas.yview)
        table = Frame(canvas)
        table.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=table, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, pady=(4, 10))
        scrollbar.pack(side="right", fill="y", pady=(4, 10))

        def scroll_table(event) -> str:
            if getattr(event, "num", None) == 4:
                canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(1, "units")
            else:
                delta = int(getattr(event, "delta", 0) or 0)
                if delta:
                    canvas.yview_scroll(int(-1 * (delta / 120)), "units")
            return "break"

        for widget in (dialog, container, canvas, table):
            widget.bind("<MouseWheel>", scroll_table)
            widget.bind("<Button-4>", scroll_table)
            widget.bind("<Button-5>", scroll_table)

        for row_index, product in enumerate(products):
            code_var = StringVar(value=product.code or "")
            price_var = StringVar(value="" if money(product.retail_price) in (None, 0) else str(product.retail_price))
            Label(table, text=product.source_sheet, width=24, anchor="w").grid(row=row_index, column=0, sticky="ew", padx=(6, 8), pady=3)
            Entry(table, textvariable=code_var, width=18).grid(row=row_index, column=1, sticky="ew", padx=(0, 8), pady=3)
            Label(table, text=product.name, width=50, anchor="w").grid(row=row_index, column=2, sticky="ew", padx=(0, 8), pady=3)
            Entry(table, textvariable=price_var, width=14).grid(row=row_index, column=3, sticky="ew", padx=(0, 6), pady=3)
            rows.append((product, code_var, price_var))

        buttons = Frame(container)
        buttons.pack(fill="x")

        def save() -> None:
            edited: list[SourceProduct] = []
            for product, code_var, price_var in rows:
                code = code_var.get().strip()
                price_text = price_var.get().strip()
                if not code and not price_text:
                    continue
                if not valid_product_code(code):
                    messagebox.showerror(APP_TITLE, f"Некорректный код товара:\n\n{product.name}", parent=dialog)
                    return
                if not price_text:
                    messagebox.showerror(APP_TITLE, f"Укажите цену для товара:\n\n{product.name}", parent=dialog)
                    return
                price = price_text.replace(" ", "").replace(",", ".")
                try:
                    number = float(price)
                except ValueError:
                    messagebox.showerror(APP_TITLE, f"Некорректная цена для товара:\n\n{product.name}", parent=dialog)
                    return
                if number <= 0:
                    messagebox.showerror(APP_TITLE, f"Цена должна быть больше 0:\n\n{product.name}", parent=dialog)
                    return
                fixed_price = int(number) if number.is_integer() else number
                edited.append(replace(product, code=code, retail_price=fixed_price))
            result["products"] = edited
            dialog.destroy()

        def skip_all() -> None:
            result["products"] = []
            dialog.destroy()

        Button(buttons, text="Включить исправленные", width=24, command=save).pack(side="right")
        Button(buttons, text="Пропустить все", width=16, command=skip_all).pack(side="right", padx=(0, 8))
        dialog.protocol("WM_DELETE_WINDOW", skip_all)
        self._center_window(dialog)
        self.root.wait_window(dialog)
        return result["products"]

    def ask_edit_groups(self, groups: list[GroupConfig]) -> list[GroupConfig] | None:
        dialog = Toplevel(self.root)
        dialog.title("Группы для выгрузки")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.geometry("860x520")
        dialog.minsize(720, 360)

        result: dict[str, list[GroupConfig] | None] = {"groups": None}
        rows: list[tuple[GroupConfig, StringVar, StringVar]] = []

        container = Frame(dialog, padx=14, pady=12)
        container.pack(fill="both", expand=True)
        Label(
            container,
            text="Проверьте группы перед выгрузкой. ID группы сохраняется как текст и будет подставлен всем товарам из листа.",
            anchor="w",
            justify="left",
            wraplength=820,
        ).pack(fill="x", pady=(0, 10))

        header = Frame(container)
        header.pack(fill="x")
        Label(header, text="Лист прайса", width=28, anchor="w").grid(row=0, column=0, sticky="ew", padx=(0, 8))
        Label(header, text="Группа на выходе", width=36, anchor="w").grid(row=0, column=1, sticky="ew", padx=(0, 8))
        Label(header, text="ID группы", width=18, anchor="w").grid(row=0, column=2, sticky="ew")

        canvas = Canvas(container, highlightthickness=1, highlightbackground="#cfcfcf")
        scrollbar = Scrollbar(container, orient="vertical", command=canvas.yview)
        table = Frame(canvas)
        table.bind("<Configure>", lambda _event: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=table, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True, pady=(4, 10))
        scrollbar.pack(side="right", fill="y", pady=(4, 10))

        def scroll_table(event) -> str:
            if getattr(event, "num", None) == 4:
                canvas.yview_scroll(-1, "units")
            elif getattr(event, "num", None) == 5:
                canvas.yview_scroll(1, "units")
            else:
                delta = int(getattr(event, "delta", 0) or 0)
                if delta:
                    canvas.yview_scroll(int(-1 * (delta / 120)), "units")
            return "break"

        for widget in (dialog, container, canvas, table):
            widget.bind("<MouseWheel>", scroll_table)
            widget.bind("<Button-4>", scroll_table)
            widget.bind("<Button-5>", scroll_table)

        for row_index, group in enumerate(groups):
            name_var = StringVar(value=group.group_name)
            id_var = StringVar(value=group.group_id)
            Label(table, text=group.source_sheet, width=28, anchor="w").grid(row=row_index, column=0, sticky="ew", padx=(6, 8), pady=3)
            name_entry = Entry(table, textvariable=name_var, width=42)
            id_entry = Entry(table, textvariable=id_var, width=20)
            self._bind_entry_shortcuts(name_entry)
            self._bind_entry_shortcuts(id_entry)
            name_entry.bind("<MouseWheel>", scroll_table)
            name_entry.bind("<Button-4>", scroll_table)
            name_entry.bind("<Button-5>", scroll_table)
            id_entry.bind("<MouseWheel>", scroll_table)
            id_entry.bind("<Button-4>", scroll_table)
            id_entry.bind("<Button-5>", scroll_table)
            name_entry.grid(row=row_index, column=1, sticky="ew", padx=(0, 8), pady=3)
            id_entry.grid(row=row_index, column=2, sticky="ew", padx=(0, 6), pady=3)
            rows.append((group, name_var, id_var))

        buttons = Frame(container)
        buttons.pack(fill="x")

        def save() -> None:
            edited: list[GroupConfig] = []
            for group, name_var, id_var in rows:
                name = name_var.get().strip()
                group_id = id_var.get().strip()
                if not name:
                    messagebox.showerror(APP_TITLE, f"Укажите название группы для листа:\n\n{group.source_sheet}", parent=dialog)
                    return
                edited.append(GroupConfig(group.source_sheet, name, group_id))
            result["groups"] = edited
            dialog.destroy()

        def cancel() -> None:
            result["groups"] = None
            dialog.destroy()

        Button(buttons, text="Продолжить", width=18, command=save).pack(side="right")
        Button(buttons, text="Отмена", width=14, command=cancel).pack(side="right", padx=(0, 8))
        dialog.protocol("WM_DELETE_WINDOW", cancel)
        self._center_window(dialog)
        self.root.wait_window(dialog)
        return result["groups"]

    def ask_include_without_code(self, issue: SheetIssue) -> tuple[bool, bool]:
        dialog = Toplevel(self.root)
        dialog.title("Товары без кода")
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        apply_to_all = BooleanVar(value=False)
        result = {"include": False}

        frame = Frame(dialog, padx=18, pady=14)
        frame.pack(fill="both", expand=True)
        Label(
            frame,
            text=(
                f"На листе \"{issue.sheet_name}\" найдено товаров без кода: {issue.count}.\n\n"
                "Включить эти товары в выходной файл?"
            ),
            justify="left",
            wraplength=520,
        ).pack(anchor="w")
        Checkbutton(frame, text="Применить этот ответ ко всем последующим листам", variable=apply_to_all).pack(
            anchor="w", pady=(12, 8)
        )

        buttons = Frame(frame)
        buttons.pack(anchor="e")

        def choose(include: bool) -> None:
            result["include"] = include
            dialog.destroy()

        Button(buttons, text="Да, включить", width=16, command=lambda: choose(True)).pack(side="left", padx=(0, 8))
        Button(buttons, text="Нет, пропустить", width=16, command=lambda: choose(False)).pack(side="left")

        dialog.protocol("WM_DELETE_WINDOW", lambda: choose(False))
        self._center_window(dialog)
        self.root.wait_window(dialog)
        return result["include"], apply_to_all.get()

    def _show_error(self, exc: Exception) -> None:
        self.convert_button.config(state="normal")
        self.status.set("Ошибка конвертации.")
        messagebox.showerror(APP_TITLE, f"Не удалось выполнить конвертацию:\n\n{exc}")

    def _show_success(
        self,
        output: Path,
        count: int,
        missing_group_ids: dict[str, int],
        skipped_without_code: dict[str, int],
        skipped_without_price: dict[str, int],
        skipped_invalid_code: dict[str, int],
        skipped_without_group_id: dict[str, int],
        exported_images_count: int,
        images_output_dir: str | None,
    ) -> None:
        self.convert_button.config(state="normal")
        self.status.set(f"Готово. Товаров: {count}.")
        notes = []
        if skipped_without_code:
            rows = "\n".join(f"- {name}: {qty}" for name, qty in sorted(skipped_without_code.items()))
            notes.append(f"Пропущены товары без кода:\n{rows}")
        if skipped_without_price:
            rows = "\n".join(f"- {name}: {qty}" for name, qty in sorted(skipped_without_price.items()))
            notes.append(f"Пропущены товары без цены или с ценой 0:\n{rows}")
        if skipped_invalid_code:
            rows = "\n".join(f"- {name}: {qty}" for name, qty in sorted(skipped_invalid_code.items()))
            notes.append(f"Пропущены товары с невалидным кодом:\n{rows}")
        if missing_group_ids:
            rows = "\n".join(f"- {name}: {qty}" for name, qty in sorted(missing_group_ids.items()))
            notes.append(f"Пропущены товары из групп без ID:\n{rows}")
        if exported_images_count and images_output_dir:
            notes.append(f"Картинки товаров сохранены: {exported_images_count}\nПапка:\n{images_output_dir}")
        suffix = "\n\n" + "\n\n".join(notes) if notes else ""
        messagebox.showinfo(APP_TITLE, f"Конвертация завершена.\n\nФайл:\n{output.resolve()}{suffix}")

    def run(self) -> None:
        self.root.mainloop()


def main() -> None:
    ConverterApp().run()


if __name__ == "__main__":
    main()
