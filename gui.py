from __future__ import annotations

import threading
from dataclasses import replace
from pathlib import Path
from tkinter import BooleanVar, Button, Canvas, Checkbutton, Entry, Frame, Label, Radiobutton, Scrollbar, StringVar, Tk, Toplevel, filedialog, messagebox

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
        self.root.geometry("760x210")
        self.root.minsize(680, 210)

        self.source_path = StringVar()
        self.output_path = StringVar()
        self.status = StringVar(value="Выберите исходный прайс для конвертации.")
        self.apply_without_code_to_all: bool | None = None

        self._build_ui()

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

        self.convert_button = Button(container, text="Конвертировать", command=self.start_convert, height=2)
        self.convert_button.grid(row=2, column=0, columnspan=3, sticky="ew", pady=(16, 8))

        Label(container, textvariable=self.status, anchor="w").grid(row=3, column=0, columnspan=3, sticky="ew")

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
        self.status.set("Конвертация выполняется...")
        thread = threading.Thread(target=self._convert_in_background, args=(source, output, edited_group_configs), daemon=True)
        thread.start()

    def _convert_in_background(self, source: Path, output: Path, group_configs: list[GroupConfig]) -> None:
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
                edited.append(replace(product, code=code, retail_price=fixed_price, wholesale_price=fixed_price))
            result["products"] = edited
            dialog.destroy()

        def skip_all() -> None:
            result["products"] = []
            dialog.destroy()

        Button(buttons, text="Включить исправленные", width=24, command=save).pack(side="right")
        Button(buttons, text="Пропустить все", width=16, command=skip_all).pack(side="right", padx=(0, 8))
        dialog.protocol("WM_DELETE_WINDOW", skip_all)
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

        for row_index, group in enumerate(groups):
            name_var = StringVar(value=group.group_name)
            id_var = StringVar(value=group.group_id)
            Label(table, text=group.source_sheet, width=28, anchor="w").grid(row=row_index, column=0, sticky="ew", padx=(6, 8), pady=3)
            Entry(table, textvariable=name_var, width=42).grid(row=row_index, column=1, sticky="ew", padx=(0, 8), pady=3)
            Entry(table, textvariable=id_var, width=20).grid(row=row_index, column=2, sticky="ew", padx=(0, 6), pady=3)
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
