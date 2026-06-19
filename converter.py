from __future__ import annotations

import argparse
import hashlib
import html
import json
import re
import sys
from concurrent.futures import ThreadPoolExecutor, as_completed
from io import BytesIO
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable

from PIL import Image as PILImage
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.worksheet.worksheet import Worksheet


PRODUCT_SHEET = "Export Products Sheet"
GROUP_SHEET = "Export Groups Sheet"
DEFAULT_SCHEMA_FILE = "satu_schema.json"
DEFAULT_GROUP_CACHE_FILE = "group_cache.json"


@dataclass
class SourceProduct:
    code: str | None
    name: str
    name_says_unavailable: bool
    unit: str | None
    stock: Any
    retail_price: Any
    wholesale_price: Any
    sko_price: Any
    comment: str | None
    quantity: Any
    source_sheet: str
    row_number: int


@dataclass
class ExportedProduct:
    product: SourceProduct
    output_row: int


@dataclass
class ConversionStats:
    converted_count: int
    missing_group_ids: dict[str, int]
    skipped_without_code: dict[str, int]
    skipped_without_price: dict[str, int]
    skipped_invalid_code: dict[str, int]
    skipped_without_group_id: dict[str, int]
    exported_images_count: int
    images_output_dir: str | None


@dataclass
class WorksheetImage:
    image: Any
    image_bytes: bytes
    start_row: int
    end_row: int
    start_col: int | None
    end_col: int | None


@dataclass
class ImageExportCandidate:
    output_row: int
    code: str
    image_bytes: bytes
    extension: str


@dataclass
class ImageWriteTask:
    output_row: int
    image_bytes: bytes
    target_path: Path
    relative_path: str


@dataclass
class GroupConfig:
    source_sheet: str
    group_name: str
    group_id: str


@dataclass
class SheetColumns:
    header_row: int
    code: int | None
    name: int
    unit: int | None
    stock: int | None
    retail_price: int
    wholesale_price: int | None
    sko_price: int | None
    comment: int | None
    quantity: int | None


@dataclass
class SheetIssue:
    sheet_name: str
    issue_type: str
    count: int


IssueHandler = Callable[[SheetIssue], bool]
DuplicateChoiceHandler = Callable[[str, list[SourceProduct]], SourceProduct]
ProductIssueHandler = Callable[[list[SourceProduct]], list[SourceProduct]]
ProgressCallback = Callable[[int, int], None]


EMU_PER_PIXEL = 9525
IMAGE_WORKER_COUNT = 4
UNAVAILABLE_NAME_PATTERN = re.compile(
    r"(?i)(?:\s*[\(\[\{,;:–—-]\s*)?\bнет\s+в\s+наличии\b(?:\s*[\)\]\},;:–—-]\s*)?"
)


def normalize_text(value: Any) -> str:
    if value is None:
        return ""
    return re.sub(r"\s+", " ", str(value).strip())


def normalize_key(value: Any) -> str:
    return normalize_text(value).casefold()


def name_without_availability_marker(value: Any) -> tuple[str, bool]:
    text = normalize_text(value)
    cleaned, replacements = UNAVAILABLE_NAME_PATTERN.subn(" ", text)
    cleaned = normalize_text(re.sub(r"\s+([,.;:])", r"\1", cleaned))
    return cleaned, replacements > 0


def compact_key(value: Any) -> str:
    return re.sub(r"[^0-9a-zа-яё]+", "", normalize_key(value))


def unit_key(value: Any) -> str:
    return compact_key(normalize_text(value).replace("²", "2").replace("³", "3"))


ALLOWED_UNITS = {
    "шт.",
    "100 шт.",
    "10 шт.",
    "тыс. шт.",
    "т.у.шт.",
    "т",
    "кг",
    "г",
    "куб.м",
    "л",
    "кв.м",
    "кв.см",
    "кв.фут",
    "кв.дм",
    "м",
    "км",
    "дал",
    "мешок",
    "пара",
    "чел.",
    "упаковка",
    "тысяча",
    "сотка",
    "га",
    "пог.м",
    "ящик",
    "ведро",
    "банка",
    "бут",
    "канистра",
    "пач",
    "мм",
    "мл",
    "гр/кв.м",
    "кг/кв.м",
    "100 г",
    "комплект",
    "набор",
    "моток",
    "рулон",
    "услуга",
    "см",
    "секция",
    "бухта",
    "объект",
    "страница",
    "т/км",
    "сутки",
    "ватт",
    "лист",
    "карат",
    "минута",
    "кВт",
    "мВт",
    "бобина",
    "паллетоместо",
    "смена",
    "куб.дм",
    "рейс",
    "колесо",
    "ярд",
    "баллон",
    "бочка",
    "коробка",
    "10 см",
    "автоцистерна",
    "еврокуб",
    "100 мл",
    "ед.",
    "час",
    "день",
    "неделя",
    "месяц",
    "два месяца",
    "квартал",
    "полгода",
    "год",
    "номер",
    "птицеместо",
    "стакан",
    "флакон",
    "ампула",
    "тираж",
    "выезд",
    "доза",
    "шприц-туба",
    "таблетка",
    "тюбик",
    "блистер",
    "50 г.",
    "партия",
    "посевная единица",
    "слово",
    "50 шт.",
    "20 шт.",
    "5 г.",
    "2 г.",
    "10 г.",
    "гигакалория",
}

UNIT_ALIASES = {
    "шт": "шт.",
    "штука": "шт.",
    "штуки": "шт.",
    "штук": "шт.",
    "ед": "ед.",
    "компл": "комплект",
    "комп": "комплект",
    "комплект": "комплект",
    "комплекты": "комплект",
    "м2": "кв.м",
    "м²": "кв.м",
    "квм": "кв.м",
    "квметр": "кв.м",
    "квметра": "кв.м",
    "квадратныйметр": "кв.м",
    "квадратныхметров": "кв.м",
    "м3": "куб.м",
    "м³": "куб.м",
    "кубм": "куб.м",
    "кубметр": "куб.м",
    "погм": "пог.м",
    "пм": "пог.м",
    "мп": "пог.м",
    "упак": "упаковка",
    "уп": "упаковка",
    "упаковка": "упаковка",
    "пачка": "пач",
    "пачки": "пач",
    "рул": "рулон",
    "рулон": "рулон",
    "наб": "набор",
    "набор": "набор",
    "секц": "секция",
    "секция": "секция",
    "бухт": "бухта",
    "бухта": "бухта",
    "л": "л",
    "литр": "л",
    "литры": "л",
    "кг": "кг",
    "гр": "г",
    "г": "г",
    "т": "т",
    "м": "м",
    "см": "см",
    "мм": "мм",
    "мл": "мл",
    "лист": "лист",
    "кор": "коробка",
    "короб": "коробка",
    "коробка": "коробка",
    "квт": "кВт",
    "мвт": "мВт",
}


def normalize_unit(value: Any) -> str | None:
    unit = normalize_text(value)
    if not unit:
        return None
    allowed_by_key = {unit_key(allowed): allowed for allowed in ALLOWED_UNITS}
    aliases_by_key = {unit_key(alias): target for alias, target in UNIT_ALIASES.items()}
    key = unit_key(unit)
    if key in aliases_by_key:
        return aliases_by_key[key]
    if key in allowed_by_key:
        return allowed_by_key[key]
    return "шт."


def normalize_output_unit(value: Any) -> str:
    return normalize_unit(value) or "шт."


def safe_filename_part(value: Any) -> str:
    text = normalize_text(value)
    safe = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "_", text)
    safe = safe.strip(" ._")
    return safe or "product"


def image_extension(image: Any) -> str:
    path = normalize_text(getattr(image, "path", ""))
    suffix = Path(path).suffix.lower()
    if suffix in {".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".webp"}:
        return suffix
    return ".png"


def marker_row(marker: Any) -> int | None:
    row = getattr(marker, "row", None)
    if row is None:
        return None
    return int(row) + 1


def marker_end_row(marker: Any) -> int | None:
    row = getattr(marker, "row", None)
    if row is None:
        return None
    row_offset = getattr(marker, "rowOff", 0) or 0
    return int(row) + (1 if row_offset else 0)


def marker_col(marker: Any) -> int | None:
    col = getattr(marker, "col", None)
    if col is None:
        return None
    return int(col) + 1


def marker_end_col(marker: Any) -> int | None:
    col = getattr(marker, "col", None)
    if col is None:
        return None
    col_offset = getattr(marker, "colOff", 0) or 0
    return int(col) + (1 if col_offset else 0)


def row_height_pixels(ws: Worksheet, row_number: int) -> float:
    height = ws.row_dimensions[row_number].height or ws.sheet_format.defaultRowHeight or 15
    return float(height) * 96 / 72


def emu_to_pixels(value: Any) -> float | None:
    if value is None:
        return None
    return float(value) / EMU_PER_PIXEL


def image_display_height_pixels(image: Any) -> float | None:
    anchor = getattr(image, "anchor", None)
    ext = getattr(anchor, "ext", None)
    height = emu_to_pixels(getattr(ext, "cy", None))
    if height:
        return height
    fallback = getattr(image, "height", None)
    return float(fallback) if fallback else None


def image_display_width_pixels(image: Any) -> float | None:
    anchor = getattr(image, "anchor", None)
    ext = getattr(anchor, "ext", None)
    width = emu_to_pixels(getattr(ext, "cx", None))
    if width:
        return width
    fallback = getattr(image, "width", None)
    return float(fallback) if fallback else None


def column_width_pixels(ws: Worksheet, column_number: int) -> float:
    from openpyxl.utils import get_column_letter

    width = ws.column_dimensions[get_column_letter(column_number)].width or 8.43
    return float(width) * 7 + 5


def estimated_end_row(ws: Worksheet, start_row: int, image: Any) -> int:
    height = image_display_height_pixels(image)
    if not height:
        return start_row

    remaining = float(height)
    row_number = start_row
    while remaining > row_height_pixels(ws, row_number) and row_number < ws.max_row:
        remaining -= row_height_pixels(ws, row_number)
        row_number += 1
    return row_number


def estimated_end_col(ws: Worksheet, start_col: int, image: Any) -> int:
    width = image_display_width_pixels(image)
    if not width:
        return start_col

    remaining = float(width)
    col_number = start_col
    max_col = max(ws.max_column, start_col)
    while remaining > column_width_pixels(ws, col_number) and col_number < max_col:
        remaining -= column_width_pixels(ws, col_number)
        col_number += 1
    return col_number


def product_row_range(ws: Worksheet, row_number: int, columns: SheetColumns | None) -> tuple[int, int]:
    if columns is None:
        return row_number, row_number

    relevant_columns = [
        column + 1
        for column in (columns.code, columns.name, columns.retail_price)
        if column is not None
    ]
    start_row = row_number
    end_row = row_number
    for merged_range in ws.merged_cells.ranges:
        if not (merged_range.min_row <= row_number <= merged_range.max_row):
            continue
        if not any(merged_range.min_col <= column <= merged_range.max_col for column in relevant_columns):
            continue
        start_row = min(start_row, merged_range.min_row)
        end_row = max(end_row, merged_range.max_row)
    return start_row, end_row


def is_decorative_new_badge(image_bytes: bytes) -> bool:
    try:
        with PILImage.open(BytesIO(image_bytes)) as image:
            rgba = image.convert("RGBA")
            width, height = rgba.size
            pixels = list(rgba.getdata())
    except Exception:
        return False

    if not pixels or width < 80 or height < 80:
        return False

    opaque = [pixel for pixel in pixels if pixel[3] > 20]
    if not opaque:
        return False

    transparent_ratio = (len(pixels) - len(opaque)) / len(pixels)
    red_ratio = sum(1 for r, g, b, _a in opaque if r > 150 and r > g * 1.6 and r > b * 1.6) / len(opaque)
    white_ratio = sum(1 for r, g, b, _a in opaque if r > 210 and g > 210 and b > 210) / len(opaque)
    squareish = 0.75 <= width / height <= 1.35

    return squareish and transparent_ratio > 0.45 and red_ratio > 0.65 and white_ratio > 0.015


def is_number(value: Any) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def money(value: Any) -> int | float | None:
    if value is None or value == "":
        return None
    if is_number(value):
        return int(value) if float(value).is_integer() else value
    text = normalize_text(value).replace("\u00a0", "").replace(" ", "").replace(",", ".")
    try:
        number = float(text)
    except ValueError:
        return None
    return int(number) if number.is_integer() else number


def valid_price(value: Any) -> bool:
    price = money(value)
    return price is not None and price != 0


def valid_product_code(value: Any) -> bool:
    code = normalize_text(value)
    if not code:
        return False
    key = compact_key(code)
    if key in {"0", "код", "кодтовара", "артикул"}:
        return False
    digits = re.sub(r"\D+", "", code)
    has_letters = bool(re.search(r"[A-Za-zА-Яа-яЁё]", code))
    if digits and not has_letters and re.fullmatch(r"0+", digits):
        return False
    return True


def quantity(value: Any) -> int | float | None:
    return money(value)


def availability(stock: Any, qty: Any, force_unavailable: bool = False) -> str:
    if force_unavailable:
        return "-"
    stock_text = normalize_key(stock)
    numeric_qty = quantity(qty)
    if numeric_qty and numeric_qty > 0:
        return "+"
    if any(marker in stock_text for marker in ("нет", "отсутств", "уточ")):
        return "-"
    if any(marker in stock_text for marker in ("налич", "есть")):
        return "+"
    return "+"


def first_matching_header(headers: dict[str, int], names: tuple[str, ...]) -> int | None:
    for name in names:
        for header, index in headers.items():
            if name in header:
                return index
    return None


def detect_columns(row: tuple[Any, ...], row_index: int) -> SheetColumns | None:
    headers = {compact_key(value): index for index, value in enumerate(row) if compact_key(value)}

    code = first_matching_header(headers, ("код", "артикул"))
    name = first_matching_header(headers, ("наименование", "название", "товар"))
    retail = first_matching_header(headers, ("ценарозница", "розница", "цена"))

    unit = first_matching_header(headers, ("едизм", "единицаизмерения", "изм"))
    stock = first_matching_header(headers, ("остаток", "наличие", "склад"))
    if retail is None:
        return None
    wholesale = next_non_empty_column(row, retail, 1)
    sko = next_non_empty_column(row, retail, 2)
    comment = first_matching_header(headers, ("ссылкахкомментарий", "ссылкакомментарий", "допкомментарий", "комментарий"))
    quantity_col = first_matching_header(headers, ("количество", "колво", "остатокчисло"))

    if retail is None:
        return None

    if name is None and code is not None:
        stop_at = min(index for index in (unit, stock, retail) if index is not None)
        for index in range(code + 1, stop_at):
            if normalize_text(row[index]):
                name = index
                break

    if name is None:
        return None

    return SheetColumns(
        header_row=row_index,
        code=code,
        name=name,
        unit=unit,
        stock=stock,
        retail_price=retail,
        wholesale_price=wholesale,
        sko_price=sko,
        comment=comment,
        quantity=quantity_col,
    )


def find_sheet_columns(ws: Worksheet) -> SheetColumns | None:
    for row_index, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 20), values_only=True), start=1):
        columns = detect_columns(row, row_index)
        if columns is not None:
            return columns
    return None


def find_header_row(ws: Worksheet) -> int | None:
    columns = find_sheet_columns(ws)
    return columns.header_row if columns else None


def cell_value(row: tuple[Any, ...], index: int | None) -> Any:
    if index is None or index >= len(row):
        return None
    return row[index]


def next_non_empty_column(row: tuple[Any, ...], start_index: int | None, steps_ahead: int) -> int | None:
    if start_index is None:
        return None
    found = 0
    for index in range(start_index + 1, len(row)):
        if normalize_text(row[index]):
            found += 1
            if found == steps_ahead:
                return index
    return None


def optional_price(value: Any) -> int | float | None:
    price = money(value)
    if price is None or price <= 0:
        return None
    return price


def default_issue_handler(issue: SheetIssue) -> bool:
    return issue.issue_type != "without_code"


def product_from_row(row: tuple[Any, ...], columns: SheetColumns, source_sheet: str, row_number: int) -> SourceProduct:
    name, name_says_unavailable = name_without_availability_marker(cell_value(row, columns.name))
    return SourceProduct(
        code=normalize_text(cell_value(row, columns.code)) or None,
        name=name,
        name_says_unavailable=name_says_unavailable,
        unit=normalize_unit(cell_value(row, columns.unit)),
        stock=cell_value(row, columns.stock),
        retail_price=cell_value(row, columns.retail_price),
        wholesale_price=cell_value(row, columns.wholesale_price),
        sko_price=cell_value(row, columns.sko_price),
        comment=normalize_text(cell_value(row, columns.comment)) or None,
        quantity=cell_value(row, columns.quantity),
        source_sheet=source_sheet,
        row_number=row_number,
    )


def count_by_sheet(products: list[SourceProduct]) -> dict[str, int]:
    counts: dict[str, int] = {}
    for product in products:
        counts[product.source_sheet] = counts.get(product.source_sheet, 0) + 1
    return counts


def read_source_products(
    path: Path,
    issue_handler: IssueHandler | None = None,
    product_issue_handler: ProductIssueHandler | None = None,
) -> tuple[list[SourceProduct], dict[str, int], dict[str, int], dict[str, int]]:
    workbook = load_workbook(path, read_only=True, data_only=True)
    products: list[SourceProduct] = []
    skipped_without_code: dict[str, int] = {}
    products_with_issues: list[SourceProduct] = []
    issue_flags: list[tuple[bool, bool]] = []

    for ws in workbook.worksheets:
        columns = find_sheet_columns(ws)
        if columns is None:
            continue

        for row_number, row in enumerate(ws.iter_rows(min_row=columns.header_row + 1, max_row=ws.max_row, values_only=True), start=columns.header_row + 1):
            name, _ = name_without_availability_marker(cell_value(row, columns.name))

            if not name:
                continue
            product = product_from_row(row, columns, ws.title, row_number)
            bad_price = not valid_price(product.retail_price)
            bad_code = not valid_product_code(product.code)
            if bad_price or bad_code:
                products_with_issues.append(product)
                issue_flags.append((bad_price, bad_code))
                continue
            products.append(product)

    workbook.close()
    edited_products: list[SourceProduct] = []
    if products_with_issues and product_issue_handler:
        edited_products = [
            product
            for product in product_issue_handler(products_with_issues)
            if valid_price(product.retail_price) and valid_product_code(product.code)
        ]
        products.extend(edited_products)

    fixed_counts: dict[tuple[str, str], int] = {}
    for product in edited_products:
        fixed_key = (normalize_key(product.source_sheet), normalize_key(product.name))
        fixed_counts[fixed_key] = fixed_counts.get(fixed_key, 0) + 1
    skipped_without_price: dict[str, int] = {}
    skipped_invalid_code: dict[str, int] = {}
    for product, (bad_price, bad_code) in zip(products_with_issues, issue_flags):
        fixed_key = (normalize_key(product.source_sheet), normalize_key(product.name))
        if fixed_counts.get(fixed_key, 0) > 0:
            fixed_counts[fixed_key] -= 1
            continue
        if bad_price:
            skipped_without_price[product.source_sheet] = skipped_without_price.get(product.source_sheet, 0) + 1
        if bad_code:
            skipped_invalid_code[product.source_sheet] = skipped_invalid_code.get(product.source_sheet, 0) + 1

    return products, skipped_without_code, skipped_without_price, skipped_invalid_code


def read_schema(schema_path: Path) -> dict[str, Any]:
    with schema_path.open("r", encoding="utf-8") as file:
        return json.load(file)


def deduplicate_products(
    products: list[SourceProduct],
    duplicate_handler: DuplicateChoiceHandler | None = None,
) -> list[SourceProduct]:
    by_code: dict[str, list[SourceProduct]] = {}
    without_code: list[SourceProduct] = []
    for product in products:
        if not product.code:
            without_code.append(product)
            continue
        by_code.setdefault(product.code, []).append(product)

    result: list[SourceProduct] = []
    for code, matches in by_code.items():
        if len(matches) == 1:
            result.append(matches[0])
            continue
        names = {normalize_key(product.name) for product in matches}
        if len(names) == 1:
            result.append(matches[0])
            continue
        if duplicate_handler is None:
            readable = "\n".join(f"- {product.name} ({product.source_sheet})" for product in matches)
            raise ValueError(f"Найден код товара с разными названиями: {code}\n{readable}")
        result.append(duplicate_handler(code, matches))

    result.extend(without_code)
    return result


def groups_from_schema(schema: dict[str, Any]) -> dict[str, tuple[Any, str]]:
    groups: dict[str, tuple[Any, str]] = {}
    for group in schema.get("groups", []):
        group_id = normalize_text(group.get("number"))
        group_name = normalize_text(group.get("name"))
        if group_name:
            groups[normalize_key(group_name)] = (group_id, group_name)
    return groups


def load_aliases(path: Path | None) -> dict[str, str]:
    if path is None or not path.exists():
        return {}
    with path.open("r", encoding="utf-8") as file:
        data = json.load(file)
    return {normalize_key(key): str(value) for key, value in data.items()}


def app_base_dir() -> Path:
    return Path(getattr(sys, "_MEIPASS", Path(__file__).parent))


def default_schema_path() -> Path:
    external = Path(sys.argv[0]).resolve().with_name(DEFAULT_SCHEMA_FILE)
    if external.exists():
        return external
    return app_base_dir() / DEFAULT_SCHEMA_FILE


def default_aliases_path() -> Path:
    external = Path(sys.argv[0]).resolve().with_name("group_aliases.json")
    if external.exists():
        return external
    return app_base_dir() / "group_aliases.json"


def default_group_cache_path() -> Path:
    return Path(sys.argv[0]).resolve().with_name(DEFAULT_GROUP_CACHE_FILE)


def load_group_cache(path: Path | None = None) -> dict[str, GroupConfig]:
    cache_path = path or default_group_cache_path()
    if not cache_path.exists():
        return {}
    with cache_path.open("r", encoding="utf-8") as file:
        data = json.load(file)
    raw_groups = data.get("groups", data)
    if not isinstance(raw_groups, dict):
        return {}

    cache: dict[str, GroupConfig] = {}
    for source_sheet, item in raw_groups.items():
        if not isinstance(item, dict):
            continue
        group_name = normalize_text(item.get("group_name") or item.get("name"))
        group_id = normalize_text(item.get("group_id") or item.get("id"))
        if not group_name:
            continue
        source_name = normalize_text(item.get("source_sheet") or source_sheet)
        cache[normalize_key(source_name)] = GroupConfig(source_name, group_name, group_id)
    return cache


def save_group_cache(configs: list[GroupConfig], path: Path | None = None) -> None:
    cache_path = path or default_group_cache_path()
    cache_path.parent.mkdir(parents=True, exist_ok=True)
    data = {
        "groups": {
            config.source_sheet: {
                "source_sheet": config.source_sheet,
                "group_name": config.group_name,
                "group_id": config.group_id,
            }
            for config in configs
        }
    }
    with cache_path.open("w", encoding="utf-8") as file:
        json.dump(data, file, ensure_ascii=False, indent=2)


def detect_source_groups(
    source_path: Path,
    aliases_path: Path | None = None,
    schema_path: Path | None = None,
    cache_path: Path | None = None,
) -> list[GroupConfig]:
    workbook = load_workbook(source_path, read_only=True, data_only=True)
    schema = read_schema(schema_path or default_schema_path())
    groups = groups_from_schema(schema)
    aliases = load_aliases(aliases_path)
    cache = load_group_cache(cache_path)
    detected: list[GroupConfig] = []

    for ws in workbook.worksheets:
        if find_sheet_columns(ws) is None:
            continue
        cached = cache.get(normalize_key(ws.title))
        if cached:
            detected.append(GroupConfig(ws.title, cached.group_name, cached.group_id))
            continue
        group_id, group_name = group_for_sheet(ws.title, groups, aliases)
        detected.append(GroupConfig(ws.title, group_name, normalize_text(group_id)))

    workbook.close()
    return detected


def group_for_sheet(
    sheet_name: str,
    groups: dict[str, tuple[Any, str]],
    aliases: dict[str, str],
    group_configs: dict[str, GroupConfig] | None = None,
) -> tuple[Any, str]:
    sheet_key = normalize_key(sheet_name)
    if group_configs and sheet_key in group_configs:
        config = group_configs[sheet_key]
        return config.group_id, config.group_name
    target_name = aliases.get(sheet_key, normalize_text(sheet_name))
    group = groups.get(normalize_key(target_name))
    if group:
        return group
    return None, target_name


def description_for(product: SourceProduct) -> str:
    parts = [f"<p>{html.escape(product.name)}</p>"]
    if product.comment:
        parts.append(f"<p>{html.escape(product.comment)}</p>")
    return "\n".join(parts)


def write_value(row: list[Any], indexes: dict[str, int], column: str, value: Any) -> None:
    index = indexes.get(column)
    if index is not None:
        row[index - 1] = value


def product_to_target_row(
    product: SourceProduct,
    max_columns: int,
    indexes: dict[str, int],
    groups: dict[str, tuple[Any, str]],
    aliases: dict[str, str],
    group_configs: dict[str, GroupConfig] | None = None,
) -> list[Any]:
    group_id, group_name = group_for_sheet(product.source_sheet, groups, aliases, group_configs)
    qty = quantity(product.quantity)
    retail = money(product.retail_price)
    wholesale = optional_price(product.wholesale_price)
    sko = optional_price(product.sko_price)

    row: list[Any] = [None] * max_columns
    write_value(row, indexes, "Код_товара", product.code)
    write_value(row, indexes, "Название_позиции", product.name)
    write_value(row, indexes, "Поисковые_запросы", product.name)
    write_value(row, indexes, "Описание", description_for(product))
    write_value(row, indexes, "Тип_товара", "u")
    write_value(row, indexes, "Цена", retail)
    write_value(row, indexes, "Валюта", "KZT")
    write_value(row, indexes, "Единица_измерения", normalize_output_unit(product.unit))
    write_value(row, indexes, "Оптовая_цена", wholesale)
    write_value(row, indexes, "\u0421\u041a\u041e_\u0446\u0435\u043d\u0430", sko)
    write_value(row, indexes, "Минимальный_заказ_опт", retail)
    write_value(row, indexes, "Наличие", availability(product.stock, qty, product.name_says_unavailable))
    write_value(row, indexes, "Количество", qty)
    write_value(row, indexes, "Номер_группы", group_id)
    write_value(row, indexes, "Название_группы", group_name)
    write_value(row, indexes, "Цена_от", "-")
    write_value(row, indexes, "Личные_заметки", f"Источник: {product.source_sheet}")

    if product.code:
        write_value(row, indexes, "Уникальный_идентификатор", product.code)

    return row


def write_header(ws: Worksheet, headers: list[str]) -> None:
    ws.append(headers)
    fill = PatternFill("solid", fgColor="D9EAF7")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = fill
    ws.freeze_panes = "A2"


def create_output_workbook(schema: dict[str, Any]) -> tuple[Workbook, Worksheet, dict[str, int]]:
    workbook = Workbook()
    product_ws = workbook.active
    product_ws.title = PRODUCT_SHEET
    headers = [str(header) if header is not None else "" for header in schema["product_headers"]]
    write_header(product_ws, headers)
    indexes = {header: index for index, header in enumerate(headers, start=1) if header}

    group_ws = workbook.create_sheet(GROUP_SHEET)
    group_headers = [
        "Номер_группы",
        "Название_группы",
        "Идентификатор_группы",
        "Номер_родителя",
        "Идентификатор_родителя",
        "HTML_заголовок_группы",
        "HTML_описание_группы",
        "Описание_группы_до_списка_товарных_позиций",
        "Описание_группы_после_списка_товарных_позиций",
        "Ссылка_изображения_группы",
    ]
    write_header(group_ws, group_headers)
    for group in schema.get("groups", []):
        if not normalize_text(group.get("number")):
            continue
        group_ws.append(
            [
                group.get("number"),
                group.get("name"),
                group.get("identifier"),
                group.get("parent_number"),
                group.get("parent_identifier"),
                group.get("html_title"),
                group.get("html_description"),
                group.get("description_before"),
                group.get("description_after"),
                group.get("image_url"),
            ]
        )

    return workbook, product_ws, indexes


def append_configured_groups(workbook: Workbook, configs: list[GroupConfig]) -> None:
    if GROUP_SHEET not in workbook.sheetnames:
        return
    group_ws = workbook[GROUP_SHEET]
    existing_rows = {normalize_key(group_ws.cell(row=row, column=2).value): row for row in range(2, group_ws.max_row + 1)}
    for config in configs:
        group_name = normalize_text(config.group_name)
        group_id = normalize_text(config.group_id)
        if not group_name or not group_id:
            continue
        group_key = normalize_key(group_name)
        if group_key in existing_rows:
            if group_id:
                group_ws.cell(row=existing_rows[group_key], column=1, value=group_id)
            continue
        group_ws.append([config.group_id, group_name, None, None, None, None, None, None, None, None])
        existing_rows[group_key] = group_ws.max_row


def images_output_dir_for(output_path: Path) -> Path:
    return output_path.with_name(f"{output_path.stem}_images")


class ExcelImageExtractor:
    def __init__(self, ws: Worksheet):
        self.ws = ws

    def collect(self) -> list[WorksheetImage]:
        collected: list[WorksheetImage] = []
        for image in list(getattr(self.ws, "_images", [])):
            worksheet_image = self._worksheet_image(image)
            if worksheet_image is not None:
                collected.append(worksheet_image)
        return collected

    def _worksheet_image(self, image: Any) -> WorksheetImage | None:
        anchor = getattr(image, "anchor", None)
        start_marker = getattr(anchor, "_from", None)
        start_row = marker_row(start_marker)
        if start_row is None:
            return None

        start_col = marker_col(start_marker)
        end_marker = getattr(anchor, "_to", None)
        end_row = marker_end_row(end_marker) if end_marker is not None else None
        end_col = marker_end_col(end_marker) if end_marker is not None else None

        if end_row is None:
            end_row = estimated_end_row(self.ws, start_row, image)
        if start_col is not None and end_col is None:
            end_col = estimated_end_col(self.ws, start_col, image)

        try:
            image_bytes = image._data()
        except Exception:
            return None

        return WorksheetImage(
            image=image,
            image_bytes=image_bytes,
            start_row=start_row,
            end_row=max(start_row, end_row),
            start_col=start_col,
            end_col=end_col,
        )


class ProductImageMatcher:
    MAX_NEARBY_ROW_GAP = 1

    def __init__(self, images: list[WorksheetImage], columns: SheetColumns | None):
        self.images = images
        self.min_col, self.max_col = self._source_table_bounds(columns)

    def images_for_range(self, start_row: int, end_row: int) -> list[WorksheetImage]:
        ranked_matches: list[tuple[tuple[int, int, int, int], WorksheetImage]] = []
        for image in self.images:
            if not self._inside_source_table(image):
                continue
            score = self._match_score(image, start_row, end_row)
            if score is None:
                continue
            ranked_matches.append((score, image))

        ranked_matches.sort(key=lambda item: item[0])
        return [image for _score, image in ranked_matches]

    def _match_score(
        self,
        image: WorksheetImage,
        start_row: int,
        end_row: int,
    ) -> tuple[int, int, int, int] | None:
        if image.end_row < start_row:
            gap = start_row - image.end_row
        elif image.start_row > end_row:
            gap = image.start_row - end_row
        else:
            gap = 0

        if gap > self.MAX_NEARBY_ROW_GAP:
            return None

        target_center_twice = start_row + end_row
        image_center_twice = image.start_row + image.end_row
        center_distance_twice = min(
            abs(image_center_twice - (start_row * 2)),
            abs(image_center_twice - (end_row * 2)),
            abs(image_center_twice - target_center_twice),
        )
        span = max(0, image.end_row - image.start_row)
        position_bias = 0
        if image.start_row > end_row:
            position_bias = 2
        elif image.end_row < start_row:
            position_bias = 1
        return (center_distance_twice, span, position_bias, gap)

    def _source_table_bounds(self, columns: SheetColumns | None) -> tuple[int | None, int | None]:
        if columns is None:
            return None, None
        detected_columns = [
            value
            for value in (
                columns.code,
                columns.name,
                columns.unit,
                columns.stock,
                columns.retail_price,
                columns.wholesale_price,
                columns.sko_price,
                columns.comment,
                columns.quantity,
            )
            if value is not None
        ]
        if not detected_columns:
            return None, None
        return 1, max(detected_columns) + 4

    def _inside_source_table(self, image: WorksheetImage) -> bool:
        if self.min_col is None or self.max_col is None:
            return True
        if image.start_col is None or image.end_col is None:
            return True
        return image.start_col <= self.max_col and image.end_col >= self.min_col


def prepare_images_output_dir(output_dir: Path) -> None:
    output_dir.mkdir(parents=True, exist_ok=True)
    for existing in output_dir.iterdir():
        if existing.is_file():
            existing.unlink()


def analyze_sheet_image_exports(
    source_path: Path,
    sheet_name: str,
    sheet_products: list[ExportedProduct],
) -> list[ImageExportCandidate]:
    workbook = load_workbook(source_path, read_only=False, data_only=True)
    try:
        ws = workbook[sheet_name]
        images = ExcelImageExtractor(ws).collect()
        if not images:
            return []

        columns = find_sheet_columns(ws)
        matcher = ProductImageMatcher(images, columns)
        candidates: list[ImageExportCandidate] = []
        for exported in sheet_products:
            if not exported.product.code:
                continue
            start_row, end_row = product_row_range(ws, exported.product.row_number, columns)
            for worksheet_image in matcher.images_for_range(start_row, end_row):
                if is_decorative_new_badge(worksheet_image.image_bytes):
                    continue
                candidates.append(
                    ImageExportCandidate(
                        output_row=exported.output_row,
                        code=exported.product.code,
                        image_bytes=worksheet_image.image_bytes,
                        extension=image_extension(worksheet_image.image),
                    )
                )
        return candidates
    finally:
        workbook.close()


def prepare_image_write_tasks(
    output_path: Path,
    candidates: list[ImageExportCandidate],
    allow_multiple_images_per_product: bool = False,
) -> list[ImageWriteTask]:
    output_dir = images_output_dir_for(output_path)
    filename_counts: dict[str, int] = {}
    selected_counts_by_code: dict[str, int] = {}
    image_hashes_by_code: dict[str, set[str]] = {}
    tasks: list[ImageWriteTask] = []

    for candidate in sorted(candidates, key=lambda item: item.output_row):
        code_key = normalize_text(candidate.code)
        image_hash = hashlib.sha256(candidate.image_bytes).hexdigest()
        seen_hashes = image_hashes_by_code.setdefault(code_key, set())
        if image_hash in seen_hashes:
            continue
        if not allow_multiple_images_per_product and selected_counts_by_code.get(code_key, 0) >= 1:
            continue
        seen_hashes.add(image_hash)
        selected_counts_by_code[code_key] = selected_counts_by_code.get(code_key, 0) + 1

        base_name = safe_filename_part(candidate.code)
        filename_counts[base_name] = filename_counts.get(base_name, 0) + 1
        suffix = "" if filename_counts[base_name] == 1 else f"_{filename_counts[base_name]}"
        filename = f"{base_name}{suffix}{candidate.extension}"
        tasks.append(
            ImageWriteTask(
                output_row=candidate.output_row,
                image_bytes=candidate.image_bytes,
                target_path=output_dir / filename,
                relative_path=f"{output_dir.name}/{filename}",
            )
        )

    return tasks


def write_image_task(task: ImageWriteTask) -> tuple[int, str]:
    task.target_path.write_bytes(task.image_bytes)
    return task.output_row, task.relative_path


def export_product_images(
    source_path: Path,
    output_path: Path,
    exported_products: list[ExportedProduct],
    progress_callback: ProgressCallback | None = None,
    allow_multiple_images_per_product: bool = False,
) -> tuple[int, Path | None, dict[int, str]]:
    products_by_sheet: dict[str, list[ExportedProduct]] = {}
    for exported in exported_products:
        if not exported.product.code:
            continue
        products_by_sheet.setdefault(normalize_key(exported.product.source_sheet), []).append(exported)

    if not products_by_sheet:
        if progress_callback:
            progress_callback(0, 0)
        return 0, None, {}

    image_candidates: list[ImageExportCandidate] = []
    max_workers = min(IMAGE_WORKER_COUNT, len(products_by_sheet))
    total_sheets = len(products_by_sheet)
    if progress_callback:
        progress_callback(0, -total_sheets)
    with ThreadPoolExecutor(max_workers=max_workers) as executor:
        futures = [
            executor.submit(analyze_sheet_image_exports, source_path, sheet_products[0].product.source_sheet, sheet_products)
            for sheet_products in products_by_sheet.values()
        ]
        for processed_count, future in enumerate(as_completed(futures), start=1):
            image_candidates.extend(future.result())
            if progress_callback:
                progress_callback(processed_count, -total_sheets)

    image_tasks = prepare_image_write_tasks(output_path, image_candidates, allow_multiple_images_per_product)
    total_images = len(image_tasks)
    if progress_callback:
        progress_callback(0, total_images)
    if not image_tasks:
        return 0, None, {}

    output_dir = images_output_dir_for(output_path)
    prepare_images_output_dir(output_dir)

    written = 0
    image_paths_by_output_row: dict[int, str] = {}
    write_results: dict[int, tuple[int, str]] = {}
    with ThreadPoolExecutor(max_workers=min(IMAGE_WORKER_COUNT, len(image_tasks))) as executor:
        futures = {executor.submit(write_image_task, task): index for index, task in enumerate(image_tasks)}
        for processed_count, future in enumerate(as_completed(futures), start=1):
            task_index = futures[future]
            output_row, image_path = future.result()
            write_results[task_index] = (output_row, image_path)
            written += 1
            if progress_callback:
                progress_callback(processed_count, total_images)

    for task_index in sorted(write_results):
        output_row, image_path = write_results[task_index]
        image_paths_by_output_row.setdefault(output_row, image_path)

    return written, output_dir if written else None, image_paths_by_output_row


def convert(
    source_path: Path,
    output_path: Path,
    aliases_path: Path | None,
    schema_path: Path | None = None,
    issue_handler: IssueHandler | None = None,
    group_configs: list[GroupConfig] | None = None,
    group_cache_path: Path | None = None,
    duplicate_handler: DuplicateChoiceHandler | None = None,
    product_issue_handler: ProductIssueHandler | None = None,
    progress_callback: ProgressCallback | None = None,
    image_progress_callback: ProgressCallback | None = None,
    allow_multiple_images_per_product: bool = False,
) -> ConversionStats:
    products, skipped_without_code, skipped_without_price, skipped_invalid_code = read_source_products(source_path, issue_handler, product_issue_handler)
    products = deduplicate_products(products, duplicate_handler)
    total_products = len(products)
    if progress_callback:
        progress_callback(0, total_products)
    schema = read_schema(schema_path or default_schema_path())
    groups = groups_from_schema(schema)
    aliases = load_aliases(aliases_path)
    config_by_sheet = {normalize_key(config.source_sheet): config for config in group_configs or []}
    missing_group_ids: dict[str, int] = {}
    skipped_without_group_id: dict[str, int] = {}

    workbook, ws, indexes = create_output_workbook(schema)
    append_configured_groups(workbook, group_configs or [])
    max_columns = len(schema["product_headers"])
    converted_count = 0
    exported_products: list[ExportedProduct] = []

    for processed_count, product in enumerate(products, start=1):
        target_row = product_to_target_row(product, max_columns, indexes, groups, aliases, config_by_sheet)
        group_id = target_row[indexes["Номер_группы"] - 1]
        group_name = target_row[indexes["Название_группы"] - 1]
        if not group_id:
            missing_group_ids[str(group_name)] = missing_group_ids.get(str(group_name), 0) + 1
            skipped_without_group_id[str(group_name)] = skipped_without_group_id.get(str(group_name), 0) + 1
            if progress_callback:
                progress_callback(processed_count, total_products)
            continue
        ws.append(target_row)
        converted_count += 1
        exported_products.append(ExportedProduct(product=product, output_row=ws.max_row))
        if progress_callback:
            progress_callback(processed_count, total_products)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    exported_images_count, images_output_dir, image_paths_by_output_row = export_product_images(
        source_path,
        output_path,
        exported_products,
        image_progress_callback,
        allow_multiple_images_per_product,
    )
    image_column = indexes.get("Ссылка_изображения")
    if image_column is not None:
        for output_row, image_path in image_paths_by_output_row.items():
            ws.cell(row=output_row, column=image_column, value=image_path)
    workbook.save(output_path)
    workbook.close()
    if group_configs is not None:
        save_group_cache(group_configs, group_cache_path)
    return ConversionStats(
        converted_count=converted_count,
        missing_group_ids=missing_group_ids,
        skipped_without_code=skipped_without_code,
        skipped_without_price=skipped_without_price,
        skipped_invalid_code=skipped_invalid_code,
        skipped_without_group_id=skipped_without_group_id,
        exported_images_count=exported_images_count,
        images_output_dir=str(images_output_dir) if images_output_dir else None,
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Convert a 1C/price-list XLSX workbook into a satu.kz export/import XLSX workbook."
    )
    parser.add_argument("source", type=Path, help="Input price-list workbook.")
    parser.add_argument("-o", "--output", type=Path, default=Path("output/satu-products.xlsx"), help="Output XLSX path.")
    parser.add_argument(
        "--aliases",
        type=Path,
        default=None,
        help="JSON file mapping source sheet names to target satu.kz group names.",
    )
    parser.add_argument(
        "--schema",
        type=Path,
        default=None,
        help="JSON file with satu.kz product columns and groups.",
    )
    parser.add_argument(
        "--include-without-code",
        action="store_true",
        help="Include source rows that have a product name and price but no code.",
    )
    parser.add_argument(
        "--allow-multiple-images-per-product",
        action="store_true",
        help="Export several different images for one product code using _2, _3 suffixes.",
    )
    return parser


def main() -> None:
    args = build_parser().parse_args()
    aliases = args.aliases or default_aliases_path()
    handler = (lambda issue: True) if args.include_without_code else None
    try:
        stats = convert(
            args.source,
            args.output,
            aliases,
            args.schema,
            handler,
            allow_multiple_images_per_product=args.allow_multiple_images_per_product,
        )
    except ValueError as exc:
        print(f"Error: {exc}", file=sys.stderr)
        raise SystemExit(1) from None
    print(f"Converted products: {stats.converted_count}")
    print(f"Output: {args.output.resolve()}")
    if stats.skipped_without_code:
        print("Skipped rows without code:")
        for sheet_name, count in sorted(stats.skipped_without_code.items()):
            print(f"  - {sheet_name}: {count}")
    if stats.skipped_without_price:
        print("Skipped rows without price or with zero price:")
        for sheet_name, count in sorted(stats.skipped_without_price.items()):
            print(f"  - {sheet_name}: {count}")
    if stats.skipped_invalid_code:
        print("Skipped rows with invalid product code:")
        for sheet_name, count in sorted(stats.skipped_invalid_code.items()):
            print(f"  - {sheet_name}: {count}")
    if stats.missing_group_ids:
        print("Skipped products from groups without satu.kz group number:")
        for group_name, count in sorted(stats.missing_group_ids.items()):
            print(f"  - {group_name}: {count}")
    if stats.exported_images_count:
        print(f"Exported product images: {stats.exported_images_count}")
        print(f"Images output directory: {stats.images_output_dir}")


if __name__ == "__main__":
    main()
