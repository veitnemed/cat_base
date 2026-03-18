import json
import os
import sys
from datetime import datetime, timedelta
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.style import Style

# Для работы с PNG-иконками
from PIL import Image, ImageTk

# БЛОК РУЧНОЙ НАСТРОЙКИ
BG_MAIN = BG_WORK_AREA = BG_HEADER = BG_COUNTER = "#ededed"
BG_ENTRY, BG_SETTINGS = "#ffffff", "#ededed"
TEXT_MAIN = TEXT_LABEL = TEXT_HEADER = TEXT_FACTORY = "#212529"
TEXT_BUTTON, TEXT_DATE, TEXT_COUNTER = "#ffffff", "#555555", "#25405c"
BORDER_MAIN = BORDER_INPUT = BORDER_WORK_AREA = "#212529"
BUTTON_SERP = BUTTON_SAVE = BUTTON_SEARCH = "#505b65"
BUTTON_SUCCESS, BUTTON_DANGER = "#dde6dd", "#d5c0c0"
BUTTON_SECONDARY, BUTTON_DRAFT, BUTTON_NOT_READY = "#6c757d", "#ce9713", "#972f2f"

FONT_SIZE_SMALL, FONT_SIZE_MEDIUM, FONT_SIZE_LARGE = 10, 11, 12
FONT_SIZE_XL = FONT_SIZE_XXL = 12
FONT_PRIMARY = ("Arial", FONT_SIZE_MEDIUM)
FONT_BOLD = ("Arial", FONT_SIZE_MEDIUM, "bold")
FONT_COUNTER, FONT_BLOCK = ("Arial", 12), ("Arial", 12)
FONT_SERP_NUMBER = FONT_SAVE_BUTTON = ("Arial", 12)
FONT_HEADER1 = FONT_HEADER2 = ("Arial", FONT_SIZE_XL, "bold")
FONT_HEADER3, FONT_FACTORY = ("Arial", 11, "italic"), ("Arial", FONT_SIZE_LARGE, "bold")
FONT_SETTINGS, FONT_BUTTON = ("Arial", FONT_SIZE_XL, "bold"), (
    "Arial",
    FONT_SIZE_MEDIUM,
)
FONT_STATUS_BUTTON = ("Arial", 11, "bold")

blockcount_mode = False
storage_count_lbl = None
request_table = None
filter_var = None
work_header_label = None
storage_header_label = None
manual_override_ping = False
last_window_state = None
request_filter = "Все"
TEXT_ADD_SERP, TEXT_SAVE, TEXT_SEARCH = "+", "Сохранить", "Поиск"
TEXT_WORK_TITLE, TEXT_STORAGE_TITLE = "Изделия в работе", "Ждут отправки на склад"
TEXT_WORK_LABEL, TEXT_STORAGE_LABEL = "В работе", "Собранно"
TEXT_YEAR, TEXT_NUMBER, TEXT_SEARCH_LABEL = "Год:", "Номер:", "Поиск:"

storage_count_cache = 0
BORDER_WIDTH, ENTRY_BORDER_WIDTH = 10, 10
ROW_HEIGHT, PADY_HEADER, PADY_PRODUCT = 40, 5, 3
repair_blocks = set()
REPAIR_TEXT = "РЕМОНТ"
REPAIR_BOOTSTYLE = "secondary"

# КОД ПРИЛОЖЕНИЯ
base_path = (
    os.path.dirname(sys.executable)
    if getattr(sys, "frozen", False)
    else os.path.dirname(os.path.abspath(__file__))
)
DATA_FILE = os.path.join(base_path, "data.json")
BACKUP_DIR = os.path.join(os.path.expanduser("~"), "Serp_Backup")
BACKUP_FILE = os.path.join(BACKUP_DIR, "data_backup.json")

MONTHS_GENITIVE = [
    "января",
    "февраля",
    "марта",
    "апреля",
    "мая",
    "июня",
    "июля",
    "августа",
    "сентября",
    "октября",
    "ноября",
    "декабря",
]

storage_visible = True
marked_statuses = {}
show_draft_group = False
request_filter = "Все"
up112_hints = set()
sort_buttons = {}

# ===== Новое: подсветка ССБ 112 из внешнего файла =====
highlight_112 = (
    set()
)  # множество ключей (номер, завод), для которых ССБ 112 нужно окрасить в info
UP112_PATH = r"C:\ssb_data\global_data\up_112"


# ======================================================
def on_window_state_change(event=None):
    """Авто-раскрытие при zoomed и авто-сворачивание при возврате в normal."""
    global last_window_state, storage_visible
    try:
        state = root.state()  # 'normal' | 'zoomed' | 'iconic'
        if state != last_window_state:
            if state == "zoomed":
                if not storage_visible:
                    toggle_storage_visibility()  # раскрыть правую колонку
            elif state == "normal":
                if storage_visible:
                    toggle_storage_visibility()  # свернуть правую колонку
            last_window_state = state
    except Exception as e:
        print("on_window_state_change() error:", e)


def resource_path(relative_path):
    return (
        os.path.join(sys._MEIPASS, relative_path)
        if hasattr(sys, "_MEIPASS")
        else os.path.join(os.path.abspath("."), relative_path)
    )


if not os.path.exists(DATA_FILE):
    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(
            {
                "products": {},
                "assembled": [],
                "storage": [],
                "comments": {},
                "assembly_dates": {},
                "storage_dates": {},
                "draft_products": {},
                "redy_products": {},
                "settings": {
                    "factory_order": [
                        "ВЕКТОР",
                        "ИНТЕГРАЛ",
                        "РЗП",
                        "КНИИТМУ",
                        "СВТ",
                        "СИГНАЛ",
                    ],
                    "sort_order": "new_first",
                    "archive_view_mode": "journal",
                    "data_path": base_path,
                    "scaling_factor": 2.0,
                    "show_draft_group": False,
                },
            },
            f,
        )

products, assembled_products, storage_products = {}, set(), set()
draft_products, redy_products, assembly_dates, storage_dates = set(), set(), {}, {}
assembly_years, comments, product_dates = {}, {}, {}
factory_order, sort_order, archive_view_mode = (
    ["ВЕКТОР", "ИНТЕГРАЛ", "РЗП", "КНИИТМУ", "СВТ", "СИГНАЛ"],
    "new_first",
    "journal",
)
data_path, work_scroll_position, storage_scroll_position = base_path, 0, 0
archive_scroll_position, work_top_element, storage_top_element = 0, None, None
archive_top_element, long_press_timers = None, {}
work_widgets, storage_widgets, header_widgets, date_header_widgets = {}, {}, {}, {}

scaling_factor = 2.0

factory_mapping = {
    "1": "ВЕКТОР",
    "2": "ИНТЕГРАЛ",
    "3": "РЗП",
    "4": "КНИИТМУ",
    "5": "СВТ",
    "6": "СИГНАЛ",
}
factory_reverse_mapping = {v: k for k, v in factory_mapping.items()}
type_mapping, all_block_types = {"61": "ССБ 161", "12": "ССБ 112", "14": "ССБ 114"}, [
    "ССБ 161",
    "ССБ 112",
    "ССБ 114",
]

BLOCK_SYNONYMS = {
    "ССБ 161": ["ССБ161", "161", "ССБ-161", "SSB161"],
    "ССБ 112": ["ССБ112", "112", "ССБ-112", "SSB112"],
    "ССБ 114": ["ССБ114", "114", "ССБ-114", "SSB114"],
}


def ensure_block_dict(key) -> None:
    """
    Гарантирует, что products[key] — словарь с каноническими ключами из all_block_types.
    Переносит значения из синонимов и заполняет отсутствующие False.
    """
    blk = products.get(key)
    if not isinstance(blk, dict):
        products[key] = {t: False for t in all_block_types}
        return

    # Переносим значения из синонимов в канонические ключи
    for canon, syns in BLOCK_SYNONYMS.items():
        if canon not in blk:
            for s in syns:
                if s in blk:
                    blk[canon] = bool(blk.pop(s))
                    break

    # Гарантируем наличие всех канонических ключей
    for t in all_block_types:
        blk.setdefault(t, False)


def normalize_all_products() -> None:
    """
    Прогоняет ensure_block_dict по всем изделиям. Если схема менялась — сохраняет файл.
    """
    changed_schema = False
    for k in list(products.keys()):
        blk = products.get(k)
        before = set(blk.keys()) if isinstance(blk, dict) else set()
        ensure_block_dict(k)
        after = set(products[k].keys()) if isinstance(products[k], dict) else set()
        if before != after:
            changed_schema = True
    if changed_schema:
        save_data()


def set_sort_mode(mode: str):
    """Единая точка переключения режима левой колонки."""
    global work_sort_mode, show_draft_group, blockcount_mode
    mode = mode.lower()
    if mode not in ("drafts", "blocks", "factories"):
        return
    work_sort_mode = mode

    # Приводим существующие флаги к выбранному режиму
    if mode == "drafts":
        show_draft_group = True
        blockcount_mode = False
    elif mode == "blocks":
        show_draft_group = False
        blockcount_mode = True
    else:  # "factories"
        show_draft_group = False
        blockcount_mode = False

    _refresh_sort_buttons()
    update_product_list()  # пересобрать список слева


def _refresh_sort_buttons():
    """Подсветка активной кнопки."""
    try:
        btn_mode_drafts.configure(
            bootstyle="primary" if work_sort_mode == "drafts" else "secondary"
        )
        btn_mode_blocks.configure(
            bootstyle="primary" if work_sort_mode == "blocks" else "secondary"
        )
        btn_mode_factories.configure(
            bootstyle="primary" if work_sort_mode == "factories" else "secondary"
        )
    except Exception:
        pass


def parse_serial_number(serial):
    if not serial.isdigit() or len(serial) != 9:
        return None
    code, item_number, factory_code, year_code = (
        serial[:2],
        serial[-4:],
        serial[-5],
        serial[-7:-5],
    )
    if code not in type_mapping or factory_code not in factory_mapping:
        return None
    return {
        "Тип изделия": type_mapping[code],
        "Год выпуска": year_code,
        "Завод": factory_mapping[factory_code],
        "Номер изделия": item_number,
        "Ключ изделия": (item_number, factory_mapping[factory_code]),
        "Серийный номер": serial,
    }


def save_data(show_message=False):
    global data_path, scaling_factor, show_draft_group, up112_hints
    settings = {
        "factory_order": factory_order,
        "sort_order": sort_order,
        "archive_view_mode": archive_view_mode,
        "data_path": data_path,
        "scaling_factor": scaling_factor,
        "show_draft_group": show_draft_group,
    }
    data = {
        "products": {json.dumps(k): v for k, v in products.items()},
        "assembled": list(assembled_products),
        "storage": list(storage_products),
        "draft": list(draft_products),
        "redy": list(redy_products),
        "comments": {json.dumps(k): v for k, v in comments.items()},
        "product_dates": {json.dumps(k): v for k, v in product_dates.items()},
        "assembly_dates": {json.dumps(k): v for k, v in assembly_dates.items()},
        "storage_dates": {json.dumps(k): v for k, v in storage_dates.items()},
        "assembly_years": {json.dumps(k): v for k, v in assembly_years.items()},
        "marked_statuses": {json.dumps(k): v for k, v in marked_statuses.items()},
        # ↓↓↓ Сохраняем подсказки для 112
        "up112_hints": list(up112_hints),
        "repair_blocks": [[json.dumps(k), b] for (k, b) in repair_blocks],
        "settings": settings,
    }

    DATA_FILE = os.path.join(data_path, "data.json")
    BACKUP_DIR = os.path.join(data_path, "Serp_Backup")
    BACKUP_FILE = os.path.join(BACKUP_DIR, "data_backup.json")

    with open(DATA_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    try:
        os.makedirs(BACKUP_DIR, exist_ok=True)
        with open(BACKUP_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception as e:
        print("Ошибка бэкапа:", e)

    if show_message:
        messagebox.showinfo("Сообщение", "Данные сохранены!")


def load_data():
    global products, assembled_products, storage_products, comments, product_dates, assembly_dates
    global storage_dates, assembly_years, draft_products, redy_products, factory_order, marked_statuses
    global sort_order, archive_view_mode, data_path, scaling_factor, show_draft_group, up112_hints
    global repair_blocks

    if not os.path.exists(DATA_FILE):
        return

    # 1) читаем файл
    with open(DATA_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    # 2) разворачиваем структуры
    products = {tuple(json.loads(k)): v for k, v in data.get("products", {}).items()}
    assembled_products = set(tuple(x) for x in data.get("assembled", []))
    storage_products = set(tuple(x) for x in data.get("storage", []))
    draft_products = set(tuple(x) for x in data.get("draft", []))
    redy_products = set(tuple(x) for x in data.get("redy", []))
    comments = {tuple(json.loads(k)): v for k, v in data.get("comments", {}).items()}
    product_dates = {
        tuple(json.loads(k)): v for k, v in data.get("product_dates", {}).items()
    }
    assembly_dates = {
        tuple(json.loads(k)): v for k, v in data.get("assembly_dates", {}).items()
    }
    storage_dates = {
        tuple(json.loads(k)): v for k, v in data.get("storage_dates", {}).items()
    }
    assembly_years = {
        tuple(json.loads(k)): v for k, v in data.get("assembly_years", {}).items()
    }
    marked_statuses = {
        tuple(json.loads(k)): v for k, v in data.get("marked_statuses", {}).items()
    }
    up112_hints = set(tuple(x) for x in data.get("up112_hints", []))

    # пары вида ((номер, завод), "ССБ 1xx")
    repair_blocks = set(
        (tuple(json.loads(k)), b) for (k, b) in data.get("repair_blocks", [])
    )

    # 3) миграция 'КНИТМУ' -> 'КНИИТМУ' для всех коллекций, включая repair_blocks/marked_statuses/up112_hints
    for key in list(products.keys()):
        if "КНИТМУ" in key:
            new_key = (key[0], "КНИИТМУ")

            # dict-ы по ключу
            products[new_key] = products.pop(key)
            if key in comments:
                comments[new_key] = comments.pop(key)
            if key in product_dates:
                product_dates[new_key] = product_dates.pop(key)
            if key in assembly_dates:
                assembly_dates[new_key] = assembly_dates.pop(key)
            if key in storage_dates:
                storage_dates[new_key] = storage_dates.pop(key)
            if key in assembly_years:
                assembly_years[new_key] = assembly_years.pop(key)
            if key in marked_statuses:
                marked_statuses[new_key] = marked_statuses.pop(key)

            # множества по ключу
            if key in draft_products:
                draft_products.remove(key)
                draft_products.add(new_key)
            if key in redy_products:
                redy_products.remove(key)
                redy_products.add(new_key)
            if key in assembled_products:
                assembled_products.remove(key)
                assembled_products.add(new_key)
            if key in storage_products:
                storage_products.remove(key)
                storage_products.add(new_key)
            if key in up112_hints:
                up112_hints.remove(key)
                up112_hints.add(new_key)

            # пары (key, block_type) в repair_blocks
            for blk in ["ССБ 112", "ССБ 114", "ССБ 161"]:
                if (key, blk) in repair_blocks:
                    repair_blocks.remove((key, blk))
                    repair_blocks.add((new_key, blk))

    # 4) настройки
    settings = data.get("settings", {})
    factory_order = settings.get(
        "factory_order", ["ВЕКТОР", "ИНТЕГРАЛ", "РЗП", "КНИИТМУ", "СВТ", "СИГНАЛ"]
    )
    sort_order = settings.get("sort_order", "new_first")
    archive_view_mode = settings.get("archive_view_mode", "journal")
    data_path = settings.get("data_path", base_path)
    scaling_factor = settings.get("scaling_factor", 2.0)
    show_draft_group = settings.get("show_draft_group", False)
    # 5) миграция: для старых записей в архиве добавить 'дату сборки' = 'дата отправки'
    changed = False
    for key in assembled_products:
        if key not in storage_dates and key in assembly_dates:
            storage_dates[key] = assembly_dates[key]
            changed = True
    if changed:
        save_data()
    normalize_all_products()


def process_serial():
    """
    Поддерживает 2 формата ввода в поле 'Номер':
      1) ОДИН номер:   <номер><код_завода>
         примеры: '4516', '12', '999991' — последняя цифра всегда код завода (1..6).
      2) ДИАПАЗОН:     <начало>-<конец><код_завода>
         пример: '45-506'  -> добавит 0045..0050, завод по коду '6' => СИГНАЛ.
         Ввод без ведущих нулей; мы сами дополним до 4 знаков.

    Проверки:
      - корректность формата и кода завода,
      - валидность года (2 цифры),
      - границы 0..9999, конец >= начало,
      - ограничение по размеру диапазона,
      - отсутствие дублей: если хоть один СЕРП из диапазона уже существует/в складе/в архиве — НИЧЕГО не добавляем и показываем список конфликтов.
    """
    MAX_BATCH = 300  # разумный предел за один раз, при необходимости увеличь

    raw = entry_var.get().strip()
    year = year_var.get().strip()

    # год обязателен (ровно 2 цифры)
    if not (year.isdigit() and len(year) == 2):
        messagebox.showerror("Ошибка", "Некорректный год (нужно ровно 2 цифры).")
        return

    # ===== режим диапазона: <start>-<end><factory> =====
    if "-" in raw:
        s = raw.replace(" ", "")
        try:
            left, right = s.split("-", 1)
        except ValueError:
            messagebox.showerror(
                "Ошибка", "Неверный формат диапазона. Пример: 45-506 (6 — код завода)."
            )
            return

        if not (left.isdigit() and right.isdigit()):
            messagebox.showerror(
                "Ошибка", "Диапазон должен содержать только цифры. Пример: 45-506."
            )
            return

        if len(right) < 2:
            messagebox.showerror(
                "Ошибка",
                "Во второй части нужна цифра завода в самом конце. Пример: 45-506 (6 — код завода).",
            )
            return

        start_num = int(left)
        factory_code = right[-1]
        end_num_str = right[:-1]

        if factory_code not in factory_mapping:
            messagebox.showerror(
                "Ошибка", "Код завода должен быть 1–6 (последняя цифра)."
            )
            return
        if end_num_str == "":
            messagebox.showerror(
                "Ошибка", "После тире должен быть номер и код завода. Пример: 45-506."
            )
            return

        end_num = int(end_num_str)

        # валидации границ
        if not (0 <= start_num <= 9999 and 0 <= end_num <= 9999):
            messagebox.showerror(
                "Ошибка", "Номера в диапазоне должны быть от 0 до 9999."
            )
            return
        if end_num < start_num:
            messagebox.showerror("Ошибка", "Конец диапазона меньше начала.")
            return

        count = end_num - start_num + 1
        if count > MAX_BATCH:
            messagebox.showerror(
                "Слишком много",
                f"Слишком длинный диапазон ({count}). Максимум за раз — {MAX_BATCH}.",
            )
            return

        factory_name = factory_mapping[factory_code]

        # проверка на конфликты
        conflicts = []  # список текстов вида '0045 (в работе)' и т.п.
        for n in range(start_num, end_num + 1):
            serial_number = f"{n:04d}"
            key = (serial_number, factory_name)

            if key in assembled_products:
                conflicts.append(f"СЕРП {serial_number} — уже в архиве")
            elif key in storage_products:
                conflicts.append(
                    f"СЕРП {serial_number} — собран и ждёт отправки на склад"
                )
            elif key in products:
                # уже есть в базе «в работе»
                conflicts.append(f"СЕРП {serial_number} — уже добавлен")

        if conflicts:
            # ничего не добавляем — показываем полный список проблем
            lines = "\n".join(conflicts[:50])
            more = "" if len(conflicts) <= 50 else f"\n…и ещё {len(conflicts) - 50}"
            messagebox.showerror(
                "Дубликаты в диапазоне",
                f"Ничего не добавлено. В диапазоне уже существуют:\n{lines}{more}",
            )
            return

        # добавляем все разом
        for n in range(start_num, end_num + 1):
            serial_number = f"{n:04d}"
            key = (serial_number, factory_name)
            products[key] = {t: False for t in all_block_types}
            product_dates[key] = datetime.now().strftime("%d.%m")
            assembly_years[key] = year
            # перестраховка: чистим статусы, если вдруг где-то были
            storage_products.discard(key)
            assembled_products.discard(key)
            draft_products.discard(key)
            redy_products.discard(key)

        save_data()
        update_product_list()
        entry_var.set("")
        messagebox.showinfo(
            "Готово",
            f"Добавлено изделий: {count}\n"
            f"{start_num:04d}–{end_num:04d} ({factory_name})",
        )
        return

    # ===== одиночный ввод (как раньше) =====
    raw_number = raw
    if not raw_number.isdigit() or len(raw_number) < 2 or int(raw_number) > 99999:
        messagebox.showerror(
            "Ошибка", "Введите число от 2 до 99999 (номер+код завода в конце)."
        )
        return

    factory_code = raw_number[-1]
    if factory_code not in factory_mapping:
        messagebox.showerror("Ошибка", "Код завода должен быть 1–6 (последняя цифра).")
        return

    serial_number = raw_number[:-1].zfill(4)
    factory_name = factory_mapping[factory_code]
    key = (serial_number, factory_name)

    if key in assembled_products:
        messagebox.showwarning("Ошибка", "СЕРП с таким номером уже собран (в архиве)!")
        entry_var.set("")
        return

    if key in storage_products:
        messagebox.showwarning(
            "Ошибка", "СЕРП с таким номером уже в правой колонке (собран)!"
        )
        entry_var.set("")
        return

    if key in products:
        messagebox.showwarning(
            "Ошибка", f"СЕРП с таким номером уже существует! Ключ: {key}"
        )
        entry_var.set("")
        return

    # всё ок — добавляем один
    products[key] = {t: False for t in all_block_types}
    product_dates[key] = datetime.now().strftime("%d.%m")
    assembly_years[key] = year
    storage_products.discard(key)
    assembled_products.discard(key)
    draft_products.discard(key)
    redy_products.discard(key)

    save_data()
    update_product_list()
    entry_var.set("")


def toggle_block(key, block_type):
    ensure_block_dict(key)
    products[key][block_type] = not bool(products[key].get(block_type, False))
    if block_type == "ССБ 112" and products[key][block_type]:
        up112_hints.discard(key)

    save_data()
    update_row_widgets(key)


def on_block_left_click(key, block_type):
    """
    ЛКМ по кнопке блока:
    - если блок в режиме РЕМОНТ -> снять ремонт (без изменения установлен/не установлен)
    - иначе -> как раньше: toggle установлен/не установлен
    """
    if (key, block_type) in repair_blocks:
        repair_blocks.discard((key, block_type))
        save_data()
        update_row_widgets(key)
    else:
        toggle_block(key, block_type)


def toggle_block_repair(key, block_type):
    """
    ПКМ: переключить РЕМОНТ только для СЕРЫХ кнопок.
    - если кнопка зелёная (установлено) — игнорируем
    - если кнопка синяя (подсказка 112) — игнорируем
    - только серую можно отдать в ремонт
    """
    # установлен?
    installed = bool(products.get(key, {}).get(block_type, False))
    if installed:
        return

    # синяя подсказка допустима только для 112, её в ремонт не отдаём
    info_tint = (block_type == "ССБ 112") and (key in up112_hints)
    if info_tint:
        return

    # ок — именно СЕРАЯ: переключаем ремонт
    if (key, block_type) in repair_blocks:
        repair_blocks.discard((key, block_type))
    else:
        repair_blocks.add((key, block_type))

    save_data()
    update_row_widgets(key)


def delete_serp(key):
    if messagebox.askyesno("Подтверждение", f"Удалить СЕРП №{key[0]} ({key[1]})?"):
        save_scroll_positions()  # ← сохранить
        products.pop(key, None)
        comments.pop(key, None)
        product_dates.pop(key, None)
        draft_products.discard(key)
        redy_products.discard(key)
        save_data()
        update_product_list()
        restore_scroll_positions()  # ← восстановить


def save_all_comments():
    for area in [work_frame, storage_frame]:
        for widget in area.winfo_children():
            if isinstance(widget, tb.Frame) and hasattr(widget, "key"):
                for child in widget.winfo_children():
                    if isinstance(child, tb.Entry) and hasattr(child, "key_reference"):
                        comments[child.key_reference] = child.get()
    save_data(False)


def mark_storage(key):
    save_all_comments()
    save_scroll_positions()

    # Переносим изделие в правую колонку «собрано»
    storage_products.add(key)
    storage_dates[key] = datetime.now().strftime("%d.%m.%y")
    assembled_products.discard(key)  # точно не в архиве

    ensure_mutual_state(key)
    save_data()

    # Актуализировать оба экрана
    update_product_list(preserve_scroll=True)
    update_assembly_archive()
    restore_scroll_positions()


def mark_assembled(key):
    save_all_comments()
    save_scroll_positions()

    # Отдать на склад (архив)
    assembled_products.add(key)
    assembly_dates[key] = datetime.now().strftime("%d.%m.%y")

    # Из правой колонки убрать
    storage_products.discard(key)

    ensure_mutual_state(key)
    save_data()

    update_assembly_archive()
    update_product_list(preserve_scroll=True)
    restore_scroll_positions()


def return_to_work(key):
    save_all_comments()
    save_scroll_positions()
    if key in storage_products:
        was_draft = key in draft_products
        storage_products.remove(key)
        storage_dates.pop(key, None)
        if was_draft and key not in draft_products:
            draft_products.add(key)
    save_data()
    update_product_list(preserve_scroll=True)
    restore_scroll_positions()


def toggle_draft_status(key):
    save_all_comments()
    if key in draft_products:
        draft_products.remove(key)
    else:
        draft_products.add(key)
    save_data()
    update_row_widgets(key)


def back_to_draft(key):
    save_all_comments()
    if key in redy_products:
        redy_products.remove(key)
        draft_products.add(key)
    save_data()
    update_row_widgets(key)


def toggle_marked_status(key):
    # Ремонт больше не вешается на статус; ремонт ставится на кнопках блоков ПКМ.
    return


def format_date_genitive(date_str):
    """Форматирует дату из 'дд.мм' или 'дд.мм.гг' в 'дд месяц_родительный'"""
    if not date_str or not isinstance(date_str, str):
        return date_str
    date_str = date_str.rstrip(".")
    parts = date_str.split(".")
    if len(parts) >= 2:
        try:
            day = int(parts[0])
            month = int(parts[1])
            if 1 <= month <= 12:
                return f"{day} {MONTHS_GENITIVE[month - 1]}"
        except ValueError:
            pass
    return date_str


def send_all_to_storage():
    if not storage_products:
        messagebox.showinfo("Информация", "Нет изделий для отправки на склад")
        return
    if messagebox.askyesno(
        "Подтверждение", "Вы точно хотите отдать все СЕРПЫ на склад?"
    ):
        for key in list(storage_products):
            mark_assembled(key)
        messagebox.showinfo("Успех", "Все изделия отправлены на склад")


def parse_date_str(date_str, key=None):
    try:
        parts = date_str.split(".")
        if len(parts) == 2:
            day, month = parts
            if key and key in assembly_years:
                year = (
                    int("20" + assembly_years[key])
                    if len(assembly_years[key]) == 2
                    else int(assembly_years[key])
                )
            else:
                year = datetime.now().year
            return datetime(year, int(month), int(day))
        elif len(parts) == 3:
            day, month, year = parts
            year_full = int("20" + year) if len(year) == 2 else int(year)
            return datetime(year_full, int(month), int(day))
        else:
            return datetime.min
    except Exception:
        return datetime.min


def format_short_serp(key):
    number, factory_name = key
    factory_code = factory_reverse_mapping.get(factory_name, "?")
    return f"СЕРП {number} [{factory_code}]"


def create_product_row(
    parent, key, blocks, current_comments, in_storage, show_factory_code=False
):
    # подстраховка
    if key not in products:
        products[key] = {t: False for t in all_block_types}

    row = tb.Frame(parent, padding=5, style="Main.TFrame")
    row.pack(fill=X, pady=PADY_PRODUCT)
    row.key = key

    # ---------- правая колонка (собранные) ----------
    if in_storage:
        label_text = f"СЕРП №{key[0]}"
        if show_factory_code:
            factory_code = factory_reverse_mapping.get(key[1], "?")
            label_text += f" [{factory_code}]"

        tb.Label(row, text=label_text, font=FONT_SERP_NUMBER, style="Main.TLabel").pack(
            side=LEFT, padx=5
        )

        comment_frame = tb.Frame(row)
        comment_frame.pack(side=LEFT, padx=5, fill=X, expand=True)
        comment_entry = tb.Entry(
            comment_frame, width=15, font=FONT_PRIMARY, bootstyle="light"
        )
        comment_entry.pack(fill=BOTH, expand=True)
        comment_entry.key_reference = key
        comment_entry.insert(0, current_comments.get(key, comments.get(key, "")))
        comment_entry.bind(
            "<FocusOut>", lambda e, k=key: comments.update({k: comment_entry.get()})
        )
        comment_entry.bind(
            "<Return>",
            lambda e, k=key: (comments.update({k: comment_entry.get()}), root.focus()),
        )

        tb.Button(
            row,
            text="В работу",
            width=10,
            bootstyle="info",
            command=lambda k=key: return_to_work(k),
        ).pack(side=LEFT, padx=2)
        tb.Button(
            row,
            text="Склад",
            width=8,
            bootstyle="secondary",
            command=lambda k=key: mark_assembled(k),
        ).pack(side=LEFT, padx=2)
        tb.Button(
            row,
            text="✕",
            width=3,
            bootstyle="secondary",
            command=lambda k=key: delete_serp(k),
        ).pack(side=LEFT, padx=2)
        return row

    # ---------- левая колонка ("в работе") ----------
    is_draft = key in draft_products
    status_text, status_style = (
        ("ЗАГОТОВКА", "warning") if is_draft else ("НЕ ГОТОВО", "danger")
    )
    status_btn = tb.Button(
        row,
        text=status_text,
        width=12,
        bootstyle=status_style,
        command=lambda k=key: toggle_draft_status(k),
    )
    status_btn.pack(side=LEFT, padx=(0, 5))

    code_text = (
        f" [{factory_reverse_mapping.get(key[1], '?')}]" if show_factory_code else ""
    )
    tb.Label(
        row,
        text=f"СЕРП №{key[0]}{code_text} [",
        font=FONT_SERP_NUMBER,
        style="Main.TLabel",
    ).pack(side=LEFT, padx=2)

    block_btns = []
    for block_type in all_block_types:
        installed = bool(products[key].get(block_type, False))
        in_repair = (key, block_type) in repair_blocks
        info_tint = (block_type == "ССБ 112") and (key in up112_hints) and not installed

        if in_repair:
            btn_text = REPAIR_TEXT
            btn_style = REPAIR_BOOTSTYLE
        else:
            btn_text = block_type
            btn_style = (
                "success" if installed else ("info" if info_tint else "secondary")
            )

        block_btn = tb.Button(
            row,
            text=btn_text,
            width=8,
            bootstyle=btn_style,
            command=lambda k=key, b=block_type: on_block_left_click(k, b),
        )
        block_btn.pack(side=LEFT, padx=1)

        # ПКМ — разрешаем только для серых (не установлен и без синей подсказки)
        if (not installed) and (not info_tint):
            block_btn.bind(
                "<Button-3>", lambda e, k=key, b=block_type: toggle_block_repair(k, b)
            )

        block_btns.append(block_btn)

    row.block_btns = block_btns
    tb.Label(row, text="]", font=FONT_SERP_NUMBER, style="Main.TLabel").pack(
        side=LEFT, padx=2
    )

    comment_frame = tb.Frame(row)
    comment_frame.pack(side=LEFT, padx=5, fill=X, expand=True)
    comment_entry = tb.Entry(
        comment_frame, width=15, font=FONT_PRIMARY, bootstyle="light"
    )
    comment_entry.pack(fill=BOTH, expand=True)
    comment_entry.key_reference = key
    comment_entry.insert(0, current_comments.get(key, comments.get(key, "")))
    comment_entry.bind(
        "<FocusOut>", lambda e, k=key: comments.update({k: comment_entry.get()})
    )
    comment_entry.bind(
        "<Return>",
        lambda e, k=key: (comments.update({k: comment_entry.get()}), root.focus()),
    )

    # Галочка
    ready = all_blocks_done(key)
    if ready:
        check_btn = tb.Button(
            row,
            text="✓",
            width=3,
            bootstyle="info",
            command=lambda k=key: mark_storage(k),
        )
    else:
        check_btn = tb.Button(
            row, text="✓", width=3, bootstyle="light", state="disabled"
        )
        check_btn.bind(
            "<ButtonPress-1>",
            lambda e, k=key, bbs=block_btns, cb=check_btn: start_long_press(
                e, k, bbs, cb
            ),
        )
        check_btn.bind("<ButtonRelease-1>", lambda e, k=key: cancel_long_press(e, k))
    check_btn.pack(side=LEFT, padx=2)
    row.check_btn = check_btn

    tb.Button(
        row,
        text="✕",
        width=3,
        bootstyle="secondary",
        command=lambda k=key: delete_serp(k),
    ).pack(side=LEFT, padx=2)

    return row


def start_long_press(event, key, block_btns, check_btn):
    global long_press_timers
    if key in long_press_timers:
        root.after_cancel(long_press_timers[key])
    long_press_timers[key] = root.after(
        750, lambda k=key, bbs=block_btns, cb=check_btn: complete_long_press(k, bbs, cb)
    )


def cancel_long_press(event, key):
    global long_press_timers
    if key in long_press_timers:
        root.after_cancel(long_press_timers[key])
        del long_press_timers[key]


def complete_long_press(key, block_btns, check_btn):
    # 1) Снять все «ремонты»
    clear_repair_for_key(key)

    # 2) Поставить все блоки как установленные
    if key not in products:
        products[key] = {t: False for t in all_block_types}
    for t in all_block_types:
        products[key][t] = True

    # 3) Убрать «подсветку-подсказку» 112, раз уж он установлен
    try:
        up112_hints.discard(key)
    except Exception:
        pass

    # 4) Обновить внешний вид кнопок блоков
    for i, btn in enumerate(block_btns):
        try:
            btn.configure(text=all_block_types[i], bootstyle="success")
        except Exception:
            pass

    # 5) Включить галочку
    try:
        check_btn.configure(
            bootstyle="info", state="normal", command=lambda k=key: mark_storage(k)
        )
        check_btn.unbind("<ButtonPress-1>")
        check_btn.unbind("<ButtonRelease-1>")
    except Exception:
        pass

    save_data()
    update_row_widgets(key)


def save_scroll_positions():
    global work_scroll_position, storage_scroll_position, work_top_element, storage_top_element
    if work_canvas.winfo_exists():
        work_scroll_position = work_canvas.yview()[0]
    if storage_canvas.winfo_exists():
        storage_scroll_position = storage_canvas.yview()[0]
    work_top_element = get_top_visible_element(work_frame)
    storage_top_element = get_top_visible_element(storage_frame)


def get_top_visible_element(parent_frame):
    if not parent_frame.winfo_children():
        return None
    canvas = (
        work_canvas
        if parent_frame == work_frame
        else storage_canvas if parent_frame == storage_frame else archive_canvas
    )
    y_top, y_bottom = canvas.canvasy(0), canvas.canvasy(canvas.winfo_height())
    for widget in parent_frame.winfo_children():
        if isinstance(widget, tb.Frame) and hasattr(widget, "key"):
            widget_y, widget_height = widget.winfo_y(), widget.winfo_height()
            if widget_y + widget_height > y_top and widget_y < y_bottom:
                return widget.key
    return None


def restore_scroll_positions():
    if work_canvas.winfo_exists():
        work_canvas.yview_moveto(work_scroll_position)
    if storage_canvas.winfo_exists():
        storage_canvas.yview_moveto(storage_scroll_position)
    if work_top_element:
        scroll_to_element(work_frame, work_top_element)
    if storage_top_element:
        scroll_to_element(storage_frame, storage_top_element)


def scroll_to_element(parent_frame, element_key):
    parent_frame.update_idletasks()
    for widget in parent_frame.winfo_children():
        if (
            isinstance(widget, tb.Frame)
            and hasattr(widget, "key")
            and widget.key == element_key
        ):
            canvas = (
                work_canvas
                if parent_frame == work_frame
                else storage_canvas if parent_frame == storage_frame else archive_canvas
            )
            canvas.yview_moveto(widget.winfo_y() / parent_frame.winfo_height())
            break


def update_row_widgets(key):
    if key not in work_widgets:
        return

    row = work_widgets[key]
    blocks = products.get(key, {t: False for t in all_block_types})

    # Кнопки блоков
    for i, block_type in enumerate(all_block_types):
        installed = bool(blocks.get(block_type, False))
        in_repair = (key, block_type) in repair_blocks

        # определяем, будет ли синяя подсказка
        info_tint = (block_type == "ССБ 112") and (key in up112_hints) and not installed

        if in_repair:
            btn_text = REPAIR_TEXT
            btn_style = REPAIR_BOOTSTYLE
        else:
            btn_text = block_type
            btn_style = (
                "success" if installed else ("info" if info_tint else "secondary")
            )

        btn = row.block_btns[i]
        btn.configure(
            text=btn_text,
            bootstyle=btn_style,
            command=lambda k=key, b=block_type: on_block_left_click(k, b),
        )

        # ПКМ — ТОЛЬКО ДЛЯ СЕРЫХ кнопок (не установлен и без синей подсказки)
        try:
            btn.unbind("<Button-3>")
        except Exception:
            pass
        allow_repair = (not installed) and (not info_tint)
        if allow_repair:
            btn.bind(
                "<Button-3>", lambda e, k=key, b=block_type: toggle_block_repair(k, b)
            )

    # Галочка
    ready = all(products[key].get(t, False) for t in all_block_types)
    check_btn = row.check_btn
    if ready:
        check_btn.configure(
            bootstyle="info", state="normal", command=lambda k=key: mark_storage(k)
        )
        check_btn.unbind("<ButtonPress-1>")
        check_btn.unbind("<ButtonRelease-1>")
    else:
        check_btn.configure(bootstyle="light", state="disabled", command=None)
        check_btn.unbind("<ButtonPress-1>")
        check_btn.unbind("<ButtonRelease-1>")
        check_btn.bind(
            "<ButtonPress-1>",
            lambda e, k=key, bbs=row.block_btns, cb=check_btn: start_long_press(
                e, k, bbs, cb
            ),
        )
        check_btn.bind("<ButtonRelease-1>", lambda e, k=key: cancel_long_press(e, k))

    # Кнопка статуса (первая в строке)
    status_btn = row.winfo_children()[0]
    try:
        status_btn.unbind("<Button-3>")  # ПКМ по статусу не используем
    except Exception:
        pass

    if key in draft_products:
        status_btn.configure(
            text="ЗАГОТОВКА",
            bootstyle="warning",
            command=lambda k=key: toggle_draft_status(k),
        )
    else:
        status_btn.configure(
            text="НЕ ГОТОВО",
            bootstyle="danger",
            command=lambda k=key: toggle_draft_status(k),
        )


# ================== DIFF-РЕНДЕР СПИСКОВ ==================
def installed_blocks_count(blocks: dict) -> int:
    # считаем только по каноническим ключам
    return sum(1 for t in all_block_types if bool(blocks.get(t, False)))


def _get_current_comments():
    """Снять текущее содержимое комментариев из видимых строк (без пересоздания)."""
    current = {}
    for frame in (work_frame, storage_frame):
        for widget in frame.winfo_children():
            if isinstance(widget, tb.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, tb.Entry) and hasattr(child, "key_reference"):
                        current[child.key_reference] = child.get()
    return current


def _ensure_header(parent, hid, text, style):
    """Вернуть/создать заголовок с устойчивым id (hid). Меняет текст/стиль без пересоздания."""
    w = header_widgets.get(hid)
    if w is None or not w.winfo_exists():
        w = tb.Label(parent, text=text, style=style)
        w._vid = hid  # устойчивый визуальный id
        w.is_header = True
        header_widgets[hid] = w
        # пакуем «по умолчанию» (переупорядочим позже)
        w.pack(anchor="w", padx=10, pady=(10, 5))
    else:
        # только обновление параметров — без уничтожения
        if str(w.cget("text")) != text:
            w.config(text=text)
        # стиль меняем только если надо
        try:
            if str(w.cget("style")) != style:
                w.config(style=style)
        except Exception:
            pass
    return w


def _ensure_row(
    parent, key, blocks, current_comments, in_storage, show_factory_code=False
):
    src_dict = storage_widgets if in_storage else work_widgets
    other_dict = work_widgets if in_storage else storage_widgets

    # желаемый VID с учётом показа кода
    desired_vid = f"R:{'S' if in_storage else 'W'}:{key[0]}:{key[1]}:F{1 if show_factory_code else 0}"

    row = src_dict.get(key)

    # если была в другой колонке — удаляем там
    if key in other_dict:
        try:
            other_dict[key].destroy()
        except Exception:
            pass
        other_dict.pop(key, None)
        row = None

    # если строка есть, но VID (а значит и show_factory_code) не совпадает — пересоздаём
    if row is not None and row.winfo_exists():
        if getattr(row, "_vid", None) != desired_vid:
            try:
                row.destroy()
            except Exception:
                pass
            src_dict.pop(key, None)
            row = None

    if row is None or not row.winfo_exists():
        row = create_product_row(
            parent, key, blocks, current_comments, in_storage, show_factory_code
        )
        row._vid = desired_vid
        src_dict[key] = row
    else:
        update_row_widgets(key)

    return row


def _desired_work_sequence(current_comments):
    """
    Строит целевую последовательность для левой колонки ("В РАБОТЕ").
    Изменения:
      - В режиме show_draft_group заголовок "Остальные изделия" показывает количество.
      - Секция "В ремонте" размещается в КОНЦЕ (после "Остальные изделия").
    """
    seq = []

    # группируем по заводам всё, что "в работе"
    work_grouped = {factory: [] for factory in factory_order}
    for key, blocks in products.items():
        factory = key[1]
        if key in assembled_products or key in storage_products:
            continue
        work_grouped[factory].append((key, blocks))

    # ключи, где есть хотя бы один блок в ремонте
    repair_keys = {k for (k, _) in repair_blocks}
    sort_reverse = sort_order == "new_first"

    # ===== режим «красный пазл» 3-2-1-0 (ПКМ по пазлу) =====
    if blockcount_mode:
        flat_items = []
        for factory in factory_order:
            for key, blocks in work_grouped[factory]:
                flat_items.append((key, blocks))

        groups = {3: [], 2: [], 1: [], 0: []}
        for key, blocks in flat_items:
            installed_count = installed_blocks_count(blocks)
            groups[installed_count].append((key, blocks))

        headers = {
            3: "3 блока в наличии",
            2: "2 блока в наличии",
            1: "1 блок в наличии",
            0: "Нет блоков",
        }

        for count in [3, 2, 1, 0]:
            items = groups[count]
            if not items:
                continue

            items.sort(
                key=lambda item: (
                    0 if item[0] in draft_products else 1,
                    parse_date_str(product_dates.get(item[0], "01.01.00"), item[0]),
                    item[0][0],
                )
            )

            hid = f"W:grp:{count}"
            seq.append(
                {
                    "kind": "header",
                    "id": hid,
                    "text": f"{headers[count]} ({len(items)})",
                    "style": "Header3.TLabel",
                }
            )
            for key, blocks in items:
                seq.append(
                    {
                        "kind": "row",
                        "key": key,
                        "blocks": blocks,
                        "in_storage": False,
                        "show_factory_code": True,
                    }
                )
        return seq

    # ===== режим «пазл ЛКМ»: ЗАГОТОВКИ → ОСТАЛЬНЫЕ → РЕМОНТ (ремонт теперь в КОНЦЕ) =====
    if show_draft_group:
        # --- Заготовки (без ремонта) ---
        drafts = []
        for factory in factory_order:
            for key, blocks in work_grouped[factory]:
                if key in draft_products and key not in repair_keys:
                    drafts.append((key, blocks))

        if drafts:
            seq.append(
                {
                    "kind": "header",
                    "id": "W:hdr:drafts",
                    "text": f"Заготовки ({len(drafts)})",
                    "style": "FactoryHeader.TLabel",
                }
            )
            drafts.sort(
                key=lambda item: (
                    len(all_block_types) - installed_blocks_count(item[1]),
                    parse_date_str(product_dates.get(item[0], "01.01.00"), item[0]),
                )
            )
            groups = {3: [], 2: [], 1: [], 0: []}
            for key, blocks in drafts:
                cnt = installed_blocks_count(blocks)
                groups[cnt].append((key, blocks))

            for cnt in [3, 2, 1, 0]:
                if not groups[cnt]:
                    continue
                if cnt == 3:
                    text = f"Готовы к сборке ({len(groups[cnt])})"
                elif cnt == 2:
                    text = "2 блока в наличии"
                elif cnt == 1:
                    text = "1 блок в наличии"
                else:
                    text = "0 блоков в наличии"
                seq.append(
                    {
                        "kind": "header",
                        "id": f"W:hdr:drafts:{cnt}",
                        "text": text,
                        "style": "Header3.TLabel",
                    }
                )
                for key, blocks in groups[cnt]:
                    seq.append(
                        {
                            "kind": "row",
                            "key": key,
                            "blocks": blocks,
                            "in_storage": False,
                            "show_factory_code": True,
                        }
                    )

        # --- Остальные (не заготовки и не ремонт) ---
        others_by_factory = {}
        for factory in factory_order:
            for key, blocks in work_grouped[factory]:
                if key not in draft_products and key not in repair_keys:
                    others_by_factory.setdefault(factory, []).append((key, blocks))

        others_total = sum(len(v) for v in others_by_factory.values())

        if others_total > 0:
            # Заголовок теперь с количеством
            seq.append(
                {
                    "kind": "header",
                    "id": "W:hdr:others",
                    "text": f"Остальные изделия ({others_total})",
                    "style": "FactoryHeader.TLabel",
                }
            )

            for factory in factory_order:
                items = others_by_factory.get(factory, [])
                if not items:
                    continue
                seq.append(
                    {
                        "kind": "header",
                        "id": f"W:hdr:factory:{factory}",
                        "text": f"Завод: {factory}",
                        "style": "FactoryHeader.TLabel",
                    }
                )

                date_groups = {}
                for key, blocks in items:
                    date = product_dates.get(key, "??.??.??")
                    date_groups.setdefault(date, []).append((key, blocks))

                for date in sorted(
                    date_groups.keys(),
                    key=lambda d: parse_date_str(d, date_groups[d][0][0]),
                    reverse=sort_reverse,
                ):
                    seq.append(
                        {
                            "kind": "header",
                            "id": f"W:hdr:date:{factory}:{date}",
                            "text": format_date_genitive(date),
                            "style": "Header3.TLabel",
                        }
                    )
                    for key, blocks in sorted(date_groups[date], key=lambda x: x[0][0]):
                        seq.append(
                            {
                                "kind": "row",
                                "key": key,
                                "blocks": blocks,
                                "in_storage": False,
                                "show_factory_code": False,
                            }
                        )

        # --- В ремонте (ПЕРЕНЕСЕНО В КОНЕЦ) ---
        repair_items = []
        for factory in factory_order:
            for key, blocks in work_grouped[factory]:
                if key in repair_keys:
                    repair_items.append((key, blocks))

        if repair_items:
            seq.append(
                {
                    "kind": "header",
                    "id": "W:hdr:repair",
                    "text": f"В ремонте ({len(repair_items)})",
                    "style": "FactoryHeader.TLabel",
                }
            )
            repair_items.sort(
                key=lambda item: (
                    parse_date_str(product_dates.get(item[0], "01.01.00"), item[0]),
                    item[0][0],
                )
            )
            for key, blocks in repair_items:
                seq.append(
                    {
                        "kind": "row",
                        "key": key,
                        "blocks": blocks,
                        "in_storage": False,
                        "show_factory_code": True,
                    }
                )

        return seq

    # ===== обычный режим (без группировки «заготовок») =====
    for factory in factory_order:
        items = work_grouped[factory]
        if not items:
            continue
        seq.append(
            {
                "kind": "header",
                "id": f"W:hdr:factory:{factory}",
                "text": f"Завод: {factory}",
                "style": "FactoryHeader.TLabel",
            }
        )

        date_groups = {}
        for key, blocks in items:
            date = product_dates.get(key, "??.??.??")
            date_groups.setdefault(date, []).append((key, blocks))

        for date in sorted(
            date_groups.keys(),
            key=lambda d: parse_date_str(d, date_groups[d][0][0]),
            reverse=sort_reverse,
        ):
            seq.append(
                {
                    "kind": "header",
                    "id": f"W:hdr:date:{factory}:{date}",
                    "text": format_date_genitive(date),
                    "style": "Header3.TLabel",
                }
            )
            for key, blocks in sorted(date_groups[date], key=lambda x: x[0][0]):
                seq.append(
                    {
                        "kind": "row",
                        "key": key,
                        "blocks": blocks,
                        "in_storage": False,
                        "show_factory_code": False,
                    }
                )
    return seq


def all_blocks_done(key) -> bool:
    """Только по каноническим типам (чтобы лишние поля в старых данных не ломали логику)."""
    try:
        return all(products[key].get(t, False) for t in all_block_types)
    except Exception:
        return False


def clear_repair_for_key(key):
    """Снять все статусы РЕМОНТ для конкретного изделия (всех трёх блоков)."""
    global repair_blocks
    to_remove = [(k, b) for (k, b) in list(repair_blocks) if k == key]
    for pair in to_remove:
        repair_blocks.discard(pair)


def ensure_mutual_state(key):
    """
    Инвариант состояний: изделие не должно одновременно быть и в storage_products, и в assembled_products.
    """
    if key in assembled_products:
        storage_products.discard(key)
    # если изделие «в работе», оно не должно числиться в правой колонке/архиве
    if (
        key in products
        and key not in storage_products
        and key not in assembled_products
    ):
        storage_products.discard(key)
        assembled_products.discard(key)


def _desired_storage_sequence(current_comments):
    """Аналогичная последовательность для правой колонки (склад)."""
    seq = []

    # группируем
    storage_grouped = {factory: [] for factory in factory_order}
    for key, blocks in products.items():
        factory = key[1]
        if key in storage_products and key not in assembled_products:
            storage_grouped[factory].append((key, blocks))

    sort_reverse = sort_order == "new_first"

    for factory in factory_order:
        items = storage_grouped[factory]
        if not items:
            continue

        seq.append(
            {
                "kind": "header",
                "id": f"S:hdr:factory:{factory}",
                "text": f"Завод: {factory}",
                "style": "FactoryHeader.TLabel",
            }
        )

        date_groups = {}
        for key, blocks in items:
            date = storage_dates.get(key, "??.??.??")
            date_groups.setdefault(date, []).append((key, blocks))

        for date in sorted(
            date_groups.keys(),
            key=lambda d: parse_date_str(d, date_groups[d][0][0]),
            reverse=sort_reverse,
        ):
            # показываем только день.месяц
            dparts = (date or "??.??").split(".")
            dm_text = (
                f"Дата сборки: {dparts[0]}.{dparts[1]}"
                if len(dparts) >= 2
                else f"Дата сборки: {date}"
            )
            seq.append(
                {
                    "kind": "header",
                    "id": f"S:hdr:date:{factory}:{date}",
                    "text": dm_text,
                    "style": "Header3.TLabel",
                }
            )
            for key, blocks in sorted(date_groups[date], key=lambda x: x[0][0]):
                seq.append(
                    {
                        "kind": "row",
                        "key": key,
                        "blocks": blocks,
                        "in_storage": True,
                        "show_factory_code": False,
                    }
                )

    return seq


def _apply_sequence(frame, sequence, which):
    """
    Применяем diff без полной перерисовки.
    Главное отличие: мы переупорядочиваем ЦЕЛЫЕ СЕКЦИИ, а не отдельные заголовки,
    поэтому «Заготовки / В ремонте / Остальные изделия» всегда идут в правильном порядке.
    """
    current_comments = _get_current_comments()

    # 1) Гарантируем наличие нужных виджетов (без пересоздания лишнего)
    widgets_in_order, seen_vids = [], set()
    for item in sequence:
        if item["kind"] == "header":
            w = _ensure_header(frame, item["id"], item["text"], item["style"])
        else:
            w = _ensure_row(
                parent=frame,
                key=item["key"],
                blocks=item["blocks"],
                current_comments=current_comments,
                in_storage=item["in_storage"],
                show_factory_code=item.get("show_factory_code", False),
            )
        widgets_in_order.append(w)
        seen_vids.add(getattr(w, "_vid", None))

    # 2) Удаляем те, которых больше нет в целевом списке
    for child in list(frame.winfo_children()):
        vid = getattr(child, "_vid", None)
        if vid and vid not in seen_vids:
            try:
                child.destroy()
            except Exception:
                pass

    # 3) Собираем ГРУППЫ (секции)
    #    Для WORK: секции = drafts / repair / others (вся «остальная» часть с заводами и датами).
    #    Для STORAGE: секция = каждый завод.
    def is_section_header(vid: str) -> bool:
        if not vid:
            return False
        if vid.startswith("W:hdr:drafts"):  # Заготовки
            return True
        if vid.startswith("W:hdr:repair"):  # В ремонте
            return True
        if vid.startswith("W:hdr:others"):  # Остальные изделия
            return True
        if vid.startswith("S:hdr:factory:"):  # Склад: секция по заводу
            return True
        return False

    desired_groups = []
    cur = []

    for w in widgets_in_order:
        vid = getattr(w, "_vid", "")
        if getattr(w, "is_header", False) and is_section_header(vid):
            if cur:
                desired_groups.append(cur)
            cur = [w]  # начинаем новую секцию
        else:
            (cur or (cur := [])).append(w)
    if cur:
        desired_groups.append(cur)

    # 4) Вспомогалка упаковать с дефолтными паддингами
    def _pack(w, before=None):
        try:
            w.pack_forget()
        except Exception:
            pass
        if getattr(w, "is_header", False):
            # аккуратные отступы для "заводских" заголовков
            style_name = ""
            try:
                style_name = str(w.cget("style"))
            except Exception:
                pass
            if "FactoryHeader" in style_name:
                kwargs = dict(anchor="w", pady=(10, 0))
            else:
                kwargs = dict(anchor="w", padx=10, pady=(10, 5))
        else:
            kwargs = dict(fill=X, pady=PADY_PRODUCT)
        if before is not None:
            w.pack(before=before, **kwargs)
        else:
            w.pack(**kwargs)

    # 5) Минимальная перестановка секций (без полной перерисовки)
    #    Каждую секцию вставляем перед самым ранним элементом из оставшихся секций.
    def current_order():
        return [
            w for w in frame.winfo_children() if getattr(w, "_vid", None) in seen_vids
        ]

    for i in range(len(desired_groups)):
        cur_order = current_order()

        # Если секция уже единым блоком и стоит на месте — пропускаем.
        try:
            start = cur_order.index(desired_groups[i][0])
            ok = all(
                start + j < len(cur_order)
                and cur_order[start + j] is desired_groups[i][j]
                for j in range(len(desired_groups[i]))
            )
        except ValueError:
            ok = False
        if ok:
            continue

        # Якорь — «самый ранний» виджет из будущих секций
        later_firsts = [
            g[0] for g in desired_groups[i + 1 :] if g and g[0] in cur_order
        ]
        anchor = (
            min(later_firsts, key=lambda w: cur_order.index(w))
            if later_firsts
            else None
        )

        # Вставляем секцию целиком (с конца к началу, чтобы сохранить порядок)
        for w in reversed(desired_groups[i]):
            _pack(w, before=anchor)
            anchor = w


def update_product_list(preserve_scroll=True):
    """
    Дифф-обновление без полной перерисовки.
    Пересоздаём только то, что реально меняется, и аккуратно переупаковываем.
    """
    global show_draft_group, storage_count_cache

    # 0) сохранить скролл
    if preserve_scroll:
        try:
            ws = work_canvas.yview()[0]
            ss = storage_canvas.yview()[0]
        except Exception:
            ws = ss = 0.0

    # 1) служебные данные для шапок
    #   (считаем после диффа, но до отрисовки — числа должны быть актуальны)
    work_count = sum(
        1
        for k in products
        if (k not in assembled_products and k not in storage_products)
    )
    storage_count = sum(1 for k in storage_products if k not in assembled_products)

    # обновление заголовков/кнопок (не трогаем списки)
    work_header_label.config(text=f"{TEXT_WORK_TITLE} ({work_count})")
    storage_header_label.config(text=f"{TEXT_STORAGE_TITLE} ({storage_count})")

    storage_count_cache = storage_count
    try:
        if not storage_visible:
            lbl = f"◀ ({storage_count_cache})"
            toggle_storage_button.config(text=lbl, width=len(lbl))
        else:
            toggle_storage_button.config(text="▶", width=2)
    except Exception:
        pass

    # проверка группы «заготовки»
    repair_keys = {k for (k, _) in repair_blocks}
    draft_count = sum(
        1
        for k in products
        if (
            k not in assembled_products
            and k not in storage_products
            and k in draft_products
            and k not in repair_keys
        )
    )
    if show_draft_group and draft_count == 0:
        show_draft_group = False
        try:
            _refresh_sort_buttons()
        except Exception:
            pass

    # 2) строим целевые последовательности слева/справа
    work_seq = _desired_work_sequence(_get_current_comments())
    storage_seq = _desired_storage_sequence(_get_current_comments())

    # 3) применяем дифф-патч «без перерисовки»
    _apply_sequence(work_frame, work_seq, which="work")
    _apply_sequence(storage_frame, storage_seq, which="storage")

    # 4) восстановление скролла мягко
    if preserve_scroll:
        work_canvas.update_idletasks()
        try:
            work_canvas.yview_moveto(ws)
            storage_canvas.yview_moveto(ss)
        except Exception:
            pass


def search_product():
    search_text = search_var.get().strip()
    if not search_text.isdigit() or len(search_text) < 2 or len(search_text) > 5:
        messagebox.showerror("Ошибка", "Введите от 2 до 5 цифр (номер и завод)")
        return

    number_part = search_text[:-1].zfill(4)
    factory_code = search_text[-1]
    if factory_code not in factory_mapping:
        messagebox.showerror("Ошибка", "Некорректный код завода (1-6)")
        return

    factory_name = factory_mapping[factory_code]
    key = (number_part, factory_name)
    result = []

    if (
        key in products
        and key not in storage_products
        and key not in assembled_products
    ):
        result.append(f"СЕРП №{number_part} ({factory_name}) — ожидает сборки")
        result.append(f"Дата добавления: {product_dates.get(key, '??.??')}")
        blocks = products.get(key, {})
        result.append("Наличие блоков:")
        result.append(f"112: {'✓' if blocks.get('ССБ 112', False) else '✕'}")
        result.append(f"161: {'✓' if blocks.get('ССБ 161', False) else '✕'}")
        result.append(f"114: {'✓' if blocks.get('ССБ 114', False) else '✕'}")

    elif key in storage_products:
        ds = storage_dates.get(key, "??.??.??")
        result.append(
            f"СЕРП №{number_part} ({factory_name}) — собран, ждёт отправки на склад"
        )
        result.append(f"Дата сборки: {ds}")

    elif key in assembled_products:
        date_ship = assembly_dates.get(key, "??.??.??")
        date_asm = storage_dates.get(key, date_ship)  # для старых совпадает
        result.append(f"СЕРП №{number_part} ({factory_name}) — отправлен на склад")
        result.append(f"Дата отправки: {date_ship}")
        result.append(f"Дата сборки: {date_asm}")

    else:
        result.append(f"СЕРП №{number_part} ({factory_name}) не найден")

    messagebox.showinfo("Результат поиска", "\n".join(result))


def format_date_display(date_str):
    return format_date_genitive(date_str)


def export_to_excel():
    try:
        import pandas as pd
    except ImportError:
        messagebox.showerror(
            "Ошибка",
            "Для экспорта в Excel установите pandas: pip install pandas",
        )
        return

    data = []
    for key in assembled_products:
        number, factory = key
        # Дата отправки (в архив)
        date_ship = assembly_dates.get(key, "н/д")
        # Дата сборки — если нет, приравниваем к дате отправки
        date_asm = storage_dates.get(key, date_ship)

        year = assembly_years.get(key, "н/д")
        factory_code = factory_reverse_mapping.get(factory, "?")
        serial_display = f"{year}{factory_code}{number.zfill(4)}"

        data.append(
            {
                "Дата отправки": date_ship,
                "Дата сборки": date_asm,
                "Серийный номер": serial_display,
                "Завод": factory,
            }
        )

    if not data:
        messagebox.showinfo("Информация", "Нет данных для экспорта")
        return

    df = pd.DataFrame(data)
    file_path = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
    )
    if file_path:
        try:
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Успех", f"Данные экспортированы в {file_path}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {str(e)}")


def move_factory_to_top(factory):
    global factory_order
    new_order = [f for f in factory_order if f != factory]
    new_order.insert(0, factory)
    factory_order = new_order
    save_data()
    update_product_list()
    update_button_colors()
    if notebook.tab(notebook.select(), "text") == "Архив сборки":
        update_assembly_archive()


def reset_factory_order():
    global factory_order
    factory_order = ["ВЕКТОР", "ИНТЕГРАЛ", "РЗП", "КНИИТМУ", "СВТ", "СИГНАЛ"]
    save_data()
    update_product_list()
    update_button_colors()
    if notebook.tab(notebook.select(), "text") == "Архив сборки":
        update_assembly_archive()


def toggle_sort_order():
    global sort_order, sort_button, sort_down_icon, sort_up_icon

    # переключаем режим сортировки
    sort_order = "old_first" if sort_order == "new_first" else "new_first"

    # обновляем кнопку (всегда есть текст; иконка — если загружена)
    if sort_button is not None:
        # базовый текст (стрелки)
        sort_button.configure(text=("⏫" if sort_order == "new_first" else "⏬"))

        # попробовать поставить PNG-иконку
        img = None
        try:
            img = sort_down_icon if sort_order == "new_first" else sort_up_icon
        except NameError:
            img = None

        if img:
            sort_button.configure(image=img, compound="left")
        else:
            sort_button.configure(image="", compound=None)

    save_data()
    save_scroll_positions()
    update_product_list()
    restore_scroll_positions()


def set_sort_mode(mode: str):
    """Единая точка переключения режима левой колонки."""
    global work_sort_mode, show_draft_group, blockcount_mode
    mode = mode.lower()
    if mode not in ("drafts", "blocks", "factories"):
        return
    work_sort_mode = mode

    # Синхронизируем с уже существующими флагами
    if mode == "drafts":
        show_draft_group = True
        blockcount_mode = False
    elif mode == "blocks":
        show_draft_group = False
        blockcount_mode = True
    else:  # "factories"
        show_draft_group = False
        blockcount_mode = False

    _refresh_sort_buttons()
    update_product_list()  # перерисовать левую колонку


def _refresh_sort_buttons():
    """Подсветить активную кнопку режима."""
    try:
        btn_mode_drafts.configure(
            bootstyle="primary" if work_sort_mode == "drafts" else "secondary"
        )
        btn_mode_blocks.configure(
            bootstyle="primary" if work_sort_mode == "blocks" else "secondary"
        )
        btn_mode_factories.configure(
            bootstyle="primary" if work_sort_mode == "factories" else "secondary"
        )
    except Exception:
        pass


# === Режим "красный пазл" (группы 3-2-1-0) ===
blockcount_mode = False  # глобальный флаг ПКМ-режима

work_sort_mode = "factories"

# Заглушки (старые кнопки больше не используем)
draft_group_button = None
sort_button = None

# Ссылки на новые кнопки (инициализируются при создании UI)
btn_mode_drafts = None
btn_mode_blocks = None
btn_mode_factories = None


def on_draft_puzzle_left_click(event=None):
    """
    ЛКМ по пазлу: работает как раньше (вкл/выкл 'заготовки'),
    но только если НЕ включён красный режим.
    """
    if blockcount_mode:
        return "break"  # игнорируем ЛКМ в красном режиме
    toggle_draft_group()


def toggle_blockcount_mode(event=None):
    """
    ПКМ по пазлу: включить/выключить красный режим 3-2-1-0.
    При включении красного режима всегда возвращаем кнопке серый/оранжевый вид
    и выключаем группировку 'заготовок'.
    """
    global blockcount_mode, show_draft_group
    blockcount_mode = not blockcount_mode

    if blockcount_mode:
        # включили красный режим: отключаем 'заготовки' и красим пазл
        show_draft_group = False
        draft_group_button.configure(text="🧩", bootstyle="danger")
    else:
        # выключили красный режим: возвращаем исходный вид пазла
        draft_group_button.configure(
            text="🧩", bootstyle=("warning" if show_draft_group else "secondary")
        )

    update_product_list()


def save_settings():
    global data_path, scaling_factor
    data_path = data_path_var.get()
    save_data()
    messagebox.showinfo("Настройки", "Настройки успешно сохранены!")
    update_product_list()


def update_button_colors():
    for factory, btn in factory_buttons.items():
        btn.configure(
            bootstyle="primary" if factory == factory_order[0] else "secondary"
        )


def get_week_range():
    today = datetime.now()
    start = today - timedelta(days=today.weekday())
    end = start + timedelta(days=6)
    return start, end


def count_assembled_this_month():
    now = datetime.now()
    count = 0
    for key in assembled_products:
        date_str = assembly_dates.get(key, "")
        try:
            if "." in date_str:
                parts = date_str.split(".")
                if len(parts) == 2:
                    day, month = int(parts[0]), int(parts[1])
                    year = now.year
                elif len(parts) == 3:
                    day, month, year_part = int(parts[0]), int(parts[1]), int(parts[2])
                    year = 2000 + year_part if year_part < 100 else year_part
                else:
                    continue
                if month == now.month and year == now.year:
                    count += 1
        except:
            continue
    return count


def count_assembled_this_week():
    start, end = get_week_range()
    count = 0
    for key in assembled_products:
        date_str = assembly_dates.get(key, "")
        try:
            if "." in date_str:
                parts = date_str.split(".")
                if len(parts) == 2:
                    day, month = int(parts[0]), int(parts[1])
                    year = datetime.now().year
                elif len(parts) == 3:
                    day, month, year_part = int(parts[0]), int(parts[1]), int(parts[2])
                    year = 2000 + year_part if year_part < 100 else year_part
                else:
                    continue
                date = datetime(year, month, day)
                if start <= date <= end:
                    count += 1
        except:
            continue
    return count


def clear_archive():
    global archive_scroll_position
    if not messagebox.askyesno(
        "Подтверждение", "Вы уверены, что хотите полностью очистить архив сборки?"
    ):
        return
    save_scroll_positions()
    archive_keys = set(assembled_products)
    assembled_products.clear()
    for key in archive_keys:
        products.pop(key, None)
        comments.pop(key, None)
        product_dates.pop(key, None)
        assembly_dates.pop(key, None)
        assembly_years.pop(key, None)
    update_product_list()
    update_assembly_archive()
    save_data()


def delete_from_archive(key):
    global archive_scroll_position
    if not messagebox.askyesno("Подтверждение", "Удалить этот СЕРП из архива?"):
        return
    save_scroll_positions()
    if key in assembled_products:
        assembled_products.remove(key)
    assembly_dates.pop(key, None)
    assembly_years.pop(key, None)
    comments.pop(key, None)
    products.pop(key, None)
    update_assembly_archive()
    update_product_list()
    save_data()


def set_archive_mode(mode):
    global archive_view_mode, archive_scroll_position
    save_scroll_positions()
    archive_view_mode = mode
    update_assembly_archive()


def update_assembly_archive():
    for widget in archive_frame.winfo_children():
        widget.destroy()

    main_container = tb.Frame(archive_frame, style="Main.TFrame")
    main_container.pack(fill=BOTH, expand=True, padx=10, pady=10)
    main_container.configure(style="Main.TFrame")

    # Левая колонка (список)
    left_frame = tb.Frame(main_container, style="Main.TFrame")
    left_frame.pack(side=LEFT, fill=BOTH, expand=True, padx=(0, 10))
    left_frame.configure(style="Main.TFrame")

    # Правая колонка (статистика + кнопки)
    right_frame = tb.Frame(main_container, style="Main.TFrame", width=300)
    right_frame.pack(side=RIGHT, fill=Y, padx=(10, 0))
    right_frame.pack_propagate(False)
    right_frame.configure(style="Main.TFrame")

    top_spacer = tb.Frame(right_frame, style="Main.TFrame", height=0)
    top_spacer.pack(side=TOP, fill=BOTH, expand=True)

    bottom_container = tb.Frame(right_frame, style="Main.TFrame")
    bottom_container.pack(side=BOTTOM, fill=X, pady=(0, 25))

    # Переключатель режима
    mode_frame = tb.Frame(left_frame, style="Main.TFrame")
    mode_frame.pack(fill=X, pady=(0, 10))
    mode_frame.configure(style="Main.TFrame")

    tb.Label(
        mode_frame,
        text="Режим просмотра:",
        font=("Segoe UI", 12, "bold"),
        style="Main.TLabel",
    ).pack(side=LEFT, padx=10, pady=10)

    tb.Button(
        mode_frame,
        text="Журнал",
        width=12,
        bootstyle="primary" if archive_view_mode == "journal" else "secondary",
        command=lambda: set_archive_mode("journal"),
    ).pack(side=LEFT, padx=5, pady=10)

    tb.Button(
        mode_frame,
        text="История",
        width=12,
        bootstyle="primary" if archive_view_mode == "history" else "secondary",
        command=lambda: set_archive_mode("history"),
    ).pack(side=LEFT, padx=(5, 10), pady=10)

    # Прокручиваемая область слева
    list_container = tb.Frame(left_frame, style="Main.TFrame")
    list_container.pack(fill=BOTH, expand=True)
    list_container.configure(style="Main.TFrame")

    list_canvas = tb.Canvas(list_container, background=BG_MAIN, highlightthickness=0)
    list_scrollbar = tb.Scrollbar(
        list_container,
        orient=VERTICAL,
        command=list_canvas.yview,
        bootstyle="light-round",
    )
    list_content = tb.Frame(list_canvas, style="Main.TFrame")
    list_content.configure(style="Main.TFrame")
    list_canvas.create_window((0, 0), window=list_content, anchor=NW)

    def on_canvas_configure(event):
        list_canvas.itemconfig("all", width=event.width)

    def on_frame_configure(event):
        list_canvas.configure(scrollregion=list_canvas.bbox("all"))

    list_canvas.bind("<Configure>", on_canvas_configure)
    list_content.bind("<Configure>", on_frame_configure)
    list_canvas.configure(yscrollcommand=list_scrollbar.set)
    list_scrollbar.pack(side=RIGHT, fill=Y)
    list_canvas.pack(side=LEFT, fill=BOTH, expand=True)

    # Правая колонка — статистика
    stats_header = tb.Label(
        right_frame, text="Статистика архива", font=("Arial", 14), style="Main.TLabel"
    )
    stats_header.pack(pady=(0, 10))

    stats_container = tb.Frame(right_frame, style="Main.TFrame")
    stats_container.pack(fill=X, pady=(0, 15), padx=5)
    stats_container.configure(style="Main.TFrame")

    week_count = count_assembled_this_week()
    month_count = count_assembled_this_month()
    total_count = len(assembled_products)

    months = [
        "январь",
        "февраль",
        "март",
        "апрель",
        "май",
        "июнь",
        "июль",
        "август",
        "сентябрь",
        "октябрь",
        "ноябрь",
        "декабрь",
    ]
    month_name = months[datetime.now().month - 1]

    stats_data = [
        ("За неделю:", f"{week_count} шт."),
        (f"За {month_name}:", f"{month_count} шт."),
        ("В архиве:", f"{total_count} шт."),
    ]

    for label, value in stats_data:
        stat_row = tb.Frame(stats_container, style="Main.TFrame")
        stat_row.pack(fill=X, pady=4, padx=10)
        stat_row.configure(style="Main.TFrame")
        tb.Label(stat_row, text=label, font=("Segoe UI", 11), style="Main.TLabel").pack(
            side=LEFT
        )
        tb.Label(stat_row, text=value, font=("Segoe UI", 11), style="Main.TLabel").pack(
            side=RIGHT
        )

    clear_btn = tb.Button(
        right_frame,
        text="Очистить архив",
        width=20,
        bootstyle="secondary",
        command=clear_archive,
    )
    clear_btn.pack(pady=10)

    export_btn = tb.Button(
        right_frame,
        text="Экспорт в Excel",
        width=20,
        bootstyle="success",
        command=export_to_excel,
    )
    export_btn.pack(pady=10)

    # Данные архива
    all_items = list(assembled_products)

    if archive_view_mode == "history":
        # Сортируем по дате «отправлено на склад» (assembly_dates используется как дата архива)
        all_items.sort(
            key=lambda k: parse_date_str(assembly_dates.get(k, "01.01.00")),
            reverse=True,
        )

        # Группируем по дате
        date_groups = {}
        for key in all_items:
            date_str = assembly_dates.get(key, "??.??.??")
            date_groups.setdefault(date_str, []).append(key)

        # Список дат (сверху — самые свежие)
        sorted_dates = sorted(
            date_groups.keys(), key=lambda d: parse_date_str(d), reverse=True
        )

        for date_str in sorted_dates:
            day_count = len(date_groups.get(date_str, []))

            # Заголовок группы по дате
            date_header = tb.Frame(list_content, style="Main.TFrame")
            date_header.pack(fill=X, padx=10, pady=(15, 5))
            dm = (date_str or "").split(".")
            ds_short = f"{dm[0]}.{dm[1]}" if len(dm) >= 2 else (date_str or "??.??")

            tb.Label(
                date_header,
                text=f"Отправлено на склад {day_count} шт.  ({ds_short}) ",
                font=FONT_BOLD,
                style="Main.TLabel",
            ).pack(anchor="w")

            # далее как у вас: группировка по заводам и создание строк
            factory_groups = {}
            for key in date_groups[date_str]:
                factory = key[1]
                factory_groups.setdefault(factory, []).append(key)

            sorted_factories = sorted(
                factory_groups.keys(),
                key=lambda f: (
                    factory_order.index(f) if f in factory_order else len(factory_order)
                ),
            )

            if archive_view_mode == "history":
                ...
                for factory in sorted_factories:
                    sorted_items = sorted(
                        factory_groups[factory], key=lambda k: int(k[0])
                    )
                    for key in sorted_items:
                        # дата сборки: сначала пробуем storage_dates, иначе assembly_dates
                        date_asm = storage_dates.get(key) or assembly_dates.get(
                            key, "??.??.??"
                        )
                        # показываем только дд.мм
                        dm = (date_asm or "").split(".")
                        ds_short = (
                            f"{dm[0]}.{dm[1]}"
                            if len(dm) >= 2
                            else (date_asm or "??.??")
                        )

                        # ВОТ ЭТИ ДВЕ СТРОКИ НУЖНЫ:
                        row = tb.Frame(list_content, style="Main.TFrame")
                        row.pack(fill=X, pady=3, padx=20)

                        tb.Label(
                            row,
                            text=f"СЕРП №{key[0]} ({key[1]}) // ДС: {ds_short}",
                            font=FONT_PRIMARY,
                            style="Main.TLabel",
                        ).pack(side="left", fill=X, expand=True)

                        tb.Button(
                            row,
                            text="✕",
                            width=3,
                            bootstyle="secondary",
                            command=lambda k=key: delete_from_archive(k),
                        ).pack(side="right", padx=(10, 0))

    else:
        # Режим «Журнал» без изменений
        grouped = {f: [] for f in factory_order}
        for key in all_items:
            if key[1] in grouped:
                grouped[key[1]].append(key)

        for factory in factory_order:
            if not grouped[factory]:
                continue

            factory_header = tb.Frame(list_content, style="Main.TFrame")
            factory_header.pack(fill=X, padx=10, pady=(15, 5))
            factory_header.configure(style="Main.TFrame")
            tb.Label(
                factory_header, text=factory, font=FONT_BOLD, style="Main.TLabel"
            ).pack(anchor=W)

            for key in sorted(grouped[factory], key=lambda x: int(x[0])):
                number, factory_name = key
                factory_code = factory_reverse_mapping.get(factory_name, "?")
                year = assembly_years.get(key, "??")
                serial_display = f"{year}{factory_code}{number.zfill(4)}"
                date = assembly_dates.get(key, "??.??.??")

                row = tb.Frame(list_content, style="Main.TFrame")
                row.pack(fill=X, pady=3, padx=20)
                row.configure(style="Main.TFrame")

                # Здесь оставляем старую подпись (если захотите — тоже поменяем)
                # Если нет даты сборки — берём дату отправки на хранение
                date_asm = assembly_dates.get(key) or storage_dates.get(key, "??.??.??")
                # если нет 'даты сборки' — берём 'дату отправки'

                tb.Label(
                    row,
                    text=f"СЕРП №{serial_display} - Дата сборки: {date_asm}",
                    font=FONT_PRIMARY,
                    style="Main.TLabel",
                ).pack(side="left", fill=X, expand=True)

                tb.Button(
                    row,
                    text="✕",
                    width=3,
                    bootstyle="secondary",
                    command=lambda k=key: delete_from_archive(k),
                ).pack(side="right", padx=(10, 0))

    list_canvas.update_idletasks()
    list_canvas.yview_moveto(0)

    def update_bg(widget):
        try:
            widget.configure(background=BG_MAIN)
        except:
            pass
        for child in widget.winfo_children():
            update_bg(child)

    update_bg(main_container)


def apply_request_filter():
    global request_filter
    request_filter = filter_var.get()
    update_request_table()


def update_request_table():
    if request_table is None:
        return

    request_table.delete(*request_table.get_children())
    missing_power, missing_detector, missing_cap = [], [], []

    for key, blocks in products.items():
        if key in assembled_products or key in storage_products:
            continue

        number, factory_name = key
        factory_code = [k for k, v in factory_mapping.items() if v == factory_name][0]
        display_code = f"{number} [{factory_code}]"

        if request_filter != "Все":
            if request_filter == "Заготовки":
                if key not in draft_products:
                    continue
            elif request_filter in factory_mapping.values():
                if factory_name != request_filter:
                    continue

        if not blocks["ССБ 112"]:
            missing_power.append(
                "" + display_code if key in draft_products else display_code
            )
        if not blocks["ССБ 161"]:
            missing_detector.append(
                "" + display_code if key in draft_products else display_code
            )
        if not blocks["ССБ 114"]:
            missing_cap.append(
                "" + display_code if key in draft_products else display_code
            )

    def sort_by_factory(x: str):
        try:
            return int(x.split("[")[-1].split("]")[0])
        except:
            return 999

    missing_power.sort(key=sort_by_factory)
    missing_detector.sort(key=sort_by_factory)
    missing_cap.sort(key=sort_by_factory)

    max_len = max(len(missing_power), len(missing_detector), len(missing_cap))
    for i in range(max_len):
        row = (
            missing_power[i] if i < len(missing_power) else "",
            missing_detector[i] if i < len(missing_detector) else "",
            missing_cap[i] if i < len(missing_cap) else "",
        )
        request_table.insert("", "end", values=row)


def on_tab_change(event):
    save_scroll_positions()
    selected_tab = notebook.tab(notebook.select(), "text")
    if selected_tab == "Запросить блоки":
        update_request_table()
    elif selected_tab == "Архив сборки":
        update_assembly_archive()


def _on_mousewheel(event):
    x, y = root.winfo_pointerx(), root.winfo_pointery()
    widget = root.winfo_containing(x, y)
    while widget:
        if isinstance(widget, tb.Canvas):
            step = 0
            if event.delta:
                step = int(-1 * (event.delta / 120))
            else:
                if event.num == 4:
                    step = -1
                elif event.num == 5:
                    step = 1

            y0, y1 = widget.yview()
            # Блокируем прокрутку, если уже на самом верху/низу
            if (step < 0 and y0 <= 0.0) or (step > 0 and y1 >= 1.0):
                return "break"

            widget.yview_scroll(step, "units")
            return "break"
        widget = widget.master


def toggle_draft_group():
    global show_draft_group
    has_drafts = len(draft_products) > 0

    if not has_drafts:
        if show_draft_group:
            show_draft_group = False
            draft_group_button.configure(bootstyle="secondary")
            save_data()
            update_product_list()
        else:
            messagebox.showinfo("Информация", "Нет заготовок для отображения")
        return

    show_draft_group = not show_draft_group
    draft_group_button.configure(
        bootstyle="warning" if show_draft_group else "secondary"
    )
    save_data()
    update_product_list()


def collapse_storage_on_start():
    """Сворачивает правую колонку при старте, если она сейчас видима."""
    global storage_visible
    root.update_idletasks()
    if storage_visible:
        toggle_storage_visibility()


def toggle_storage_visibility():
    global storage_visible, storage_count_cache
    if storage_visible:
        storage_frame_container.grid_remove()
        main_area.columnconfigure(0, weight=1)
        main_area.columnconfigure(1, weight=0)
        label = f"◀ ({storage_count_cache})"
        toggle_storage_button.config(text=label, width=len(label))
        storage_visible = False
    else:
        storage_frame_container.grid()
        main_area.columnconfigure(0, weight=5)
        main_area.columnconfigure(1, weight=2)
        toggle_storage_button.config(text="▶", width=2)
        storage_visible = True


def center_window(initial=False):
    """Ставит окно по центру.
    initial=True — использовать запрошенный размер (без мигания)."""
    try:
        if initial:
            # размеры, которые запрашивает layout, пока окно скрыто
            w = max(100, root.winfo_reqwidth())
            h = max(100, root.winfo_reqheight())
            # запасной план: если вдруг всё ещё маленькие — возьмём из текущей geometry()
            if w < 200 or h < 200:
                geom = root.geometry()
                if "x" in geom:
                    wh = geom.split("+")[0]
                    w2, h2 = [int(x) for x in wh.split("x")]
                    w, h = max(w, w2), max(h, h2)
        else:
            root.update_idletasks()
            w = max(100, root.winfo_width())
            h = max(100, root.winfo_height())
            if w < 200 or h < 200:
                w = max(w, root.winfo_reqwidth())
                h = max(h, root.winfo_reqheight())

        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        x = (sw - w) // 2
        y = (sh - h) // 2
        root.geometry(f"{w}x{h}+{x}+{y}")
    except Exception as e:
        print("center_window() error:", e)


load_data()

# --- СБРОС РЕЖИМОВ ПЕРЕД ПОСТРОЕНИЕМ UI ---
show_draft_group = False
blockcount_mode = False
work_sort_mode = "factories"
root = tb.Window(themename="litera")
root.tk.call("tk", "scaling", scaling_factor)
root.title("serp-base")
root.withdraw()  # скрываем окно на время постройки UI
root.geometry("600x700")
root.minsize(200, 600)
root.configure(background=BG_MAIN)
try:
    root.iconbitmap(resource_path("new.ico"))
except:
    print("Иконка не найдена")

# Загрузка изображений для кнопок
try:
    sklad_on_img = Image.open(resource_path("SKLAD32_ON.png")).resize(
        (32, 32), Image.LANCZOS
    )
    sklad_off_img = Image.open(resource_path("SKLAD32_OFF.png")).resize(
        (32, 32), Image.LANCZOS
    )
    sklad_on_icon = ImageTk.PhotoImage(sklad_on_img)
    sklad_off_icon = ImageTk.PhotoImage(sklad_off_img)

    zag_on_img = Image.open(resource_path("ZAG32_ON.png")).resize(
        (32, 32), Image.LANCZOS
    )
    zag_off_img = Image.open(resource_path("ZAG32_OFF.png")).resize(
        (32, 32), Image.LANCZOS
    )
    zag_on_icon = ImageTk.PhotoImage(zag_on_img)
    zag_off_icon = ImageTk.PhotoImage(zag_off_img)

    sort_down_img = Image.open(resource_path("SORT32_ON.png")).resize(
        (32, 32), Image.LANCZOS
    )
    sort_up_img = Image.open(resource_path("SORT32_OFF.png")).resize(
        (32, 32), Image.LANCZOS
    )
    sort_down_icon = ImageTk.PhotoImage(sort_down_img)
    sort_up_icon = ImageTk.PhotoImage(sort_up_img)
except Exception as e:
    print(f"Ошибка загрузки изображений: {e}")
    sklad_on_icon = sklad_off_icon = zag_on_icon = zag_off_icon = sort_down_icon = (
        sort_up_icon
    ) = None

style = tb.Style()
style.configure("Main.TFrame", background=BG_MAIN)
style.configure("WorkArea.TFrame", background=BG_WORK_AREA)
style.configure("Header.TFrame", background=BG_HEADER)
style.configure("Settings.TFrame", background=BG_SETTINGS)
style.configure("Wide.Vertical.TScrollbar", width=14)
style.configure("Save.TButton", font=FONT_SAVE_BUTTON)
style.configure("Status.TButton", font=FONT_STATUS_BUTTON)
style.configure("Block.TButton", font=FONT_BLOCK)
style.configure(
    "Header1.TLabel", font=FONT_HEADER1, background=BG_MAIN, foreground=TEXT_HEADER
)
style.configure(
    "Header2.TLabel", font=FONT_HEADER2, background=BG_MAIN, foreground=TEXT_HEADER
)
style.configure(
    "Header3.TLabel", font=FONT_HEADER3, background=BG_MAIN, foreground=TEXT_DATE
)
style.configure(
    "FactoryHeader.TLabel",
    font=FONT_FACTORY,
    background=BG_MAIN,
    foreground=TEXT_FACTORY,
)
style.configure(
    "Main.TLabel", background=BG_MAIN, foreground=TEXT_MAIN, font=FONT_PRIMARY
)
style.configure(
    "WorkArea.TLabel", background=BG_WORK_AREA, foreground=TEXT_MAIN, font=FONT_PRIMARY
)
style.configure(
    "Bold.TLabel", font=FONT_SETTINGS, background=BG_MAIN, foreground=TEXT_HEADER
)
style.configure(
    "BlockCounter.TLabel",
    font=FONT_COUNTER,
    foreground=TEXT_COUNTER,
    background=BG_MAIN,
)

notebook = tb.Notebook(root, bootstyle="primary")
notebook.pack(fill=BOTH, expand=True, padx=10, pady=10)
frame_main = tb.Frame(notebook, style="Main.TFrame")
frame_assembly = tb.Frame(notebook, style="Main.TFrame")
frame_request = tb.Frame(notebook, style="Main.TFrame")
frame_settings = tb.Frame(notebook, style="Settings.TFrame")
notebook.add(frame_main, text="Основной учёт")
notebook.add(frame_assembly, text="Архив сборки")
notebook.add(frame_request, text="Запросить блоки")
notebook.add(frame_settings, text="Настройки")

request_panel = tb.Frame(frame_request, style="Main.TFrame")
request_panel.pack(fill=BOTH, expand=True, padx=10, pady=10)

# Панель фильтров
filter_frame = tb.Frame(request_panel, style="Main.TFrame")
filter_frame.pack(fill=X, pady=(0, 10))

tb.Label(filter_frame, text="Фильтр:", style="Main.TLabel").pack(side=LEFT, padx=5)

filter_var = tb.StringVar(value="Все")
filter_combo = tb.Combobox(
    filter_frame,
    textvariable=filter_var,
    values=[
        "Все",
        "Заготовки",
        "ВЕКТОР",
        "ИНТЕГРАЛ",
        "РЗП",
        "КНИИТМУ",
        "СВТ",
        "СИГНАЛ",
    ],
    state="readonly",
    width=15,
)
filter_combo.pack(side=LEFT, padx=5)

apply_btn = tb.Button(
    filter_frame, text="Применить", bootstyle="secondary", command=apply_request_filter
)
apply_btn.pack(side=LEFT, padx=5)

# Таблица
columns = ("Блок питания ССБ112", "Детектор ССБ161", "Колпак ССБ114")
request_table = tb.Treeview(
    request_panel, columns=columns, show="headings", height=25, bootstyle="light"
)
style = tb.Style()
style.configure(
    "Treeview",
    background=BG_WORK_AREA,
    foreground=TEXT_MAIN,
    fieldbackground=BG_WORK_AREA,
    font=FONT_PRIMARY,
)
style.configure(
    "Treeview.Heading", background=BG_HEADER, foreground=TEXT_HEADER, font=FONT_BOLD
)
style.map("Treeview", background=[("selected", BG_HEADER)])

for col in columns:
    request_table.heading(col, text=col)
    request_table.column(col, width=200, anchor=CENTER)

request_table.pack(fill=BOTH, expand=True)

request_scrollbar = tb.Scrollbar(
    request_panel,
    orient=VERTICAL,
    command=request_table.yview,
    style="Wide.Vertical.TScrollbar",
)
request_table.configure(yscrollcommand=request_scrollbar.set)
request_scrollbar.pack(side=RIGHT, fill=Y)

entry_frame_container = tb.Frame(frame_main, style="Main.TFrame")
entry_frame_container.pack(fill=X, pady=10, padx=10)
entry_header = tb.Frame(entry_frame_container, style="Header.TFrame")
entry_header.pack(fill="x", pady=(0, PADY_HEADER))
tb.Label(entry_header, text="Добавить СЕРП-ВС6Д", style="Header2.TLabel").pack(
    side="left", padx=10, pady=5
)


entry_panel = tb.Frame(entry_frame_container, style="Main.TFrame")
entry_panel.pack(fill=X)
input_row = tb.Frame(entry_panel, style="Main.TFrame")
input_row.pack(fill=X, pady=(0, 10))
input_frame = tb.Frame(input_row, style="Main.TFrame")
input_frame.pack(side=LEFT, fill=X, expand=True)
tb.Label(input_frame, text=TEXT_YEAR, font=FONT_PRIMARY, style="Main.TLabel").pack(
    side=LEFT, padx=5
)
year_var = tb.StringVar()
year_combo = tb.Combobox(
    input_frame, textvariable=year_var, font=FONT_PRIMARY, width=5, bootstyle="light"
)
year_combo["values"] = [f"{y:02}" for y in range(24, 100)]
year_combo.current(1)
year_combo.pack(side=LEFT, padx=5)
tb.Label(input_frame, text=TEXT_NUMBER, font=FONT_PRIMARY, style="Main.TLabel").pack(
    side=LEFT, padx=5
)
entry_var = tb.StringVar()
entry = tb.Entry(
    input_frame, textvariable=entry_var, font=FONT_PRIMARY, width=10, bootstyle="light"
)
entry.pack(side=LEFT, padx=5)
entry.bind("<Return>", lambda e: process_serial())
add_button = tb.Button(
    input_frame,
    text=TEXT_ADD_SERP,
    width=3,
    bootstyle="primary",
    padding=(5, 5),
    command=process_serial,
)
add_button.pack(side=LEFT, padx=5)

# --- СТРОКА 1 над списками: слева сортировка, справа — Сохранить ---
save_row = tb.Frame(entry_panel, style="Main.TFrame")
save_row.pack(fill=X, pady=(0, 10))  # между input_row и нижней строкой (поиском)

# левая "колонка" инструментов над списком
left_tools_col = tb.Frame(save_row, style="Main.TFrame")
left_tools_col.pack(side=LEFT, padx=5)

# Ряд 1: кнопки режимов
modes_row = tb.Frame(left_tools_col, style="Main.TFrame")
modes_row.pack(anchor="w")

btn_mode_drafts = tb.Button(
    modes_row,
    text="Заготовки",
    width=12,
    bootstyle="secondary",
    command=lambda: set_sort_mode("drafts"),
)
btn_mode_drafts.pack(side=LEFT, padx=3)
sort_buttons["drafts"] = btn_mode_drafts

btn_mode_blocks = tb.Button(
    modes_row,
    text="Блоки",
    width=10,
    bootstyle="secondary",
    command=lambda: set_sort_mode("blocks"),
)
btn_mode_blocks.pack(side=LEFT, padx=3)
sort_buttons["blocks"] = btn_mode_blocks

btn_mode_factories = tb.Button(
    modes_row,
    text="Заводы",
    width=10,
    bootstyle="secondary",
    command=lambda: set_sort_mode("factories"),
)
btn_mode_factories.pack(side=LEFT, padx=3)
sort_buttons["factories"] = btn_mode_factories

# Ряд 2: "🙂 112" + сортировка по дате
tools_row2 = tb.Frame(left_tools_col, style="Main.TFrame")
tools_row2.pack(anchor="w", pady=(4, 0))

btn_112 = tb.Button(
    tools_row2,
    text="🙂 112",
    width=10,
    bootstyle="secondary",
    command=lambda: toggle_112_icons(),
)
btn_112.pack(side="left", padx=3)

sort_button = tb.Button(
    tools_row2,
    text=("⏫" if sort_order == "new_first" else "⏬"),
    width=6,
    bootstyle="secondary",
    command=toggle_sort_order,
)
try:
    if sort_down_icon and sort_up_icon:
        sort_button.configure(
            image=(sort_down_icon if sort_order == "new_first" else sort_up_icon),
            compound="left",
        )
except Exception:
    pass
sort_button.pack(side="left", padx=3)


def toggle_112_icons():
    # TODO: сюда вставь свою реальную логику "кнопки 112"
    # (например, массовая подсветка up112_hints из файла и т.п.)
    messagebox.showinfo("112", "Кнопка 112 нажата (заглушка).")


# справа — кнопка Сохранить (как и было)
tb.Button(
    save_row,
    text=TEXT_SAVE,
    width=12,
    bootstyle="secondary",
    style="Save.TButton",
    command=lambda: save_data(True),
).pack(side="right", padx=5)

# --- СТРОКА 2: только ПОИСК справа ---
toolbar_row = tb.Frame(entry_panel, style="Main.TFrame")
toolbar_row.pack(fill=X, pady=(5, 10))

search_frame = tb.Frame(toolbar_row, style="Main.TFrame")
search_frame.pack(side=RIGHT, anchor="e")

search_var = tb.StringVar()
tb.Label(
    search_frame, text=TEXT_SEARCH_LABEL, font=FONT_PRIMARY, style="Main.TLabel"
).pack(side=LEFT, padx=5)

search_entry = tb.Entry(
    search_frame,
    textvariable=search_var,
    font=FONT_PRIMARY,
    width=20,
    bootstyle="light",
)
search_entry.pack(side=LEFT, padx=5)
search_entry.bind("<Return>", lambda e: search_product())

tb.Button(
    search_frame, text="🔍", width=4, bootstyle="secondary", command=search_product
).pack(side=LEFT)


main_area = tb.Frame(frame_main, style="Main.TFrame")
main_area.pack(fill=BOTH, expand=True, padx=10, pady=5)
main_area.columnconfigure(0, weight=5)
main_area.columnconfigure(1, weight=2)
main_area.rowconfigure(0, weight=1)

work_frame_container = tb.Frame(main_area, style="Main.TFrame")
work_frame_container.grid(row=0, column=0, sticky="nsew", padx=(0, 5), pady=0)
work_frame_container.grid_rowconfigure(1, weight=1)

work_header = tb.Frame(work_frame_container, style="Header.TFrame")
work_header.pack(fill="x", pady=(0, PADY_HEADER))
work_header_label = tb.Label(work_header, text=TEXT_WORK_TITLE, style="Header1.TLabel")
work_header_label.pack(side="left", padx=10, pady=5)
work_container = tb.Frame(work_frame_container, style="WorkArea.TFrame")
work_container.pack(fill="both", expand=True)

toggle_storage_button = tb.Button(
    work_header, text="▶", width=2, command=toggle_storage_visibility
)
toggle_storage_button.pack(side="right", padx=5, pady=5)

work_canvas = tb.Canvas(work_container, background=BG_WORK_AREA)
work_scrollbar = tb.Scrollbar(
    work_container,
    orient=VERTICAL,
    command=work_canvas.yview,
    bootstyle="dark-round",
    style="Wide.Vertical.TScrollbar",
)
work_frame = tb.Frame(work_canvas, style="WorkArea.TFrame")
work_frame.bind(
    "<Configure>", lambda e: work_canvas.configure(scrollregion=work_canvas.bbox("all"))
)
work_canvas.create_window((0, 0), window=work_frame, anchor=NW)
work_canvas.configure(yscrollcommand=work_scrollbar.set, background=BG_WORK_AREA)
work_scrollbar.pack(side=RIGHT, fill=Y)
work_canvas.pack(side=LEFT, fill=BOTH, expand=True)
work_canvas.config(height=300)

storage_frame_container = tb.Frame(main_area, style="Main.TFrame")
storage_frame_container.grid(row=0, column=1, sticky="nsew", padx=(5, 0), pady=0)
storage_frame_container.grid_rowconfigure(1, weight=1)
storage_header = tb.Frame(storage_frame_container, style="Header.TFrame")
storage_header_label = tb.Label(
    storage_header, text=TEXT_STORAGE_TITLE, style="Header1.TLabel"
)
storage_header.pack(fill="x", pady=(0, PADY_HEADER))
storage_header_label = tb.Label(
    storage_header, text=TEXT_STORAGE_TITLE, style="Header1.TLabel"
)
storage_header_label.pack(side="left", padx=10, pady=5)
storage_container = tb.Frame(storage_frame_container, style="WorkArea.TFrame")
storage_container.pack(fill="both", expand=True)
storage_canvas = tb.Canvas(storage_container, background=BG_WORK_AREA)
storage_scrollbar = tb.Scrollbar(
    storage_container,
    orient=VERTICAL,
    command=storage_canvas.yview,
    bootstyle="dark-round",
    style="Wide.Vertical.TScrollbar",
)
storage_frame = tb.Frame(storage_canvas, style="WorkArea.TFrame")
storage_frame.bind(
    "<Configure>",
    lambda e: storage_canvas.configure(scrollregion=storage_canvas.bbox("all")),
)
storage_canvas.create_window((0, 0), window=storage_frame, anchor=NW)
storage_canvas.configure(yscrollcommand=storage_scrollbar.set, background=BG_WORK_AREA)
storage_scrollbar.pack(side=RIGHT, fill=Y)
storage_canvas.pack(side=LEFT, fill=BOTH, expand=True)
storage_canvas.config(height=300)

archive_frame = tb.Frame(frame_assembly, style="Main.TFrame")
archive_frame.pack(fill=BOTH, expand=True, padx=10, pady=10)
factory_buttons = {}
dpi_var = None
scaling_var = None


def browse_data_path():
    path = filedialog.askdirectory()
    if path:
        data_path_var.set(path)


def apply_scaling():
    global scaling_factor
    new_scaling = scaling_var.get()
    scaling_factor = new_scaling
    save_data()
    messagebox.showinfo(
        "Масштабирование",
        "Масштабирование будет применено после перезапуска приложения.",
    )


def create_settings_tab():
    global data_path_var, scaling_var

    settings_container = tb.Frame(frame_settings, style="Settings.TFrame")
    settings_container.pack(fill=BOTH, expand=True, padx=20, pady=20)

    path_frame = tb.LabelFrame(
        settings_container,
        text="Путь сохранения данных",
        bootstyle="primary",
        padding=15,
        style="Settings.TLabelframe",
    )
    path_frame.pack(fill=X, pady=10)

    tb.Label(path_frame, text="Текущий путь сохранения:", style="Settings.TLabel").pack(
        anchor=W, pady=5
    )

    data_path_var = tb.StringVar(value=data_path)
    path_entry = tb.Entry(
        path_frame,
        textvariable=data_path_var,
        font=FONT_PRIMARY,
        bootstyle="light",
        state="readonly",
        width=50,
    )
    path_entry.pack(fill=X, padx=5, pady=5, expand=True)

    tb.Button(
        path_frame, text="Выбрать папку", bootstyle="primary", command=browse_data_path
    ).pack(side=LEFT, padx=5, pady=5)

    scaling_frame = tb.LabelFrame(
        settings_container,
        text="Масштабирование интерфейса",
        bootstyle="primary",
        padding=15,
        style="Settings.TLabelframe",
    )
    scaling_frame.pack(fill=X, pady=10)

    scale_value_label = tb.Label(
        scaling_frame,
        text=f"Текущее значение: {scaling_factor:.1f}",
        style="Settings.TLabel",
    )
    scale_value_label.pack(anchor=W, pady=5)

    tb.Label(
        scaling_frame,
        text="Коэффициент масштабирования (рекомендуется 1.5-2.5):",
        style="Settings.TLabel",
    ).pack(anchor=W, pady=5)

    scaling_var = tb.DoubleVar(value=scaling_factor)
    scale = tb.Scale(
        scaling_frame,
        from_=1.0,
        to=4.0,
        orient=tk.HORIZONTAL,
        variable=scaling_var,
        bootstyle="primary",
        command=lambda v: scale_value_label.config(
            text=f"Текущее значение: {float(v):.1f}"
        ),
    )
    scale.pack(fill=X, padx=5, pady=5)

    tb.Button(
        scaling_frame,
        text="Применить масштабирование",
        bootstyle="success",
        command=apply_scaling,
    ).pack(side=LEFT, anchor=W, padx=5, pady=5)

    factory_frame = tb.LabelFrame(
        settings_container,
        text="Порядок отображения заводов",
        bootstyle="primary",
        padding=15,
        style="Settings.TLabelframe",
    )
    factory_frame.pack(fill=X, pady=10)

    tb.Label(
        factory_frame,
        text="Нажмите на завод, чтобы сделать его первым в списке:",
        style="Settings.TLabel",
    ).pack(anchor=W, pady=5)

    btn_frame = tb.Frame(factory_frame, style="Settings.TFrame")
    btn_frame.pack(fill=X, pady=10)

    for factory in ["ВЕКТОР", "ИНТЕГРАЛ", "РЗП", "КНИИТМУ", "СВТ", "СИГНАЛ"]:
        btn = tb.Button(
            btn_frame,
            text=factory,
            width=12,
            bootstyle="primary-outline",
            command=lambda f=factory: move_factory_to_top(f),
        )
        btn.pack(side=LEFT, padx=5, pady=5)
        factory_buttons[factory] = btn

    reset_factory_order_btn = tb.Button(
        factory_frame, text="Сброс", bootstyle="warning", command=reset_factory_order
    )
    reset_factory_order_btn.pack(side=LEFT, anchor=W, padx=5, pady=10)

    reset_factory_order_btn.pack(pady=10)
    control_frame = tb.Frame(settings_container, style="Settings.TFrame")
    control_frame.pack(fill=X, pady=20)

    tb.Button(
        control_frame,
        text="Сохранить настройки",
        width=20,
        bootstyle="success",
        style="Bold.TButton",
        padding=5,
        command=save_settings,
    ).pack(side=RIGHT, padx=10, pady=10)

    tb.Label(
        settings_container,
        text="Версия программы: 2.6",
        style="Settings.TLabel",
        font=("Arial", 10, "italic"),
    ).pack(side=BOTTOM, anchor="e", padx=10, pady=5)


style.configure("Settings.TFrame", background=BG_SETTINGS)
style.configure(
    "Settings.TLabel", background=BG_SETTINGS, foreground=TEXT_MAIN, font=FONT_PRIMARY
)
style.configure(
    "Settings.TRadiobutton",
    background=BG_SETTINGS,
    foreground=TEXT_MAIN,
    font=FONT_PRIMARY,
)
style.configure("Settings.TLabelframe", background=BG_SETTINGS, foreground=TEXT_MAIN)
style.configure(
    "Settings.TLabelframe.Label",
    background=BG_SETTINGS,
    foreground=TEXT_MAIN,
    font=FONT_BOLD,
)

notebook.bind("<<NotebookTabChanged>>", on_tab_change)
root.bind("<MouseWheel>", _on_mousewheel)
root.bind("<Button-4>", _on_mousewheel)
root.bind("<Button-5>", _on_mousewheel)
create_settings_tab()
update_product_list()
update_assembly_archive()
update_button_colors()
root.withdraw()
root.update_idletasks()
center_window(initial=True)  # ставим в центр по запрошенному размеру — без мигания
collapse_storage_on_start()  # стартуем со свернутой колонкой
root.deiconify()  # показываем окно один раз — уже по центру


def on_closing():
    save_all_comments()
    save_data()
    root.destroy()


root.after(0, lambda: set_sort_mode("factories"))
root.protocol("WM_DELETE_WINDOW", on_closing)
# зафиксируем исходное состояние и подпишемся на изменения
last_window_state = root.state()
root.bind("<Configure>", on_window_state_change)
root.mainloop()
