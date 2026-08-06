# -*- coding: utf-8 -*-
"""
Ядро синхронизации тендеров из 1С с заметки.xlsx (лист "Тендеры").

Основной способ запуска — GUI: «Тендеры.bat» (tenders/app.py). Этот модуль
содержит всю логику (sync) и остаётся запускаемым из консоли для отладки:

    py -3.12 -m tenders.core [--dry-run] [--file ПУТЬ]

Как работает sync():
1. Получает текст буфера обмена (список тендеров из 1С, поля через табуляцию).
2. Сравнивает Номера с тем, что уже есть в Excel; НОВЫЕ тендеры получают статус
   "В работе", пометку "Добавлен ДД.ММ.ГГГГ" в столбце "Действие" и папку
   C:\\Работа\\В работе\\<Номер>\\С ЭТП с кликабельной ссылкой в столбце "Папка".
3. Пересортировывает весь лист: сначала по группе статуса (лист "Справочники"),
   затем по сроку подачи. Новые тендеры сами встают наверх — работа идёт сверху вниз.

Что НЕ трогается:
- значения и оформление существующих строк (Ссылка, Наименование, комментарии,
  Действие, Контакты, гиперссылки, примечания) сохраняются как есть. Меняется
  только положение строки на листе — по правилу сортировки.

Защита данных:
- перед сохранением оригинал копируется в подпапку backups (хранятся последние
  BACKUP_KEEP копий);
- запись атомарная: сначала во временный файл, затем замена оригинала — при
  сбое посреди записи оригинал остаётся цел.

Журналы и настройки (папка data в корне проекта):
- add_tenders.log — человекочитаемый итог каждого боевого запуска;
- history.jsonl — история запусков для статистики (по JSON-строке на запуск);
- config.json — путь к рабочему xlsx, выбранный в окне программы.

ВАЖНО: файл заметки.xlsx должен быть ЗАКРЫТ в момент запуска (кроме --dry-run).
"""

import argparse
import io
import json
import os
import re
import shutil
import sys
import datetime
import warnings
import zipfile
from copy import copy

import openpyxl
from openpyxl.comments import Comment
from openpyxl.formatting.formatting import ConditionalFormattingList
from openpyxl.styles import PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.formula import ArrayFormula

__version__ = "2.0.0"

# ------------------------- НАСТРОЙКИ -------------------------
WORK_DIR  = r"C:\Работа\В работе"
SHEET     = "Тендеры"        # рабочий лист
REF_SHEET = "Справочники"    # статусы и группы сортировки
TABLE     = "Тендеры_база"   # умная таблица на листе «Тендеры»
SUBFOLDER = "С ЭТП"

# Сколько последних резервных копий хранить в подпапке backups.
BACKUP_KEEP = 10

# Рабочие файлы программы (настройки, журнал, история) — в папке data корня проекта.
ROOT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_DIR = os.path.join(ROOT_DIR, "data")
LOG_PATH = os.path.join(DATA_DIR, "add_tenders.log")
HISTORY_PATH = os.path.join(DATA_DIR, "history.jsonl")

# Путь к рабочему xlsx: выбирается в окне программы и запоминается в config.json;
# пока не выбран — заметки.xlsx на Рабочем столе текущего пользователя.
CONFIG_PATH = os.path.join(DATA_DIR, "config.json")
DEFAULT_XLSX = os.path.join(os.path.expanduser("~"), "Desktop", "заметки.xlsx")

# Возможные значения статуса в выгрузке 1С — "якорь":
# поле сразу после статуса = ЭТП, следующее = описание.
# К списку статусов Excel (лист «Справочники») отношения не имеет.
STATUSES_1C = {"в работе", "подался", "отказ", "проиграли",
               "победа", "интересный", "интересные", "на согласовании"}

# Форматы даты "Окончание подачи" в выгрузке 1С (пробуем по порядку).
DATE_FORMATS = ("%d.%m.%Y %H:%M:%S", "%d.%m.%Y %H:%M", "%d.%m.%Y")

# Столбцы листа "Тендеры" (1=A ... 12=L)
COL_CUSTOMER, COL_NAME, COL_NUMBER, COL_LINK    = 1, 2, 3, 4
COL_ETP, COL_DEADLINE, COL_COMMENT, COL_ACTION  = 5, 6, 7, 8
COL_STATUS, COL_CONTACTS, COL_FOLDER, COL_SORT  = 9, 10, 11, 12
NCOL = 12

# Заголовки строки 1: макет фиксирован умной таблицей и сверяется перед записью.
HEADERS = ["Заказчик", "Наименование тендера", "Номер", "Ссылка", "ЭТП",
           "Окончание подачи", "Комментарий", "Действие", "Статус",
           "Контакты", "Папка", "Сорт."]

# Формула служебного столбца L "Сорт." — номер группы статуса из «Справочников».
SORT_FORMULA = ("=IFERROR(INDEX(Справочники!$E$2:$E$11,"
                "MATCH($I{row},Справочники!$C$2:$C$11,0)),99)")

# Лист «Справочники»: C — статусы в порядке сортировки, E — номер группы.
REF_STATUS_COL, REF_GROUP_COL = 3, 5
REF_ROW_FIRST, REF_ROW_LAST = 2, 11

DEFAULT_STATUS = "В работе"   # статус, который получают новые тендеры
UNKNOWN_GROUP  = 99           # группа для статуса, которого нет в справочнике
ROW_HEIGHT     = 30.0         # высота строки данных на листе «Тендеры»

# Формат "Окончание подачи": значения меньше 1 (пустые/нулевые) показываются пустыми.
DEADLINE_FORMAT = r'[<1]"";dd\.mm\.yyyy\ hh:mm'

FOLDER_TEXT   = "Открыть"     # текст ссылки в столбце "Папка"
ACTION_PREFIX = "Добавлен "   # пометка в столбце "Действие" у новых строк
# ------------------------------------------------------------


class SyncError(Exception):
    """Ожидаемая ошибка, текст которой показывается пользователю."""


def load_xlsx_path():
    """Путь к рабочему файлу: из config.json, иначе заметки.xlsx на Рабочем столе."""
    try:
        with open(CONFIG_PATH, encoding="utf-8") as f:
            path = json.load(f).get("xlsx_path")
    except (OSError, ValueError, AttributeError):
        path = None
    return path or DEFAULT_XLSX


def ensure_parent_dir(path):
    """Создаём папку для файла (data/), если её ещё нет."""
    parent = os.path.dirname(path)
    if parent:
        os.makedirs(parent, exist_ok=True)


def save_xlsx_path(path, say=print):
    """Запоминаем выбранный файл в config.json (папка data)."""
    try:
        ensure_parent_dir(CONFIG_PATH)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump({"xlsx_path": path}, f, ensure_ascii=False, indent=2)
    except OSError as ex:
        say("Не удалось сохранить настройки:", ex)


def fix_console():
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


def read_clipboard():
    """Читаем текст из буфера обмена встроенным tkinter (без внешних пакетов)."""
    import tkinter
    root = tkinter.Tk()
    root.withdraw()
    try:
        data = root.clipboard_get()
    except Exception:
        data = ""
    finally:
        root.destroy()
    return data


def sanitize_folder_name(name):
    bad = '<>:"/\\|?*'
    for ch in bad:
        name = name.replace(ch, "_")
    return name.strip().rstrip(".")


def norm_number(value):
    """Номер как строка: неразрывные пробелы и переводы строк -> обычный пробел,
    повторы пробелов схлопываются, края обрезаются — чтобы номер из 1С совпадал
    с номером, когда-то вбитым в Excel руками."""
    if value in (None, ""):
        return None
    s = re.sub(r"\s+", " ", str(value).replace("\xa0", " ")).strip()
    return s or None


def numbers_in_cell(value):
    """Номера тендеров из ячейки "Номер": в одной ячейке их бывает несколько,
    записанных через перевод строки."""
    if value in (None, ""):
        return []
    return [n for n in (norm_number(p) for p in str(value).split("\n")) if n]


def list_work_folders():
    """Список подпапок в WORK_DIR (имена)."""
    try:
        return [d for d in os.listdir(WORK_DIR)
                if os.path.isdir(os.path.join(WORK_DIR, d))]
    except FileNotFoundError:
        return []


def find_existing_folder(num, folders):
    """Ищем папку тендера по номеру: точное имя или имя вида '<номер> описание'.
    Возвращаем полный путь или None."""
    target = sanitize_folder_name(num)
    for d in folders:                       # точное совпадение
        if d == target:
            return os.path.join(WORK_DIR, d)
    for d in folders:                       # '<номер> ...' или '<номер>_...'
        if d.startswith(target + " ") or d.startswith(target + "_"):
            return os.path.join(WORK_DIR, d)
    return None


def set_folder_link(cell, path):
    """Кликабельная ссылка «Открыть» на папку тендера — абсолютным путём, как есть.

    Оформление ячейки не трогаем: у существующих строк оно своё, у новых приходит
    из строки-шаблона. Относительные пути с %20, которые видно в файле, пишет сам
    Excel, сокращая наши абсолютные при очередном ручном сохранении.
    """
    cell.value = FOLDER_TEXT
    cell.hyperlink = path


def ensure_folder(num, work_folders, created_list, say=print):
    """Возвращает путь к папке тендера, создавая её (с подпапкой С ЭТП) при отсутствии.
    В created_list добавляет путь, только если подпапка была реально создана."""
    path = find_existing_folder(num, work_folders)
    if path is None:
        path = os.path.join(WORK_DIR, sanitize_folder_name(num))
    sub = os.path.join(path, SUBFOLDER)
    newly = not os.path.isdir(sub)
    try:
        os.makedirs(sub, exist_ok=True)
    except Exception as ex:
        say("  не удалось создать папку для", num, "->", ex)
        return None
    bn = os.path.basename(path)
    if bn not in work_folders:
        work_folders.append(bn)
    if newly:
        created_list.append(path)
    return path


def parse_deadline(raw):
    """Дата из 1С -> datetime. Не распозналось — возвращаем исходную строку."""
    for fmt in DATE_FORMATS:
        try:
            return datetime.datetime.strptime(raw, fmt)
        except ValueError:
            pass
    return raw


def parse_row(line):
    """Разбираем строку из 1С (поля через TAB). Возвращаем dict или None."""
    fields = [f.strip() for f in line.split("\t")]
    if len(fields) < 6:
        return None

    number = norm_number(fields[4]) or ""
    customer = fields[5]
    if not number or number.lower() in ("номер", "да", "нет") or number.lower() in STATUSES_1C:
        return None

    deadline_raw = fields[2] if len(fields) > 2 else ""

    etp, name = "", ""
    status_idx = None
    for i, f in enumerate(fields):
        if f.lower() in STATUSES_1C:
            status_idx = i
            break
    if status_idx is not None:
        if status_idx + 1 < len(fields):
            etp = fields[status_idx + 1]
        if status_idx + 2 < len(fields):
            name = fields[status_idx + 2]
    else:
        if len(fields) > 10:
            etp = fields[10]
        if len(fields) > 11:
            name = fields[11]

    # name — описание тендера из 1С, едет в столбец B "Наименование тендера".
    return {"number": number, "customer": customer, "etp": etp,
            "deadline": parse_deadline(deadline_raw), "name": name}


def check_layout(ws):
    """Сверяем заголовки: sync() стирает всю область данных, и запись при
    разъехавшихся столбцах испортила бы файл."""
    actual = [str(ws.cell(row=1, column=c).value or "").strip() for c in range(1, NCOL + 1)]
    if actual != HEADERS:
        raise SyncError("Не совпадают заголовки листа «%s» — файл не тронут.\n"
                        "Ожидается: %s\nВ файле:   %s"
                        % (SHEET, " | ".join(HEADERS), " | ".join(actual)))


def last_data_row(ws):
    """Последняя строка данных: по границе умной таблицы, иначе — сканом."""
    table = ws.tables.get(TABLE)
    if table is not None:
        m = re.search(r"(\d+)$", table.ref)
        if m:
            return int(m.group(1))
    last = 1
    for r in range(2, ws.max_row + 1):
        if (ws.cell(row=r, column=COL_CUSTOMER).value not in (None, "") or
                ws.cell(row=r, column=COL_NUMBER).value not in (None, "")):
            last = r
    return last


def load_groups(wb, say=print):
    """Порядок сортировки с листа «Справочники»: {статус в нижнем регистре: группа}."""
    groups = {}
    if REF_SHEET in wb.sheetnames:
        ws = wb[REF_SHEET]
        for r in range(REF_ROW_FIRST, REF_ROW_LAST + 1):
            status = ws.cell(row=r, column=REF_STATUS_COL).value
            group = ws.cell(row=r, column=REF_GROUP_COL).value
            if status and isinstance(group, (int, float)) and not isinstance(group, bool):
                groups[str(status).strip().lower()] = int(group)
    if not groups:
        say("Лист «%s» не найден или пуст — сортировка только по сроку подачи." % REF_SHEET)
    return groups


# Строки без срока и с нераспознанным сроком уезжают в конец своей группы.
NO_DEADLINE = 10 ** 9
EXCEL_EPOCH = datetime.datetime(1899, 12, 30)


def deadline_key(value):
    """Срок подачи -> число для сортировки."""
    if isinstance(value, datetime.datetime):
        return (value - EXCEL_EPOCH).total_seconds() / 86400.0
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        return float(value)
    return NO_DEADLINE


def snapshot_cell(cell):
    """Полный снимок ячейки: значение + стиль + гиперссылка + комментарий."""
    cm = None
    if cell.comment is not None:
        cm = (cell.comment.text, cell.comment.author)
    return {
        "value": cell.value,
        "number_format": cell.number_format,
        "font": copy(cell.font),
        "fill": copy(cell.fill),
        "border": copy(cell.border),
        "alignment": copy(cell.alignment),
        "protection": copy(cell.protection),
        # Объект целиком, а не .target: у части ссылок адрес лежит в location.
        "hyperlink": copy(cell.hyperlink),
        "comment": cm,
    }


def apply_snapshot(cell, s):
    cell.value = s["value"]
    cell.number_format = s["number_format"]
    cell.font = s["font"]
    cell.fill = s["fill"]
    cell.border = s["border"]
    cell.alignment = s["alignment"]
    cell.protection = s["protection"]
    if s["hyperlink"] is not None:       # setter сам поправит ref на новую строку
        cell.hyperlink = s["hyperlink"]
    cell.comment = Comment(*s["comment"]) if s["comment"] else None


def apply_style_only(cell, s):
    """Применяем из снимка только оформление, не трогая значение/ссылку/комментарий."""
    cell.number_format = s["number_format"]
    cell.font = copy(s["font"])
    cell.fill = copy(s["fill"])
    cell.border = copy(s["border"])
    cell.alignment = copy(s["alignment"])
    cell.protection = copy(s["protection"])


def clear_cell(cell):
    cell.value = None
    cell.hyperlink = None
    cell.comment = None
    cell.style = "Normal"


def build_template(rows):
    """Оформление новой строки: по каждому столбцу берём первую строку данных
    БЕЗ заливки — иначе утащим ручную «зебру» и цвета условного форматирования."""
    template = []
    for c in range(1, NCOL + 1):
        pick = None
        for row in rows:
            fill = row["cells"][c - 1]["fill"]
            if fill is None or fill.patternType is None:
                pick = row["cells"][c - 1]
                break
        template.append(pick if pick is not None else rows[0]["cells"][c - 1])
    return template


# --- Выпадающий список статусов -------------------------------------------
# Excel хранит его в расширении x14 внутри <extLst> листа. openpyxl это
# расширение не понимает и при сохранении вырезает — возвращаем его вручную.

def read_ext_lst(path, sheet_index):
    """Достаём <extLst> листа из исходного xlsx (bytes; b"" — если его нет)."""
    try:
        with zipfile.ZipFile(path) as z:
            xml = z.read("xl/worksheets/sheet%d.xml" % sheet_index)
    except (OSError, KeyError, zipfile.BadZipFile):
        return b""
    m = re.search(rb"<extLst>.*?</extLst>(?=</worksheet>)", xml, re.S)
    if m is None:
        return b""
    # xr:uid ссылается на пространство имён, которого в файле openpyxl нет:
    # без удаления Excel посчитает файл повреждённым.
    return re.sub(rb' xr:uid="[^"]*"', b"", m.group())


def write_ext_lst(path, sheet_index, ext, last_row):
    """Вшиваем <extLst> обратно в сохранённый xlsx (архив пересобирается)."""
    if not ext:
        return
    # Диапазон выпадающего списка должен покрывать все строки таблицы с запасом.
    sqref = "<xm:sqref>I2:I%d</xm:sqref>" % max(1994, last_row + 200)
    ext = re.sub(rb"<xm:sqref>I2:I\d+</xm:sqref>", sqref.encode(), ext)
    target = "xl/worksheets/sheet%d.xml" % sheet_index
    with zipfile.ZipFile(path) as zin:
        parts = [(i.filename, zin.read(i.filename)) for i in zin.infolist()]
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, data in parts:
            if name == target:
                data = data.replace(b"</worksheet>", ext + b"</worksheet>")
            zout.writestr(name, data)
    with open(path, "wb") as f:
        f.write(buf.getvalue())


# Жёстко зашитые в формулах других листов границы диапазона листа «Тендеры».
TENDER_RANGE_RE = re.compile(r"(Тендеры!\$[A-Z]{1,3}\$\d+:\$[A-Z]{1,3}\$)(\d+)")


def widen_formulas(wb, last, say=print):
    """Раздвигаем диапазоны Тендеры!...$749 в формулах «Дашборда» и листов-витрин:
    иначе после роста таблицы они молча перестанут учитывать новые строки."""
    changed = 0
    for name in wb.sheetnames:
        if name == SHEET:
            continue
        for row in wb[name].iter_rows():
            for cell in row:
                value = cell.value
                if isinstance(value, ArrayFormula):
                    new = TENDER_RANGE_RE.sub(lambda m: m.group(1) + str(last), value.text)
                    if new != value.text:
                        cell.value = ArrayFormula(value.ref, new)
                        changed += 1
                elif isinstance(value, str) and value.startswith("="):
                    new = TENDER_RANGE_RE.sub(lambda m: m.group(1) + str(last), value)
                    if new != value:
                        cell.value = new
                        changed += 1
    if changed:
        say("Диапазоны формул на других листах расширены до строки %d (формул: %d)"
            % (last, changed))


def make_backup(path):
    """Копия xlsx в подпапку backups рядом с файлом; храним BACKUP_KEEP последних."""
    bdir = os.path.join(os.path.dirname(path) or ".", "backups")
    os.makedirs(bdir, exist_ok=True)
    base, ext = os.path.splitext(os.path.basename(path))
    stamp = datetime.datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    dst = os.path.join(bdir, base + "_" + stamp + ext)
    shutil.copy2(path, dst)
    old = sorted(f for f in os.listdir(bdir)
                 if f.startswith(base + "_") and f.endswith(ext))
    for f in old[:-BACKUP_KEEP]:
        try:
            os.remove(os.path.join(bdir, f))
        except OSError:
            pass
    return dst


def remove_quiet(path):
    try:
        os.remove(path)
    except OSError:
        pass


def append_log(lines, say=print):
    """Дописываем итог запуска в add_tenders.log."""
    try:
        ensure_parent_dir(LOG_PATH)
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write("\n".join(lines) + "\n\n")
    except OSError as ex:
        say("Не удалось записать лог:", ex)


# ------------------------- ИСТОРИЯ / СТАТИСТИКА -------------------------

def record_history(numbers, xlsx_path, say=print):
    """Дописываем запись об успешном запуске в history.jsonl (для статистики)."""
    rec = {"ts": datetime.datetime.now().isoformat(timespec="seconds"),
           "added": len(numbers), "numbers": list(numbers), "file": xlsx_path}
    try:
        ensure_parent_dir(HISTORY_PATH)
        with open(HISTORY_PATH, "a", encoding="utf-8") as f:
            f.write(json.dumps(rec, ensure_ascii=False) + "\n")
    except OSError as ex:
        say("Не удалось записать историю:", ex)


def load_history(path=None):
    """Читаем history.jsonl; битые строки молча пропускаем."""
    if path is None:
        path = HISTORY_PATH
    records = []
    try:
        with open(path, encoding="utf-8") as f:
            for ln in f:
                ln = ln.strip()
                if not ln:
                    continue
                try:
                    rec = json.loads(ln)
                except ValueError:
                    continue
                if isinstance(rec, dict) and "ts" in rec:
                    records.append(rec)
    except OSError:
        pass
    return records


def counts_by_day(records):
    """Свод по дням: {date: {"count": сколько добавлено, "numbers": [...]}}."""
    days = {}
    for rec in records:
        try:
            day = datetime.datetime.fromisoformat(rec["ts"]).date()
        except (KeyError, TypeError, ValueError):
            continue
        d = days.setdefault(day, {"count": 0, "numbers": []})
        d["count"] += int(rec.get("added", 0) or 0)
        d["numbers"].extend(rec.get("numbers", []))
    return days


# ------------------------- ОСНОВНАЯ ЛОГИКА -------------------------

def sync(xlsx_path, text, dry_run=False, say=print):
    """Синхронизация: текст буфера 1С -> xlsx + папки.

    Возвращает dict с итогами; ожидаемые проблемы поднимает как SyncError
    (текст готов для показа пользователю). say(*args) — вывод хода работы.
    """
    if not os.path.exists(xlsx_path):
        raise SyncError("Не найден файл " + xlsx_path)

    # В режиме предпросмотра ничего не пишем, поэтому открытый Excel не мешает.
    if not dry_run:
        try:
            with open(xlsx_path, "r+b"):
                pass
        except PermissionError:
            raise SyncError("Файл заметки.xlsx сейчас ОТКРЫТ в Excel.\n"
                            "Закройте его и запустите снова.")

    lines = [ln for ln in (text or "").replace("\r\n", "\n").split("\n") if ln.strip()]
    if not lines:
        raise SyncError("Буфер обмена пуст. Скопируйте список из 1С и повторите.")

    # Разбор буфера, сохраняя порядок 1С, без повторов внутри буфера
    entries, seen_clip = [], set()
    for ln in lines:
        e = parse_row(ln)
        if not e:
            continue
        if e["number"] in seen_clip:
            continue
        seen_clip.add(e["number"])
        entries.append(e)

    if not entries:
        raise SyncError("Не удалось распознать строки тендеров в буфере обмена.\n"
                        "Проверьте, что скопированы строки из 1С (поля через табуляцию).")

    # openpyxl предупреждает, что не понимает x14-валидацию: возвращаем её сами
    # (write_ext_lst), а предупреждение в окно программы пускать незачем.
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        wb = openpyxl.load_workbook(xlsx_path)
    if SHEET not in wb.sheetnames:
        raise SyncError("В файле нет листа «%s».\n"
                        "Похоже, выбран старый файл заметок." % SHEET)
    ws = wb[SHEET]
    check_layout(ws)
    ext = read_ext_lst(xlsx_path, wb.sheetnames.index(SHEET) + 1)

    last = last_data_row(ws)
    groups = load_groups(wb, say)

    # Снимок существующих строк данных
    existing = []
    for r in range(2, last + 1):
        cells = [snapshot_cell(ws.cell(r, c)) for c in range(1, NCOL + 1)]
        existing.append({
            "numbers": numbers_in_cell(ws.cell(r, COL_NUMBER).value),
            "status": ws.cell(r, COL_STATUS).value,
            "deadline": ws.cell(r, COL_DEADLINE).value,
            "height": ws.row_dimensions[r].height,
            "cells": cells,
        })

    existing_numbers = set()
    for row in existing:
        existing_numbers.update(row["numbers"])

    new_entries = [e for e in entries if e["number"] not in existing_numbers]
    new_numbers = [e["number"] for e in new_entries]
    existed = len(entries) - len(new_entries)

    # Итоговый порядок строк: группа статуса, затем срок подачи.
    # Третий элемент ключа — исходный порядок: при равных группе и сроке
    # существующие строки остаются на местах, а новые встают следом за ними.
    default_group = groups.get(DEFAULT_STATUS.lower(), UNKNOWN_GROUP)
    final = []
    for i, row in enumerate(existing):
        status = str(row["status"] or "").strip().lower()
        final.append((groups.get(status, UNKNOWN_GROUP),
                      deadline_key(row["deadline"]), i, "existing", row))
    for i, e in enumerate(new_entries):
        final.append((default_group, deadline_key(e["deadline"]),
                      10 ** 6 + i, "new", e))
    final.sort(key=lambda item: item[:3])

    write_last = 1 + len(final)
    row_of = {}                 # номер тендера -> строка, куда он встанет
    for pos, item in enumerate(final):
        if item[3] == "new":
            row_of[item[4]["number"]] = pos + 2

    if dry_run:
        say("[ПРЕДПРОСМОТР] Файл и папки не изменяются.")
        say("Новых тендеров:", len(new_entries))
        for e in new_entries:
            say("   +", e["number"], "—", e["customer"],
                "(встанет строкой %d)" % row_of[e["number"]])
        say("Уже было в файле:", existed)
        folders = list_work_folders()
        to_create = []
        for e in entries:
            path = find_existing_folder(e["number"], folders)
            if path is None:
                to_create.append(os.path.join(
                    WORK_DIR, sanitize_folder_name(e["number"]), SUBFOLDER))
            elif not os.path.isdir(os.path.join(path, SUBFOLDER)):
                to_create.append(os.path.join(path, SUBFOLDER))
        if to_create:
            say("Будут созданы папки:")
            for p in to_create:
                say("   ", p)
        return {"dry_run": True, "new": new_numbers, "existed": existed,
                "folders_to_create": to_create, "created_folders": [],
                "linked": 0, "backup": None}

    # Чистим всю область данных, затем пишем заново в новом порядке
    for r in range(2, max(last, write_last) + 1):
        for c in range(1, NCOL + 1):
            clear_cell(ws.cell(r, c))

    template = build_template(existing) if existing else None
    clip_numbers = {e["number"] for e in entries}
    work_folders = list_work_folders()
    created_folders = []
    mark = ACTION_PREFIX + datetime.date.today().strftime("%d.%m.%Y")
    linked = 0
    r = 2
    for _, _, _, kind, src in final:
        if kind == "existing":
            for c in range(1, NCOL + 1):
                apply_snapshot(ws.cell(r, c), src["cells"][c - 1])
            if src["height"] is not None:
                ws.row_dimensions[r].height = src["height"]
            for num in src["numbers"]:
                # Тендер из текущего списка 1С — создаём папку, если её нет.
                # Старые строки (нет в списке) — только ссылка, если папка уже есть.
                if num in clip_numbers:
                    path = ensure_folder(num, work_folders, created_folders, say)
                else:
                    path = find_existing_folder(num, work_folders)
                if path:
                    set_folder_link(ws.cell(r, COL_FOLDER), path)
                    linked += 1
                    break
        else:
            e = src
            path = ensure_folder(e["number"], work_folders, created_folders, say)
            if template is not None:
                for c in range(1, NCOL + 1):
                    apply_style_only(ws.cell(r, c), template[c - 1])
                    # Заливку не наследуем: фон белый, цвет даёт условное форматирование.
                    ws.cell(r, c).fill = PatternFill()
            ws.cell(r, COL_CUSTOMER, e["customer"])
            ws.cell(r, COL_NAME, e["name"])
            ws.cell(r, COL_NUMBER, e["number"])
            ws.cell(r, COL_ETP, e["etp"])
            ws.cell(r, COL_STATUS, DEFAULT_STATUS)
            ws.cell(r, COL_ACTION, mark)
            if e["deadline"] not in (None, ""):
                dcell = ws.cell(r, COL_DEADLINE, e["deadline"])
                # Формат ТОЛЬКО после значения: иначе openpyxl затрёт его своим.
                dcell.number_format = DEADLINE_FORMAT
            ws.row_dimensions[r].height = ROW_HEIGHT
            if path:
                set_folder_link(ws.cell(r, COL_FOLDER), path)
                linked += 1
        # Служебный столбец «Сорт.» — вычисляемый, ручных данных в нём нет.
        ws.cell(r, COL_SORT, SORT_FORMULA.format(row=r))
        r += 1

    # Раздвигаем границы умной таблицы, автофильтра и сохранённой сортировки.
    table = ws.tables.get(TABLE)
    if table is not None:
        table.ref = "A1:%s%d" % (get_column_letter(NCOL), write_last)
        if table.autoFilter is not None:
            table.autoFilter.ref = table.ref
        if table.sortState is not None:
            table.sortState.ref = "A2:%s%d" % (get_column_letter(NCOL), write_last)

    # Диапазоны условного форматирования. Пересобираем список целиком: присваивание
    # sqref на месте ломает сохранение (ключ словаря правил хэшируется по нему).
    saved = [(str(fmt.sqref), list(fmt.rules)) for fmt in ws.conditional_formatting]
    ws.conditional_formatting = ConditionalFormattingList()
    for sqref, rules in saved:
        # Подсветка срока — по столбцу F, окраска строки по статусу — по A:K.
        wide = ("F2:F%d" % write_last if sqref.startswith("F2")
                else "A2:%s%d" % (get_column_letter(COL_FOLDER), write_last))
        for rule in rules:
            ws.conditional_formatting.add(wide, rule)

    widen_formulas(wb, write_last, say)

    unknown = sorted({str(row["status"]).strip() for row in existing
                      if row["status"]
                      and str(row["status"]).strip().lower() not in groups})
    if unknown:
        say("Статусы вне справочника (такие строки уходят вниз):", ", ".join(unknown))

    # Резервная копия оригинала; без неё файл не трогаем.
    try:
        backup = make_backup(xlsx_path)
    except OSError as ex:
        raise SyncError("Не удалось создать резервную копию: %s\n"
                        "Файл НЕ изменён. Проверьте место на диске/права и повторите." % ex)

    # Атомарное сохранение: пишем во временный файл, затем заменяем оригинал.
    tmp_path = xlsx_path + ".tmp"
    try:
        wb.save(tmp_path)
        # Возвращаем в файл выпадающий список статусов — до подмены оригинала.
        write_ext_lst(tmp_path, wb.sheetnames.index(SHEET) + 1, ext, write_last)
        os.replace(tmp_path, xlsx_path)
    except PermissionError:
        remove_quiet(tmp_path)
        raise SyncError("Не удалось сохранить: файл открыт в Excel. Закройте его и повторите.")
    except Exception:
        remove_quiet(tmp_path)
        raise

    say("-" * 50)
    say("Добавлено новых тендеров:", len(new_entries))
    say("Уже было в файле:", existed)
    if new_entries:
        say("Новые номера:", ", ".join(new_numbers))
    say("Ссылок на папки проставлено (столбец %s): %d"
        % (get_column_letter(COL_FOLDER), linked))
    if created_folders:
        say("Созданы папки:")
        for f in created_folders:
            say("   ", f)
    say("Резервная копия:", backup)
    say("Готово. Файл сохранён:", xlsx_path)

    log = ["[%s] %s" % (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), xlsx_path),
           "добавлено: %d, уже было: %d" % (len(new_entries), existed)]
    if new_entries:
        log.append("новые: " + ", ".join(new_numbers))
    if created_folders:
        log.append("созданы папки: " + "; ".join(created_folders))
    log.append("бэкап: " + backup)
    append_log(log, say)
    record_history(new_numbers, xlsx_path, say)

    return {"dry_run": False, "new": new_numbers, "existed": existed,
            "created_folders": created_folders, "linked": linked, "backup": backup}


def main():
    fix_console()

    ap = argparse.ArgumentParser(
        description="Добавление тендеров из буфера обмена (1С) в заметки.xlsx")
    ap.add_argument("--dry-run", action="store_true",
                    help="показать, что будет добавлено, без записи файла и создания папок")
    ap.add_argument("--file", metavar="ПУТЬ",
                    help="путь к xlsx (по умолчанию — выбранный в окне программы)")
    args = ap.parse_args()

    try:
        sync(args.file or load_xlsx_path(), read_clipboard(), dry_run=args.dry_run)
    except SyncError as ex:
        print(ex)


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        import traceback
        traceback.print_exc()
        print("ОШИБКА:", e)
