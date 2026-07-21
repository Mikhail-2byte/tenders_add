# -*- coding: utf-8 -*-
"""Тесты ядра sync(): вставка новых тендеров и сохранность существующих данных.

Это самая рискованная часть программы — она стирает всю область данных листа и
пишет её заново, поэтому проверяем не только «добавилось», но и «ничего не пропало».

Запуск: py -3.12 -m pytest
"""
import datetime
import os

import openpyxl
import pytest
from openpyxl.comments import Comment
from openpyxl.styles import Font

import add_tenders as at


HEADERS = ["Заказчик", "Коментарий", "Номер", "Ссылка", "ЭТП", "Окончание подачи"]


def line_1c(number, customer="ООО Ромашка", deadline="10.08.2026 12:30",
            status="В работе", etp="Сбербанк-АСТ", comment="комментарий"):
    """Строка выгрузки 1С: 12 полей через TAB, номер — поле 5, заказчик — поле 6."""
    return "\t".join(["x", "Да", deadline, "x", number, customer,
                      "x", "x", "x", status, etp, comment])


@pytest.fixture
def env(tmp_path, monkeypatch):
    """Временный xlsx + временная папка работ; реальные файлы не трогаются."""
    work = tmp_path / "В работе"
    work.mkdir()
    monkeypatch.setattr(at, "WORK_DIR", str(work))
    monkeypatch.setattr(at, "HISTORY_PATH", str(tmp_path / "history.jsonl"))
    monkeypatch.setattr(at, "LOG_PATH", str(tmp_path / "add_tenders.log"))

    xlsx = tmp_path / "заметки.xlsx"
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = at.SHEET
    ws.append(HEADERS)
    ws.append(["Первый заказчик", "старый коммент", "111",
               "https://example.com/1", "b2b-center", "01.07.2026"])
    ws.append(["Второй заказчик", "ручная пометка", "222",
               "https://example.com/2", "rts-tender", "02.07.2026"])
    # Ручное оформление и данные, которые обязаны пережить перезапись листа.
    ws["D3"].hyperlink = "https://example.com/2"
    ws["D3"].font = Font(bold=True, color="FF0000")
    ws["B3"].comment = Comment("важное примечание", "менеджер")
    wb.save(xlsx)
    return {"xlsx": str(xlsx), "work": str(work), "tmp": tmp_path}


def load(env):
    return openpyxl.load_workbook(env["xlsx"])[at.SHEET]


def numbers(ws):
    """Номера тендеров сверху вниз (столбец C)."""
    out = []
    for r in range(2, ws.max_row + 1):
        v = ws.cell(r, at.COL_NUMBER).value
        if v not in (None, ""):
            out.append(str(v))
    return out


def quiet(*_):
    """Заглушка вместо print — тесты не должны шуметь в выводе."""


# ------------------------- вставка и порядок -------------------------

def test_new_tender_inserted_after_its_anchor(env):
    """Новый тендер встаёт сразу после предыдущего тендера из списка 1С."""
    text = "\n".join([line_1c("111"), line_1c("999"), line_1c("222")])
    result = at.sync(env["xlsx"], text, say=quiet)

    assert result["new"] == ["999"]
    assert result["existed"] == 2
    assert numbers(load(env)) == ["111", "999", "222"]


def test_new_tender_fields_written(env):
    text = line_1c("999", customer="Новый заказчик", etp="roseltorg",
                   comment="срочно", deadline="10.08.2026 12:30")
    at.sync(env["xlsx"], text, say=quiet)

    ws = load(env)
    row = 2 + numbers(ws).index("999")
    assert ws.cell(row, at.COL_CUSTOMER).value == "Новый заказчик"
    assert ws.cell(row, at.COL_ETP).value == "roseltorg"
    assert ws.cell(row, at.COL_COMMENT).value == "срочно"
    assert ws.cell(row, at.COL_DEADLINE).value == datetime.datetime(2026, 8, 10, 12, 30)


def test_existing_rows_survive_rewrite(env):
    """Ручные данные, стиль, гиперссылка и примечание старых строк не теряются."""
    at.sync(env["xlsx"], line_1c("999"), say=quiet)

    ws = load(env)
    row = 2 + numbers(ws).index("222")
    assert ws.cell(row, at.COL_CUSTOMER).value == "Второй заказчик"
    assert ws.cell(row, at.COL_COMMENT).value == "ручная пометка"
    assert ws.cell(row, 4).value == "https://example.com/2"
    assert ws.cell(row, 4).hyperlink.target == "https://example.com/2"
    assert ws.cell(row, 4).font.bold is True
    assert "важное примечание" in ws.cell(row, at.COL_COMMENT).comment.text


def test_second_run_adds_nothing(env):
    """Повторный запуск с тем же буфером не дублирует строки."""
    text = "\n".join([line_1c("111"), line_1c("999")])
    at.sync(env["xlsx"], text, say=quiet)
    before = numbers(load(env))

    result = at.sync(env["xlsx"], text, say=quiet)

    assert result["new"] == []
    assert numbers(load(env)) == before


# ------------------------- папки и ссылки -------------------------

def test_folder_created_and_linked(env):
    at.sync(env["xlsx"], line_1c("999"), say=quiet)

    assert os.path.isdir(os.path.join(env["work"], "999", at.SUBFOLDER))
    ws = load(env)
    folder_col = len(HEADERS) + 1        # «Папка» создаётся после последнего заголовка
    assert ws.cell(1, folder_col).value == at.FOLDER_HEADER
    row = 2 + numbers(ws).index("999")
    assert ws.cell(row, folder_col).value == "Открыть"
    assert ws.cell(row, folder_col).hyperlink.target.endswith("999")


def test_existing_folder_with_suffix_reused(env):
    """Папка «999 описание» считается папкой тендера 999 — новая не создаётся."""
    existing = os.path.join(env["work"], "999 поставка насосов")
    os.makedirs(existing)
    at.sync(env["xlsx"], line_1c("999"), say=quiet)

    assert not os.path.exists(os.path.join(env["work"], "999"))
    assert os.path.isdir(os.path.join(existing, at.SUBFOLDER))


# ------------------------- подсветка -------------------------

def test_highlight_only_on_current_run(env):
    """Зелёным залиты только тендеры текущего запуска, прошлые — очищаются."""
    at.sync(env["xlsx"], line_1c("999"), say=quiet)
    ws = load(env)
    row999 = 2 + numbers(ws).index("999")
    assert at.is_highlight_fill(ws.cell(row999, at.COL_CUSTOMER).fill)

    at.sync(env["xlsx"], line_1c("888"), say=quiet)
    ws = load(env)
    row999 = 2 + numbers(ws).index("999")
    row888 = 2 + numbers(ws).index("888")
    assert not at.is_highlight_fill(ws.cell(row999, at.COL_CUSTOMER).fill)
    assert at.is_highlight_fill(ws.cell(row888, at.COL_CUSTOMER).fill)


# ------------------------- защита данных -------------------------

def test_backup_created_and_no_tmp_left(env):
    at.sync(env["xlsx"], line_1c("999"), say=quiet)

    backups = list((env["tmp"] / "backups").iterdir())
    assert len(backups) == 1
    assert openpyxl.load_workbook(backups[0])[at.SHEET].max_row == 3   # копия ДО правки
    assert not os.path.exists(env["xlsx"] + ".tmp")


def test_dry_run_changes_nothing(env):
    before = open(env["xlsx"], "rb").read()

    result = at.sync(env["xlsx"], line_1c("999"), dry_run=True, say=quiet)

    assert result["new"] == ["999"]
    assert open(env["xlsx"], "rb").read() == before
    assert not os.path.exists(os.path.join(env["work"], "999"))
    assert not os.path.exists(at.HISTORY_PATH)


def test_history_written_for_real_run(env):
    at.sync(env["xlsx"], line_1c("999"), say=quiet)

    records = at.load_history(at.HISTORY_PATH)
    assert len(records) == 1
    assert records[0]["numbers"] == ["999"]


# ------------------------- ожидаемые ошибки -------------------------

def test_missing_file(tmp_path):
    with pytest.raises(at.SyncError):
        at.sync(str(tmp_path / "нет.xlsx"), line_1c("999"), say=quiet)


def test_empty_clipboard(env):
    with pytest.raises(at.SyncError):
        at.sync(env["xlsx"], "   \n\n", say=quiet)


def test_unrecognized_clipboard(env):
    with pytest.raises(at.SyncError):
        at.sync(env["xlsx"], "просто текст без табуляций", say=quiet)


def test_missing_sheet(env, monkeypatch):
    monkeypatch.setattr(at, "SHEET", "Такого листа нет")
    with pytest.raises(at.SyncError):
        at.sync(env["xlsx"], line_1c("999"), say=quiet)
