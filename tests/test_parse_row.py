# -*- coding: utf-8 -*-
"""Тесты разбора выгрузки 1С. Запуск: py -3.12 -m pytest test_parse_row.py"""
import datetime

from tenders.core import parse_row, norm_number, numbers_in_cell


def make_line(number="0123456789", customer="ООО Ромашка", deadline="10.08.2026",
              status="В работе", etp="Сбербанк-АСТ", name="наименование"):
    """Типовая строка 1С: 12 полей через TAB, статус на позиции 9."""
    fields = ["x", "Да", deadline, "x", number, customer,
              "x", "x", "x", status, etp, name]
    return "\t".join(fields)


def test_typical_line():
    e = parse_row(make_line())
    assert e["number"] == "0123456789"
    assert e["customer"] == "ООО Ромашка"
    assert e["etp"] == "Сбербанк-АСТ"
    # Описание из 1С едет в столбец B «Наименование тендера».
    assert e["name"] == "наименование"
    assert e["deadline"] == datetime.datetime(2026, 8, 10)


def test_status_anchor_in_other_position():
    # Статус на позиции 7 — ЭТП и наименование берутся сразу после него.
    fields = ["x", "Да", "10.08.2026", "x", "111", "Заказчик",
              "x", "Подался", "РТС-тендер", "срочно"]
    e = parse_row("\t".join(fields))
    assert e["etp"] == "РТС-тендер"
    assert e["name"] == "срочно"


def test_fallback_without_status():
    # Статуса нет — ЭТП и наименование по фиксированным индексам 10/11.
    fields = ["x", "Да", "10.08.2026", "x", "222", "Заказчик",
              "x", "x", "x", "x", "ЭТП-фоллбек", "наименование-фоллбек"]
    e = parse_row("\t".join(fields))
    assert e["etp"] == "ЭТП-фоллбек"
    assert e["name"] == "наименование-фоллбек"


def test_header_line_skipped():
    assert parse_row(make_line(number="Номер")) is None
    assert parse_row(make_line(number="Да")) is None
    assert parse_row(make_line(number="Отказ")) is None


def test_short_line_skipped():
    assert parse_row("одно поле") is None
    assert parse_row("a\tb\tc\td\te") is None


def test_empty_number_skipped():
    assert parse_row(make_line(number="")) is None


def test_deadline_with_time():
    e = parse_row(make_line(deadline="10.08.2026 12:30:00"))
    assert e["deadline"] == datetime.datetime(2026, 8, 10, 12, 30)
    e = parse_row(make_line(deadline="10.08.2026 09:00"))
    assert e["deadline"] == datetime.datetime(2026, 8, 10, 9, 0)


def test_deadline_unparsed_kept_as_text():
    e = parse_row(make_line(deadline="не дата"))
    assert e["deadline"] == "не дата"


def test_nbsp_in_number_normalized():
    # Неразрывный пробел из 1С и обычный пробел в Excel должны давать один номер.
    e = parse_row(make_line(number="123\xa0456"))
    assert e["number"] == "123 456"
    assert norm_number("123\xa0456") == norm_number("123 456") == "123 456"


def test_norm_number_collapses_whitespace():
    # В Excel номера вбиты руками: попадаются переводы строк и двойные пробелы.
    assert norm_number("111\n112") == "111 112"
    assert norm_number(" 111  112 ") == "111 112"


def test_norm_number_empty():
    assert norm_number(None) is None
    assert norm_number("") is None
    assert norm_number("  ") is None
    assert norm_number(123) == "123"


def test_numbers_in_cell_splits_by_newline():
    # В одной ячейке «Номер» бывает несколько тендеров, записанных в столбик.
    assert numbers_in_cell("111\n112 \n") == ["111", "112"]
    assert numbers_in_cell("111") == ["111"]
    assert numbers_in_cell(None) == []
    assert numbers_in_cell("") == []
