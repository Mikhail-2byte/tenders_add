# -*- coding: utf-8 -*-
"""Тесты разбора выгрузки 1С. Запуск: py -3.12 -m pytest test_parse_row.py"""
import datetime

from add_tenders import parse_row, norm_number


def make_line(number="0123456789", customer="ООО Ромашка", deadline="10.08.2026",
              status="В работе", etp="Сбербанк-АСТ", comment="комментарий"):
    """Типовая строка 1С: 12 полей через TAB, статус на позиции 9."""
    fields = ["x", "Да", deadline, "x", number, customer,
              "x", "x", "x", status, etp, comment]
    return "\t".join(fields)


def test_typical_line():
    e = parse_row(make_line())
    assert e["number"] == "0123456789"
    assert e["customer"] == "ООО Ромашка"
    assert e["etp"] == "Сбербанк-АСТ"
    assert e["comment"] == "комментарий"
    assert e["deadline"] == datetime.datetime(2026, 8, 10)
    assert e["deadline_has_time"] is False


def test_status_anchor_in_other_position():
    # Статус на позиции 7 — ЭТП и комментарий берутся сразу после него.
    fields = ["x", "Да", "10.08.2026", "x", "111", "Заказчик",
              "x", "Подался", "РТС-тендер", "срочно"]
    e = parse_row("\t".join(fields))
    assert e["etp"] == "РТС-тендер"
    assert e["comment"] == "срочно"


def test_fallback_without_status():
    # Статуса нет — ЭТП и комментарий по фиксированным индексам 10/11.
    fields = ["x", "Да", "10.08.2026", "x", "222", "Заказчик",
              "x", "x", "x", "x", "ЭТП-фоллбек", "коммент-фоллбек"]
    e = parse_row("\t".join(fields))
    assert e["etp"] == "ЭТП-фоллбек"
    assert e["comment"] == "коммент-фоллбек"


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
    assert e["deadline_has_time"] is True
    e = parse_row(make_line(deadline="10.08.2026 09:00"))
    assert e["deadline"] == datetime.datetime(2026, 8, 10, 9, 0)
    assert e["deadline_has_time"] is True


def test_deadline_unparsed_kept_as_text():
    e = parse_row(make_line(deadline="не дата"))
    assert e["deadline"] == "не дата"
    assert e["deadline_has_time"] is False


def test_nbsp_in_number_normalized():
    # Неразрывный пробел из 1С и обычный пробел в Excel должны давать один номер.
    e = parse_row(make_line(number="123\xa0456"))
    assert e["number"] == "123 456"
    assert norm_number("123\xa0456") == norm_number("123 456") == "123 456"


def test_norm_number_empty():
    assert norm_number(None) is None
    assert norm_number("") is None
    assert norm_number("  ") is None
    assert norm_number(123) == "123"
