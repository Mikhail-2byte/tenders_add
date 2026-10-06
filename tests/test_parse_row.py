# -*- coding: utf-8 -*-
"""Тесты разбора выгрузки 1С. Запуск: py -3.12 -m pytest test_parse_row.py"""
import datetime

from tenders.core import parse_row, norm_number, numbers_in_cell


def make_line(number="0123456789", customer="ООО Ромашка", deadline="10.08.2026",
              status="В работе", name="наименование", etp="Сбербанк-АСТ",
              comment="комментарий"):
    """Типовая строка 1С: 13 полей через TAB, статус на позиции 9."""
    fields = ["1", "01.08.2026", deadline, "10", number, customer,
              "Тендер 000000000025397 от 14.09.2026 11:54:11", "", "Нет",
              status, name, etp, comment]
    return "\t".join(fields)


def test_typical_line():
    e = parse_row(make_line())
    assert e["number"] == "0123456789"
    assert e["customer"] == "ООО Ромашка"
    assert e["etp"] == "Сбербанк-АСТ"
    # Наименование из 1С едет в столбец B «Наименование тендера».
    assert e["name"] == "наименование"
    assert e["comment"] == "комментарий"
    assert e["deadline"] == datetime.datetime(2026, 8, 10)


def test_status_anchor_in_other_position():
    # Статус на позиции 7 — наименование, ЭТП и комментарий идут сразу после него.
    fields = ["1", "01.08.2026", "10.08.2026", "10", "111", "Заказчик",
              "x", "Подался", "срочно", "РТС-тендер", "вал 877*1шт."]
    e = parse_row("\t".join(fields))
    assert e["name"] == "срочно"
    assert e["etp"] == "РТС-тендер"
    assert e["comment"] == "вал 877*1шт."


def test_fallback_without_status():
    # Статуса нет — наименование, ЭТП и комментарий по индексам 10/11/12.
    fields = ["1", "01.08.2026", "10.08.2026", "10", "222", "Заказчик",
              "x", "x", "x", "x", "наименование-фоллбек", "ЭТП-фоллбек",
              "комментарий-фоллбек"]
    e = parse_row("\t".join(fields))
    assert e["name"] == "наименование-фоллбек"
    assert e["etp"] == "ЭТП-фоллбек"
    assert e["comment"] == "комментарий-фоллбек"


def test_comment_missing_is_empty():
    # Строка обрывается на ЭТП — комментарий пустой, разбор не падает.
    fields = ["1", "01.08.2026", "10.08.2026", "10", "333", "Заказчик",
              "x", "", "Нет", "В работе", "Валы", "b2b-center"]
    e = parse_row("\t".join(fields))
    assert e["etp"] == "b2b-center"
    assert e["comment"] == ""


def test_current_1c_b2b_row_layout():
    """Регрессия на фактический формат 1С от 06.10.2026."""
    fields = [
        "1", "30.09.2026", "07.10.2026", "6", "4621463",
        "АРТЁМОВСКИЙ РУДНИК",
        "Тендер 000000000026123 от 01.10.2026 16:54:34", "", "Да",
        "В работе", "Дробильно-размольное&#x20;", "b2b-center",
        "Запчасти к дробилке конусной КМД-1200", "",
    ]

    e = parse_row("\t".join(fields))

    assert e["number"] == "4621463"
    assert e["name"] == "Дробильно-размольное&#x20;"
    assert e["etp"] == "b2b-center"
    assert e["comment"] == "Запчасти к дробилке конусной КМД-1200"


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
