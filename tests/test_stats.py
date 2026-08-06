# -*- coding: utf-8 -*-
"""Тесты истории/статистики. Запуск: py -3.12 -m pytest test_stats.py"""
import datetime

from tenders import core as at


def test_counts_by_day_groups_and_sums():
    recs = [
        {"ts": "2026-07-06T10:00:00", "added": 2, "numbers": ["1", "2"]},
        {"ts": "2026-07-06T15:30:00", "added": 1, "numbers": ["3"]},
        {"ts": "2026-07-07T09:00:00", "added": 0, "numbers": []},
    ]
    days = at.counts_by_day(recs)
    d6 = days[datetime.date(2026, 7, 6)]
    assert d6["count"] == 3
    assert d6["numbers"] == ["1", "2", "3"]
    assert days[datetime.date(2026, 7, 7)]["count"] == 0


def test_counts_by_day_skips_broken_records():
    recs = [
        {"ts": "мусор"},
        {"без_ts": 1},
        {"ts": None},
        {"ts": "2026-07-07T09:00:00", "added": 1, "numbers": ["5"]},
    ]
    days = at.counts_by_day(recs)
    assert list(days) == [datetime.date(2026, 7, 7)]
    assert days[datetime.date(2026, 7, 7)]["count"] == 1


def test_load_history_missing_file(tmp_path, monkeypatch):
    monkeypatch.setattr(at, "HISTORY_PATH", str(tmp_path / "нет_такого.jsonl"))
    assert at.load_history() == []


def test_record_then_load_skipping_broken_lines(tmp_path, monkeypatch):
    hist = tmp_path / "history.jsonl"
    monkeypatch.setattr(at, "HISTORY_PATH", str(hist))

    at.record_history(["100", "200"], r"C:\файл.xlsx")
    with open(hist, "a", encoding="utf-8") as f:
        f.write("это не json\n\n")
    at.record_history([], r"C:\файл.xlsx")

    recs = at.load_history()
    assert len(recs) == 2
    assert recs[0]["added"] == 2
    assert recs[0]["numbers"] == ["100", "200"]
    assert recs[1]["added"] == 0
    # ts парсится и попадает в свод по дням
    days = at.counts_by_day(recs)
    assert days[datetime.date.today()]["count"] == 2
