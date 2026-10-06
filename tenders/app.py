# -*- coding: utf-8 -*-
"""
«Тендеры» — локальная программа с окном (tkinter) поверх ядра tenders/core.py.

Вкладка «Добавление»: читает буфер обмена (список из 1С), предпросмотр или
боевое добавление в заметки.xlsx; ход работы виден в текстовом поле.
Вкладка «Статистика»: сводка, график и таблица «сколько тендеров добавлено
по дням» на основе history.jsonl.

Запуск: «Тендеры.bat» или py -3.12 -m tenders [--file ПУТЬ]
"""

import argparse
import datetime
import os
import queue
import threading
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from tkinter.scrolledtext import ScrolledText

from . import core as at
from . import b2b_links

CHART_DAYS = 14      # сколько последних дней показывать на графике
BAR_COLOR = "#4A90D9"
ICON_PATH = os.path.join(at.ROOT_DIR, "assets", "Тендеры.ico")


class TenderApp:
    def __init__(self, root, xlsx_path):
        self.root = root
        self.xlsx_path = xlsx_path
        self.msg_queue = queue.Queue()
        self.worker = None
        self._chart_days = {}

        root.title("Тендеры " + at.__version__)
        root.geometry("780x560")
        root.minsize(640, 460)
        if os.path.exists(ICON_PATH):
            try:
                root.iconbitmap(ICON_PATH)
            except tk.TclError:
                pass

        nb = ttk.Notebook(root)
        nb.pack(fill="both", expand=True)
        self.tab_add = ttk.Frame(nb)
        self.tab_stats = ttk.Frame(nb)
        nb.add(self.tab_add, text="  Добавление  ")
        nb.add(self.tab_stats, text="  Статистика  ")

        self._build_add_tab()
        self._build_stats_tab()
        self.refresh_stats()
        self._poll_queue()
        root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        """Во время записи окно не закрываем: обрыв посреди сохранения оставит
        недописанный файл .tmp, а работа пропадёт."""
        if self.worker is not None and self.worker.is_alive():
            messagebox.showwarning(
                "Идёт запись",
                "Файл сейчас сохраняется — подождите пару секунд.\n"
                "Окно закроется, когда работа закончится.")
            return
        self.root.destroy()

    # ---------------- вкладка «Добавление» ----------------

    def _build_add_tab(self):
        top = ttk.Frame(self.tab_add, padding=(10, 10, 10, 4))
        top.pack(fill="x")
        ttk.Label(top, text="1. Скопируйте список тендеров в 1С (Ctrl+C).\n"
                            "2. Нажмите «Предпросмотр» (ничего не меняет) "
                            "или «Добавить тендеры».").pack(anchor="w")

        btns = ttk.Frame(self.tab_add, padding=(10, 4))
        btns.pack(fill="x")
        self.btn_preview = ttk.Button(btns, text="Предпросмотр",
                                      command=lambda: self.run_sync(dry_run=True))
        self.btn_preview.pack(side="left")
        self.btn_add = ttk.Button(btns, text="Добавить тендеры",
                                  command=lambda: self.run_sync(dry_run=False))
        self.btn_add.pack(side="left", padx=8)
        # Чекбокс: автоматически проставлять ссылки B2B после добавления
        self.var_fetch_b2b = tk.BooleanVar(value=False)
        self.chk_b2b = ttk.Checkbutton(btns, text="Ссылки B2B", variable=self.var_fetch_b2b)
        self.chk_b2b.pack(side="left", padx=8)
        # Отдельный режим: проставить ссылки B2B-Center по номеру тендера.
        self.btn_b2b = ttk.Button(btns, text="Ссылки B2B (все)",
                                  command=self.run_b2b_links)
        self.btn_b2b.pack(side="left")

        file_row = ttk.Frame(self.tab_add, padding=(10, 0, 10, 4))
        file_row.pack(fill="x")
        self.file_text = tk.StringVar(value="Файл: " + self.xlsx_path)
        ttk.Button(file_row, text="Изменить…",
                   command=self.choose_file).pack(side="left")
        ttk.Label(file_row, textvariable=self.file_text,
                  foreground="#666").pack(side="left", padx=8)

        self.status = tk.StringVar(value="Готово")
        ttk.Label(self.tab_add, textvariable=self.status,
                  relief="sunken", anchor="w").pack(fill="x", side="bottom")

        self.output = ScrolledText(self.tab_add, height=18, state="disabled",
                                   font=("Consolas", 10))
        self.output.pack(fill="both", expand=True, padx=10, pady=(4, 10))

    def choose_file(self):
        """Выбор рабочего xlsx; путь запоминается в config.json до следующей смены."""
        path = filedialog.askopenfilename(
            parent=self.root, title="Выберите файл заметок",
            filetypes=[("Книга Excel", "*.xlsx"), ("Все файлы", "*.*")],
            initialfile=os.path.basename(self.xlsx_path),
            initialdir=os.path.dirname(self.xlsx_path))
        if not path:
            return
        self.xlsx_path = os.path.normpath(path)
        self.file_text.set("Файл: " + self.xlsx_path)
        at.save_xlsx_path(self.xlsx_path, say=self._append)

    def run_sync(self, dry_run):
        if self.worker is not None and self.worker.is_alive():
            return
        # Буфер обмена читаем в главном потоке (tkinter не потокобезопасен).
        try:
            text = self.root.clipboard_get()
        except tk.TclError:
            text = ""
        self._clear_output()
        self._set_busy(True, "Предпросмотр..." if dry_run else "Добавление...")
        # Поток НЕ daemon: если окно всё же закроют, запись успеет завершиться.
        # daemon-поток Python убивает мгновенно — прямо посреди сохранения файла.
        self.worker = threading.Thread(target=self._worker_run,
                                       args=(text, dry_run, self.var_fetch_b2b.get()))
        self.worker.start()

    def _worker_run(self, text, dry_run, fetch_b2b):
        def say(*parts):
            self.msg_queue.put(("line", " ".join(str(p) for p in parts)))
        try:
            result = at.sync(self.xlsx_path, text, dry_run=dry_run, fetch_b2b_links=fetch_b2b, say=say)
            self.msg_queue.put(("done", result))
        except at.SyncError as ex:
            self.msg_queue.put(("warn", str(ex)))
        except Exception:
            import traceback
            self.msg_queue.put(("error", traceback.format_exc()))

    def run_b2b_links(self):
        """Проставить ссылки B2B-Center в пустые «Ссылка» у строк с ЭТП = B2B.

        Идёт в интернет с паузой ~1.5 с на строку, поэтому по многим строкам
        может занять минуты — ход виден в окне вывода.
        """
        if self.worker is not None and self.worker.is_alive():
            return
        self._clear_output()
        self._set_busy(True, "Поиск ссылок B2B...")
        self.worker = threading.Thread(target=self._worker_b2b)
        self.worker.start()

    def _worker_b2b(self):
        def say(*parts):
            self.msg_queue.put(("line", " ".join(str(p) for p in parts)))
        try:
            result = b2b_links.fill_b2b_links(self.xlsx_path, dry_run=False, say=say)
            self.msg_queue.put(("done", result))
        except at.SyncError as ex:
            self.msg_queue.put(("warn", str(ex)))
        except Exception:
            import traceback
            self.msg_queue.put(("error", traceback.format_exc()))

    def _poll_queue(self):
        try:
            while True:
                kind, payload = self.msg_queue.get_nowait()
                if kind == "line":
                    self._append(payload)
                elif kind == "done":
                    if "filled" in payload:          # итог режима «Ссылки B2B»
                        self._set_busy(
                            False, "Готово: ссылок %d (не найдено %d, неоднозначно %d)"
                            % (payload.get("filled", 0), len(payload.get("not_found", [])),
                               len(payload.get("ambiguous", []))))
                    elif payload.get("dry_run"):
                        self._set_busy(False, "Предпросмотр готов")
                    else:
                        self._set_busy(False, "Готово: добавлено %d"
                                       % len(payload.get("new", [])))
                        self.refresh_stats()
                elif kind == "warn":
                    self._append(payload)
                    self._set_busy(False, "Не выполнено")
                    messagebox.showwarning("Тендеры", payload, parent=self.root)
                elif kind == "error":
                    self._append(payload)
                    self._set_busy(False, "Ошибка")
                    messagebox.showerror(
                        "Тендеры",
                        "Непредвиденная ошибка (подробности в окне вывода).",
                        parent=self.root)
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _set_busy(self, busy, status_text):
        state = "disabled" if busy else "normal"
        self.btn_preview.configure(state=state)
        self.btn_add.configure(state=state)
        self.btn_b2b.configure(state=state)
        # Чекбокс B2B тоже отключаем во время работы
        try:
            self.chk_b2b.configure(state=state)
        except AttributeError:
            pass
        self.status.set(status_text)

    def _clear_output(self):
        self.output.configure(state="normal")
        self.output.delete("1.0", "end")
        self.output.configure(state="disabled")

    def _append(self, line):
        self.output.configure(state="normal")
        self.output.insert("end", line + "\n")
        self.output.see("end")
        self.output.configure(state="disabled")

    # ---------------- вкладка «Статистика» ----------------

    def _build_stats_tab(self):
        f = ttk.Frame(self.tab_stats, padding=10)
        f.pack(fill="both", expand=True)

        row = ttk.Frame(f)
        row.pack(fill="x")
        self.summary = tk.StringVar(value="")
        ttk.Label(row, textvariable=self.summary,
                  font=("Segoe UI", 10, "bold")).pack(side="left")
        ttk.Button(row, text="Обновить", command=self.refresh_stats).pack(side="right")

        ttk.Label(f, text="Добавлено тендеров за последние %d дней:" % CHART_DAYS,
                  foreground="#666").pack(anchor="w", pady=(8, 2))
        self.canvas = tk.Canvas(f, height=170, background="white",
                                highlightthickness=1, highlightbackground="#cccccc")
        self.canvas.pack(fill="x")
        self.canvas.bind("<Configure>", lambda e: self._draw_chart())

        table = ttk.Frame(f)
        table.pack(fill="both", expand=True, pady=(10, 0))
        cols = ("date", "count", "numbers")
        self.tree = ttk.Treeview(table, columns=cols, show="headings")
        self.tree.heading("date", text="Дата")
        self.tree.heading("count", text="Добавлено")
        self.tree.heading("numbers", text="Номера")
        self.tree.column("date", width=110, anchor="center", stretch=False)
        self.tree.column("count", width=90, anchor="center", stretch=False)
        self.tree.column("numbers", width=380)
        sb = ttk.Scrollbar(table, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.tree.pack(side="left", fill="both", expand=True)

    def refresh_stats(self):
        days = at.counts_by_day(at.load_history())
        today = datetime.date.today()

        def since(d0):
            return sum(v["count"] for d, v in days.items() if d >= d0)

        self.summary.set(
            "Сегодня: %d      За 7 дней: %d      За 30 дней: %d      Всего: %d"
            % (days.get(today, {}).get("count", 0),
               since(today - datetime.timedelta(days=6)),
               since(today - datetime.timedelta(days=29)),
               sum(v["count"] for v in days.values())))

        self.tree.delete(*self.tree.get_children())
        for d in sorted(days, reverse=True):
            v = days[d]
            self.tree.insert("", "end", values=(
                d.strftime("%d.%m.%Y"), v["count"], ", ".join(v["numbers"])))

        self._chart_days = days
        self._draw_chart()

    def _draw_chart(self):
        c = self.canvas
        c.delete("all")
        w = c.winfo_width()
        h = c.winfo_height()
        if w < 60 or h < 60:            # окно ещё не разложено
            w, h = 740, 170
        today = datetime.date.today()
        dates = [today - datetime.timedelta(days=i)
                 for i in range(CHART_DAYS - 1, -1, -1)]
        counts = [self._chart_days.get(d, {}).get("count", 0) for d in dates]
        mx = max(counts + [1])

        pad = 16
        base_y = h - 24
        bar_w = (w - 2 * pad) / CHART_DAYS
        c.create_line(pad, base_y, w - pad, base_y, fill="#999999")
        for i, (d, cnt) in enumerate(zip(dates, counts)):
            x0 = pad + i * bar_w + 4
            x1 = pad + (i + 1) * bar_w - 4
            xc = (x0 + x1) / 2
            c.create_text(xc, h - 12, text=d.strftime("%d.%m"),
                          font=("Segoe UI", 7), fill="#555555")
            if cnt:
                bar_h = (base_y - 30) * cnt / mx
                c.create_rectangle(x0, base_y - bar_h, x1, base_y,
                                   fill=BAR_COLOR, outline="")
                c.create_text(xc, base_y - bar_h - 9, text=str(cnt),
                              font=("Segoe UI", 8), fill="#333333")


def main():
    ap = argparse.ArgumentParser(description="Тендеры — окно программы")
    ap.add_argument("--file", metavar="ПУТЬ",
                    help="путь к xlsx (по умолчанию — выбранный в окне программы)")
    args = ap.parse_args()

    root = tk.Tk()
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    TenderApp(root, args.file or at.load_xlsx_path())
    root.mainloop()


if __name__ == "__main__":
    main()
