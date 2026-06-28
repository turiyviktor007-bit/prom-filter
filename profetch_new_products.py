#!/usr/bin/env python3
"""
Profetch — Генератор нових товарів для Пром.юа
================================================
Порівнює ВАШ ФІД (Lexx/MrSeller — файл або URL) з товарами на Пром через API.
Генерує XLS тільки з НОВИМИ товарами (яких ще немає на Пром),
включно з усіма характеристиками з фіду.

pip install requests openpyxl
python profetch_new_products.py
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import sys
import time
import xml.etree.ElementTree as ET
import requests
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment


# ── КОЛЬОРИ ──────────────────────────────────
BG, BG2, BG3 = "#0f1117", "#1a1d27", "#22263a"
ACCENT, SUCCESS, WARNING, DANGER = "#4f8ef7", "#3ecf8e", "#f7b731", "#fc5c65"
TEXT, TEXT_DIM, BORDER = "#e8eaf6", "#6b7280", "#2d3250"

FONT_TITLE = ("Segoe UI", 13, "bold")
FONT_LABEL = ("Segoe UI", 9)
FONT_SMALL = ("Segoe UI", 8)
FONT_MONO  = ("Consolas", 9)
FONT_BIG   = ("Segoe UI", 22, "bold")
FONT_BTN   = ("Segoe UI", 10, "bold")

# Базові колонки формату Пром (1-51)
PROM_BASE_HEADERS = [
    "Код_товара", "Название_позиции", "Название_позиции_укр",
    "Поисковые_запросы", "Поисковые_запросы_укр", "Описание", "Описание_укр",
    "Тип_товара", "Цена", "Валюта", "Единица_измерения",
    "Минимальный_объем_заказа", "Оптовая_цена", "Минимальный_заказ_опт",
    "Ссылка_изображения", "Наличие", "Количество", "Номер_группы",
    "Название_группы", "Адрес_подраздела", "Возможность_поставки",
    "Срок_поставки", "Способ_упаковки", "Способ_упаковки_укр",
    "Уникальный_идентификатор", "Идентификатор_товара",
    "Идентификатор_подраздела", "Идентификатор_группы", "Производитель",
    "Страна_производитель", "Скидка", "ID_группы_разновидностей",
    "Личные_заметки", "Продукт_на_сайте", "Срок_действия_скидки_от",
    "Срок_действия_скидки_до", "Цена_от", "Ярлык", "HTML_заголовок",
    "HTML_заголовок_укр", "HTML_описание", "HTML_описание_укр",
    "Код_маркировки_(GTIN)", "Номер_устройства_(MPN)", "Вес,кг",
    "Ширина,см", "Высота,см", "Длина,см", "Где_находится_товар",
    "Товар_в_ProSale", "Почему_товар_не_в_ProSale",
]


# ── ЛОГІКА ───────────────────────────────────

def get_prom_skus(token, on_progress):
    """Завантажує всі артикули (sku) та зовнішні ID з Пром через API."""
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    skus, ext_ids = set(), set()
    limit, last_id, page = 100, None, 1

    while True:
        params = {"limit": limit, "status": "on_display,draft,deleted"}
        if last_id:
            params["last_id"] = last_id
        resp = requests.get("https://my.prom.ua/api/v1/products/list",
                            headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        products = resp.json().get("products", [])
        if not products:
            break
        for p in products:
            sku = str(p.get("sku", "")).strip()
            ext = str(p.get("external_id", "")).strip()
            if sku:
                skus.add(sku)
            if ext and ext not in ("None", ""):
                ext_ids.add(ext)
        on_progress(f"   Сторінка {page}: {len(skus)} артикулів")
        if len(products) < limit:
            break
        last_id = products[-1]["id"]
        page += 1
        time.sleep(0.4)
    return skus, ext_ids


def parse_feed(source, on_progress):
    """Парсить YML фід (файл або URL). Повертає список товарів."""
    on_progress("   Читання фіду...")
    if source.startswith("http"):
        resp = requests.get(source, timeout=60, headers={
            "User-Agent": "Mozilla/5.0 (compatible; ProfetchBot/1.0)"})
        resp.raise_for_status()
        content = resp.content
    else:
        with open(source, "rb") as f:
            content = f.read()

    root = ET.fromstring(content)
    offers = root.findall(".//offer")
    on_progress(f"   Знайдено {len(offers)} товарів у фіді")

    products = []
    for o in offers:
        offer_id = o.get("id", "").strip()
        available = o.get("available", "true")
        vendor_code = (o.findtext("vendorCode", "") or offer_id).strip()
        name = o.findtext("name", "").strip()
        price = o.findtext("price", "").strip()
        currency = o.findtext("currencyId", "UAH").strip()
        category = o.findtext("categoryId", "").strip()
        description = o.findtext("description", "").strip()
        vendor = o.findtext("vendor", "").strip()
        country = o.findtext("country_of_origin", "").strip()
        url = o.findtext("url", "").strip()

        pictures = [p.text.strip() for p in o.findall("picture") if p.text]

        params = {}
        gtin = ""
        for p in o.findall("param"):
            pname = p.get("name", "").strip()
            pval  = (p.text or "").strip()
            if pname and pval:
                params[pname] = pval
                if pname == "Штрих код" and pval.isdigit():
                    gtin = pval

        products.append({
            "id": offer_id, "available": available, "vendor_code": vendor_code,
            "name": name, "price": price, "currency": currency,
            "category": category, "description": description,
            "vendor": vendor, "country": country, "url": url,
            "pictures": pictures, "params": params, "gtin": gtin,
        })
    return products


def find_new(products, prom_skus, prom_ext_ids, only_available, on_progress):
    """Повертає товари яких немає на Пром."""
    on_progress("   Порівняння з Пром...")
    new_items, already, skipped_unavail = [], 0, 0
    for p in products:
        if only_available and p["available"] != "true":
            skipped_unavail += 1
            continue
        vc = p["vendor_code"]
        pid = p["id"]
        on_prom = (vc in prom_skus) or (pid in prom_ext_ids) or (vc in prom_ext_ids)
        if on_prom:
            already += 1
        else:
            new_items.append(p)
    return new_items, already, skipped_unavail


def generate_xls(new_items, output_path, on_progress):
    """Генерує XLS у форматі Пром з новими товарами + характеристики."""
    on_progress("   Збір характеристик...")

    # Знаходимо макс кількість характеристик серед нових товарів
    max_params = max((len(p["params"]) for p in new_items), default=0)
    max_params = max(max_params, 1)

    # Заголовки: базові + трійки характеристик
    headers = list(PROM_BASE_HEADERS)
    for _ in range(max_params):
        headers += ["Название_Характеристики", "Измерение_Характеристики",
                    "Значение_Характеристики"]

    on_progress("   Генерація XLS...")
    wb = Workbook()
    ws = wb.active
    ws.title = "Export Products Sheet"

    # Заголовок
    hfont = Font(name="Arial", bold=True, size=10, color="FFFFFF")
    hfill = PatternFill("solid", start_color="1F4E79")
    halign = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.append(headers)
    for cell in ws[1]:
        cell.font, cell.fill, cell.alignment = hfont, hfill, halign
    ws.row_dimensions[1].height = 32

    dfont = Font(name="Arial", size=9)

    for idx, p in enumerate(new_items):
        avail = "!" if p["available"] == "true" else "-"
        row = [
            p["vendor_code"],          # Код_товара
            p["name"],                 # Название_позиции
            "",                        # Название_позиции_укр
            "", "",                    # Поисковые запросы
            p["description"],          # Описание
            "",                        # Описание_укр
            "r",                       # Тип_товара
            p["price"],                # Цена
            p["currency"] or "UAH",    # Валюта
            "шт.",                     # Единица_измерения
            "", "", "",                # мин объём, опт цена, мин опт
            ", ".join(p["pictures"]),  # Ссылка_изображения
            avail,                     # Наличие
            "",                        # Количество
            p["category"],             # Номер_группы (categoryId з фіду!)
            "",                        # Название_группы
            "",                        # Адрес_подраздела
            "", "", "", "",            # поставка, упаковка
            "",                        # Уникальный_идентификатор
            p["id"],                   # Идентификатор_товара (внешний ID)
            "",                        # Идентификатор_подраздела
            "",                        # Идентификатор_группы (пусто!)
            p["vendor"],               # Производитель
            p["country"],              # Страна_производитель
            "", "", "", "",            # скидка, группа разнов, заметки, на сайте
            "", "", "",                # скидки от/до, цена от
            "",                        # Ярлык (пусто!)
            "", "", "", "",            # HTML заголовок/описание
            p["gtin"],                 # GTIN
            "",                        # MPN
            "", "", "", "",            # вес, габариты
            "Киев",                    # Где_находится
            "Нет", "",                 # ProSale
        ]
        # Характеристики
        for pname, pval in p["params"].items():
            row += [pname, "", pval]
        # Доповнюємо до довжини headers
        while len(row) < len(headers):
            row.append("")

        ws.append(row)
        excel_row = idx + 2
        for cell in ws[excel_row]:
            cell.font = dfont

    # Ширини колонок
    widths = {1: 12, 2: 50, 6: 60, 9: 10, 15: 80, 16: 8, 26: 15, 29: 18, 30: 14, 43: 15}
    for col, w in widths.items():
        ws.column_dimensions[ws.cell(row=1, column=col).column_letter].width = w

    ws.freeze_panes = "A2"
    wb.save(output_path)

    size_kb = os.path.getsize(output_path) / 1024
    on_progress(f"   Збережено: {os.path.basename(output_path)} ({size_kb:.0f} КБ)")
    return max_params


# ── GUI ──────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Profetch — Генератор нових товарів")
        self.geometry("700x760")
        self.minsize(640, 680)
        self.configure(bg=BG)

        self.feed_src   = tk.StringVar(value="")
        self.prom_token = tk.StringVar(value="")
        self.out_path   = tk.StringVar(value="")
        self.only_avail = tk.BooleanVar(value=True)
        self._running   = False

        self._build_ui()
        self._load_settings()

    def _build_ui(self):
        hdr = tk.Frame(self, bg=BG, pady=18)
        hdr.pack(fill="x", padx=24)
        tk.Label(hdr, text="⬡", font=("Segoe UI", 20), fg=ACCENT, bg=BG).pack(side="left")
        tk.Label(hdr, text="  PROFETCH", font=FONT_TITLE, fg=TEXT, bg=BG).pack(side="left")
        tk.Label(hdr, text="  Нові товари → XLS", font=("Segoe UI", 10),
                 fg=TEXT_DIM, bg=BG).pack(side="left", pady=3)

        self._sep()

        # 1. Фід
        self._section("1  Ваш фід (Lexx / MrSeller)")
        tk.Label(self, text="Виберіть XML-файл або вставте URL-посилання на фід",
                 font=FONT_SMALL, fg=TEXT_DIM, bg=BG).pack(anchor="w", padx=28, pady=(0, 4))
        r1 = tk.Frame(self, bg=BG)
        r1.pack(fill="x", padx=24, pady=(0, 8))
        e1 = self._entry(r1, self.feed_src, 52)
        e1.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self._btn_small(r1, "Огляд…", self._pick_feed).pack(side="left")

        # 2. Токен
        self._section("2  API-токен Пром.юа")
        tk.Label(self, text="Пром → Налаштування → Управління API-токенами",
                 font=FONT_SMALL, fg=TEXT_DIM, bg=BG).pack(anchor="w", padx=28, pady=(0, 4))
        r2 = tk.Frame(self, bg=BG)
        r2.pack(fill="x", padx=24, pady=(0, 8))
        self.tok_entry = self._entry(r2, self.prom_token, 52, show="•")
        self.tok_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self._btn_small(r2, "👁", self._toggle_token).pack(side="left")

        # 3. Результат
        self._section("3  Зберегти результат як")
        r3 = tk.Frame(self, bg=BG)
        r3.pack(fill="x", padx=24, pady=(0, 8))
        e3 = self._entry(r3, self.out_path, 52)
        e3.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self._btn_small(r3, "Огляд…", self._pick_out).pack(side="left")

        # Чекбокс — тільки в наявності
        chk_frame = tk.Frame(self, bg=BG)
        chk_frame.pack(fill="x", padx=24, pady=(4, 0))
        chk = tk.Checkbutton(
            chk_frame, text="Тільки товари в наявності (пропустити відсутні)",
            variable=self.only_avail, font=FONT_SMALL,
            bg=BG, fg=TEXT_DIM, selectcolor=BG3,
            activebackground=BG, activeforeground=TEXT,
            bd=0, highlightthickness=0
        )
        chk.pack(anchor="w")

        self._sep()

        self.run_btn = tk.Button(
            self, text="▶   ЗНАЙТИ НОВІ ТОВАРИ → XLS",
            font=FONT_BTN, bg=ACCENT, fg="#ffffff",
            activebackground="#7c5cfc", activeforeground="#ffffff",
            bd=0, cursor="hand2", padx=28, pady=12, command=self._start
        )
        self.run_btn.pack(fill="x", padx=24, pady=(12, 4))

        pb = tk.Frame(self, bg=BG, height=6)
        pb.pack(fill="x", padx=24, pady=(0, 8))
        pb.pack_propagate(False)
        self.pb_bg = tk.Frame(pb, bg=BG3, height=6)
        self.pb_bg.pack(fill="both", expand=True)
        self.pb_fill = tk.Frame(self.pb_bg, bg=ACCENT, height=6, width=0)
        self.pb_fill.place(x=0, y=0, height=6)
        self._pb_anim, self._pb_dir = 0, 1

        tk.Label(self, text="Журнал", font=FONT_SMALL, fg=TEXT_DIM, bg=BG).pack(anchor="w", padx=24)
        self.log_box = tk.Text(self, font=FONT_MONO, bg=BG2, fg=TEXT,
                               insertbackground=TEXT, bd=0, pady=10, padx=12,
                               relief="flat", state="disabled", height=8, selectbackground=BG3)
        self.log_box.pack(fill="both", expand=True, padx=24, pady=(4, 0))

        sf = tk.Frame(self, bg=BG3, pady=14)
        sf.pack(fill="x", padx=24, pady=8)
        self.stat_feed = self._stat_cell(sf, "У фіді", "—")
        self.stat_prom = self._stat_cell(sf, "Вже на Пром", "—")
        self.stat_new  = self._stat_cell(sf, "НОВИХ → XLS", "—", color=SUCCESS)
        for c in range(3):
            sf.columnconfigure(c, weight=1)

        self.open_btn = tk.Button(self, text="📂  Відкрити папку з результатом",
                                  font=FONT_SMALL, bg=BG2, fg=TEXT_DIM,
                                  activebackground=BG3, activeforeground=TEXT,
                                  bd=0, cursor="hand2", padx=16, pady=8,
                                  command=self._open_folder, state="disabled")
        self.open_btn.pack(pady=(0, 16))

    # ── хелпери ──
    def _sep(self):
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=24, pady=8)

    def _section(self, text):
        tk.Label(self, text=text, font=("Segoe UI", 9, "bold"),
                 fg=ACCENT, bg=BG).pack(anchor="w", padx=24, pady=(12, 4))

    def _entry(self, parent, var, width=40, show=None):
        kw = dict(textvariable=var, font=FONT_LABEL, bg=BG3, fg=TEXT,
                  insertbackground=TEXT, bd=0, relief="flat",
                  highlightthickness=1, highlightbackground=BORDER,
                  highlightcolor=ACCENT, width=width)
        if show:
            kw["show"] = show
        e = tk.Entry(parent, **kw)
        self._enable_paste(e)
        return e

    def _enable_paste(self, entry):
        def do_paste(event=None):
            try:
                text = self.clipboard_get()
                try:
                    entry.delete("sel.first", "sel.last")
                except tk.TclError:
                    pass
                entry.insert("insert", text)
            except tk.TclError:
                pass
            return "break"

        def do_copy(event=None):
            try:
                self.clipboard_clear()
                self.clipboard_append(entry.selection_get())
            except tk.TclError:
                pass
            return "break"

        def do_cut(event=None):
            do_copy()
            try:
                entry.delete("sel.first", "sel.last")
            except tk.TclError:
                pass
            return "break"

        def select_all(event=None):
            entry.select_range(0, "end")
            entry.icursor("end")
            return "break"

        menu = tk.Menu(entry, tearoff=0, bg=BG3, fg=TEXT,
                       activebackground=ACCENT, activeforeground="#ffffff", bd=0)
        menu.add_command(label="Вставити", command=do_paste)
        menu.add_command(label="Копіювати", command=do_copy)
        menu.add_command(label="Вирізати", command=do_cut)
        menu.add_separator()
        menu.add_command(label="Виділити все", command=select_all)

        def show_menu(event):
            try:
                menu.tk_popup(event.x_root, event.y_root)
            finally:
                menu.grab_release()

        entry.bind("<Button-3>", show_menu)
        for seq in ("<Control-v>", "<Control-V>"):
            entry.bind(seq, do_paste)
        for seq in ("<Control-c>", "<Control-C>"):
            entry.bind(seq, do_copy)
        for seq in ("<Control-x>", "<Control-X>"):
            entry.bind(seq, do_cut)
        for seq in ("<Control-a>", "<Control-A>"):
            entry.bind(seq, select_all)

        def keycode_handler(event):
            if event.state & 0x4:
                if event.keycode == 86:
                    return do_paste()
                if event.keycode == 67:
                    return do_copy()
                if event.keycode == 88:
                    return do_cut()
                if event.keycode == 65:
                    return select_all()
        entry.bind("<Key>", keycode_handler)

    def _btn_small(self, parent, text, cmd):
        return tk.Button(parent, text=text, font=FONT_SMALL, bg=BG3, fg=TEXT_DIM,
                         activebackground=BORDER, activeforeground=TEXT,
                         bd=0, cursor="hand2", padx=10, pady=6, relief="flat",
                         command=cmd, highlightthickness=1, highlightbackground=BORDER)

    def _stat_cell(self, parent, label, value, color=TEXT):
        frame = tk.Frame(parent, bg=BG3)
        frame.grid(row=0, column=len(parent.winfo_children()) - 1, sticky="nsew", padx=6)
        tk.Label(frame, text=label, font=FONT_SMALL, fg=TEXT_DIM, bg=BG3).pack(pady=(8, 0))
        val = tk.Label(frame, text=value, font=FONT_BIG, fg=color, bg=BG3)
        val.pack(pady=(0, 8))
        return val

    def _toggle_token(self):
        self.tok_entry.configure(show="" if self.tok_entry.cget("show") else "•")

    def _pick_feed(self):
        path = filedialog.askopenfilename(title="Виберіть XML-фід",
                                          filetypes=[("XML/YML", "*.xml *.yml"), ("Всі файли", "*.*")])
        if path:
            self.feed_src.set(path)
            if not self.out_path.get():
                self.out_path.set(os.path.splitext(path)[0] + "-new.xlsx")

    def _pick_out(self):
        path = filedialog.asksaveasfilename(title="Зберегти як", defaultextension=".xlsx",
                                            filetypes=[("Excel", "*.xlsx"), ("Всі файли", "*.*")])
        if path:
            self.out_path.set(path)

    def _open_folder(self):
        path = self.out_path.get()
        if path and os.path.exists(path) and sys.platform == "win32":
            os.startfile(os.path.dirname(os.path.abspath(path)))

    def _log(self, text, color=None):
        self.log_box.configure(state="normal")
        tag = None
        if color:
            tag = f"c_{color.replace('#','')}"
            self.log_box.tag_configure(tag, foreground=color)
        self.log_box.insert("end", text + "\n", tag or "")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _log_clear(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _pb_start(self):
        self._pb_anim, self._pb_dir = 0, 1
        self._pb_tick()

    def _pb_tick(self):
        if not self._running:
            self.pb_fill.place(x=0, y=0, height=6, width=0)
            return
        w = self.pb_bg.winfo_width()
        bar_w = max(w // 4, 60)
        self._pb_anim += self._pb_dir * 6
        if self._pb_anim >= w - bar_w:
            self._pb_dir = -1
        if self._pb_anim <= 0:
            self._pb_dir, self._pb_anim = 1, 0
        self.pb_fill.place(x=self._pb_anim, y=0, height=6, width=bar_w)
        self.after(20, self._pb_tick)

    def _settings_path(self):
        return os.path.join(os.path.expanduser("~"), ".profetch_new.ini")

    def _save_settings(self):
        try:
            with open(self._settings_path(), "w", encoding="utf-8") as f:
                f.write(f"token={self.prom_token.get()}\n")
                f.write(f"feed={self.feed_src.get()}\n")
                f.write(f"out={self.out_path.get()}\n")
        except Exception:
            pass

    def _load_settings(self):
        try:
            p = self._settings_path()
            if not os.path.exists(p):
                return
            with open(p, encoding="utf-8") as f:
                for line in f:
                    k, _, v = line.strip().partition("=")
                    if k == "token":
                        self.prom_token.set(v)
                    elif k == "feed":
                        self.feed_src.set(v)
                    elif k == "out":
                        self.out_path.set(v)
        except Exception:
            pass

    def _start(self):
        feed = self.feed_src.get().strip()
        token = self.prom_token.get().strip()
        out = self.out_path.get().strip()

        if not feed:
            messagebox.showwarning("Увага", "Виберіть фід або вставте URL")
            return
        if not feed.startswith("http") and not os.path.exists(feed):
            messagebox.showwarning("Увага", "Файл фіду не знайдено")
            return
        if not token:
            messagebox.showwarning("Увага", "Введіть API-токен Пром.юа")
            return
        if not out:
            messagebox.showwarning("Увага", "Вкажіть куди зберегти результат")
            return

        self._save_settings()
        self._log_clear()
        self._running = True
        self.run_btn.configure(state="disabled", text="⏳  Обробка...")
        self.open_btn.configure(state="disabled")
        for s in (self.stat_feed, self.stat_prom, self.stat_new):
            s.configure(text="…", fg=TEXT_DIM)
        self._pb_start()

        threading.Thread(target=self._worker, args=(feed, token, out), daemon=True).start()

    def _worker(self, feed, token, out):
        def log(msg, color=None):
            self.after(0, lambda: self._log(msg, color))

        try:
            log("⟳  Підключення до Пром.юа API...")
            prom_skus, prom_ext = get_prom_skus(token, lambda m: log(m))
            log(f"✓  Пром: {len(prom_skus)} артикулів", SUCCESS)

            log("")
            log("⟳  Читання фіду...")
            products = parse_feed(feed, lambda m: log(m))
            log(f"✓  Фід: {len(products)} товарів", SUCCESS)

            log("")
            log("⟳  Порівняння...")
            new_items, already, skipped = find_new(
                products, prom_skus, prom_ext,
                self.only_avail.get(), lambda m: log(m))
            log(f"✓  Результат:", SUCCESS)
            log(f"   У фіді:        {len(products)}")
            if skipped:
                log(f"   Пропущено (немає в наявності): {skipped}", WARNING)
            log(f"   Вже на Пром:   {already}")
            log(f"   НОВИХ:         {len(new_items)}",
                SUCCESS if new_items else TEXT_DIM)

            self.after(0, lambda: [
                self.stat_feed.configure(text=str(len(products)), fg=TEXT),
                self.stat_prom.configure(text=str(already), fg=TEXT),
                self.stat_new.configure(text=str(len(new_items)),
                                        fg=SUCCESS if new_items else TEXT_DIM)
            ])

            if new_items:
                log("")
                log("⟳  Генерація XLS...")
                os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
                maxp = generate_xls(new_items, out, lambda m: log(m))
                log(f"✓  Характеристик на товар: до {maxp}", SUCCESS)
                log("")
                log("─" * 44, TEXT_DIM)
                log("ЩО РОБИТИ ДАЛІ:", ACCENT)
                log("  1. Відкрийте файл в Excel — перевірте", TEXT)
                log("  2. Пром → Товари → Імпорт позицій", TEXT)
                log(f"  3. Завантажте: {os.path.basename(out)}", TEXT)
                log("  4. БЕЗ галочки «Тільки оновлення»", WARNING)
                log("  5. «Почати імпорт»", TEXT)
                log(f"  → {len(new_items)} нових товарів з характеристиками", SUCCESS)
                log("─" * 44, TEXT_DIM)
                self.after(0, lambda: self.open_btn.configure(state="normal"))
            else:
                log("")
                log("✅  Нових товарів немає — всі вже на Пром!", SUCCESS)

        except requests.exceptions.HTTPError as e:
            log(f"✗  HTTP помилка: {e}", DANGER)
            if "401" in str(e):
                log("   Перевірте API-токен Пром", DANGER)
        except requests.exceptions.ConnectionError:
            log("✗  Немає з'єднання з інтернетом", DANGER)
        except ET.ParseError as e:
            log(f"✗  Помилка XML: {e}", DANGER)
        except Exception as e:
            log(f"✗  Помилка: {e}", DANGER)
            import traceback
            log(traceback.format_exc(), DANGER)
        finally:
            self._running = False
            self.after(0, lambda: self.run_btn.configure(
                state="normal", text="▶   ЗНАЙТИ НОВІ ТОВАРИ → XLS"))


if __name__ == "__main__":
    App().mainloop()
