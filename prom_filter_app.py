#!/usr/bin/env python3
"""
Profetch — Prom New Products Filter
Графічна утиліта для порівняння XML фідів з товарами на Пром.юа
"""

import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import xml.etree.ElementTree as ET
from xml.dom import minidom
import requests
import time
import os
import sys

# ──────────────────────────────────────────────
#  КОЛЬОРИ / СТИЛЬ
# ──────────────────────────────────────────────
BG        = "#0f1117"
BG2       = "#1a1d27"
BG3       = "#22263a"
ACCENT    = "#4f8ef7"
ACCENT2   = "#7c5cfc"
SUCCESS   = "#3ecf8e"
WARNING   = "#f7b731"
DANGER    = "#fc5c65"
TEXT      = "#e8eaf6"
TEXT_DIM  = "#6b7280"
BORDER    = "#2d3250"

FONT_TITLE  = ("Segoe UI", 13, "bold")
FONT_LABEL  = ("Segoe UI", 9)
FONT_SMALL  = ("Segoe UI", 8)
FONT_MONO   = ("Consolas", 9)
FONT_BIG    = ("Segoe UI", 22, "bold")
FONT_MEDIUM = ("Segoe UI", 11, "bold")
FONT_BTN    = ("Segoe UI", 10, "bold")

# ──────────────────────────────────────────────
#  ЛОГІКА ПОРІВНЯННЯ
# ──────────────────────────────────────────────

def get_prom_products(token, on_progress):
    headers = {
        "Authorization": f"Bearer {token}",
        "Content-Type": "application/json"
    }
    prom_skus    = set()
    prom_ext_ids = set()
    limit  = 100
    last_id = None
    page   = 1

    while True:
        params = {"limit": limit, "status": "on_display,draft,deleted"}
        if last_id:
            params["last_id"] = last_id

        resp = requests.get(
            "https://my.prom.ua/api/v1/products/list",
            headers=headers, params=params, timeout=30
        )
        resp.raise_for_status()
        data     = resp.json()
        products = data.get("products", [])
        if not products:
            break

        for p in products:
            sku    = str(p.get("sku", "")).strip()
            ext_id = str(p.get("external_id", "")).strip()
            if sku:
                prom_skus.add(sku)
            if ext_id and ext_id not in ("None", ""):
                prom_ext_ids.add(ext_id)

        on_progress(f"Пром: завантажено сторінку {page} ({len(prom_skus)} артикулів)...")
        if len(products) < limit:
            break
        last_id = products[-1]["id"]
        page   += 1
        time.sleep(0.4)

    return prom_skus, prom_ext_ids

def detect_feed_format(root):
    if root.tag == "rss" or root.find(".//item") is not None:
        return "gmc"
    if root.tag in ("yml_catalog",) or root.find(".//offer") is not None:
        return "yml"
    if "base.google.com" in ET.tostring(root, encoding="unicode")[:500]:
        return "gmc"
    return "yml"

def parse_feed(xml_path_or_url, on_progress):
    on_progress("Читання XML фіду...")

    ua = {"User-Agent": "Mozilla/5.0 (compatible; ProfetchFilter/2.0)"}

    if xml_path_or_url.startswith("http"):
        resp = requests.get(xml_path_or_url, headers=ua, timeout=60)
        resp.raise_for_status()
        content = resp.content
    else:
        with open(xml_path_or_url, "rb") as f:
            content = f.read()

    root   = ET.fromstring(content)
    fmt    = detect_feed_format(root)
    ns     = {"g": "http://base.google.com/ns/1.0"}
    items  = []

    if fmt == "gmc":
        elements = root.findall(".//item")
        on_progress(f"Формат: Google Merchant Center, {len(elements)} товарів")
        for el in elements:
            g_id   = el.findtext("g:id",  namespaces=ns, default="").strip()
            mpn    = el.findtext("g:mpn", namespaces=ns, default="").strip()
            gtin   = el.findtext("g:gtin",namespaces=ns, default="").strip()
            title  = el.findtext("g:title",namespaces=ns, default="").strip()
            avail  = el.findtext("g:availability",namespaces=ns, default="").strip()
            sku    = mpn or gtin or g_id
            items.append({"id": g_id, "sku": sku, "title": title,
                          "avail": avail, "el": el, "fmt": "gmc"})
    else:
        elements = root.findall(".//offer")
        on_progress(f"Формат: YML / Пром, {len(elements)} товарів")
        for el in elements:
            oid   = el.get("id", "").strip()
            vc    = (el.findtext("vendorCode", default="") or
                     el.findtext("article",    default="") or oid).strip()
            title = el.findtext("name", default="").strip()
            avail = el.get("available", "true")
            items.append({"id": oid, "sku": vc, "title": title,
                          "avail": avail, "el": el, "fmt": "yml"})

    return items, root, fmt

def build_output_xml(new_items, orig_root, fmt):
    """Генерує XML тільки з новими товарами."""
    
    # 1. ПОВНИЙ СЛОВНИК НАЗВ ХАРАКТЕРИСТИК (З ВАШОГО CSV-ФАЙЛУ)
    param_mapping = {
        "цвет": "Колір",
        "материал": "Матеріал",
        "особенности": "Особливості",
        "форм-фактор": "Форм-фактор",
        "состояние": "Стан",
        "гарантия": "Гарантійний термін",
        "совместимая модель": "Модель телефону",
        "модель телефона": "Модель телефону",
        "модели телефона": "Модель телефону",
        "модель": "Модель телефону",
        "совместимость с": "Сумісність з",
        "совместимый бренд": "Сумісність з",
        "совместимость с apple iphone": "Сумісність з Apple iPhone",
        "поддержка magsafe": "Підтримка MagSafe",
        "совместимость с беспроводной зарядкой": "Сумісність з бездротовою зарядкою",
        "особенность цвета": "Особливість кольору",
        "узоры и принты": "Візерунки і принти",
        "тип застежки": "Тип застібки",
        "назначение": "Призначення",
        "страна-производитель товара": "Країна виробник",
        "страна производитель": "Країна виробник",
        "страна регистрации бренда": "Країна реєстрації бренду",
        "вид": "Тип",
        "тип": "Тип",
        "комплектация": "Комплектація",
        "диагональ экрана": "Діагональ екрана",
        "выходной разъем": "Вихідний роз'єм",
        "выходной ток": "Вихідний струм",
        "интерфейс подключения": "Інтерфейс під'єднання",
        "тип кабеля": "Тип кабелю",
        "длина кабеля": "Довжина кабелю",
        "вес в упаковке, кг": "Вага в пакованні, кг",
        "высота в упаковке (см)": "Висота в пакованні (см)",
        "глубина в упаковке (см)": "Глибина в пакованні (см)",
        "клеевой слой": "Клейовий шар",
        "наличие рамки": "Наявність рамки",
        "чувствительность": "Чутливість",
        "форма разъема": "Форма роз'єму",
        "тип разъема": "Тип роз'єму",
        "тип подключения": "Тип підключення",
        "тип крепления": "Тип кріплення",
        "сопротивление": "Опір",
        "односторонняя гарнитура": "Одностороння гарнітура",
        "наушники": "Навушники",
        "минимальная воспроизводимая частота": "Мінімальна відтворювана частота",
        "максимальная воспроизводимая частота": "Максимальна відтворювана частота",
        "микрофон": "Мікрофон",
        "материал амбушюра": "Матеріал амбушюр",
        "вид наушников": "Вид навушників"
    }

    # 2. ПОВНИЙ СЛОВНИК ЗНАЧЕНЬ (ТОЧНІ ЗБІГИ)
    val_exact = {
        # Базові
        "да": "Так", "нет": "Ні", "новое": "Нове", "б/у": "Вживане",
        
        # Призначення та застібки
        "для телефона": "Для телефону", "без застежки": "Без застібки", "без узоров и принтов": "Без візерунків і принтів",
        
        # Особливість кольору
        "матовый": "Матовий", "глянцевый": "Глянсовий", "прозрачный": "Прозорий",

        # Форм-фактор
        "панель (накладка на корпус)": "Панель (Накладка на корпус)", "панель": "Панель (Накладка на корпус)",
        "накладка": "Панель (Накладка на корпус)", "чехол-книжка": "Чохол-книжка", "бампер": "Бампер",
        "защитное стекло": "Захисне скло", "защитная пленка": "Захисна плівка", "вкладыши": "Вкладиші",

        # Матеріали (точні збіги)
        "тпу (термопластичный полиуретан)": "ТПУ (термопластичний поліуретан)", "силикон + пластик": "Силікон + пластик",
        "поликарбонат": "Полікарбонат", "силикон": "Силікон",

        # Країни
        "китай": "Китай", "вьетнам": "В'єтнам", "южная корея": "Південна Корея", "сша": "США",

        # Роз'єми та інше
        "прямой": "Прямий", "проводные": "Провідні", "без крепления": "Без кріплення", "встроенный": "Вбудований",

        # Англійські кольори з файлів
        "matt black": "Чорний", "black": "Чорний", "camo black": "Чорний", 
        "kevlar black": "Чорний", "smoke black": "Чорний", "black geo": "Чорний",
        "ash": "Сірий", "ash grey": "Сірий", "carbon fiber": "Сірий", "carbon": "Сірий",
        "titanium": "Сірий", "slate": "Сірий", "grey": "Сірий", "gray": "Сірий",
        "olive": "Зелений", "olive drab": "Оливковий", "olive green": "Зелений", "khaki": "Хакі",
        "dark orange": "Помаранчевий", "sunset": "Помаранчевий", "brown": "Коричневий",
        "clear": "Прозорий", "transparent": "Прозорий", "white": "Білий", 
        "red": "Червоний", "crimson": "Червоний", "magma": "Червоний", "blue": "Синій", 
        "mallard": "Синій", "cobalt": "Синій", "green": "Зелений", "spearmint": "М'ятний", 
        "pink": "Рожевий", "silver": "Сріблястий", "lilac": "Світло-фіолетовий",

        # Російські кольори
        "черный": "Чорний", "белый": "Білий", "красный": "Червоний", "синий": "Синій",
        "зеленый": "Зелений", "розовый": "Рожевий", "серый": "Сірий", "светло фиолетовый": "Світло-фіолетовий",
        "серебристый": "Сріблястий", "оливковый": "Оливковий", "фиолетовый": "Фіолетовий",
        "оранжевый": "Помаранчевий", "мятный": "М'ятний", "голубой": "Блакитний"
    }

    # 3. СЛОВНИК СКЛАДНИХ ЗНАЧЕНЬ ТА ОСОБЛИВОСТЕЙ (Заміна шматочків тексту)
    val_replace = {
        "Для занятий спортом": "Для занять спортом",
        "Противоударный": "Протиударний",
        "Усиленные борты чехла": "Посилені борти чохла",
        "Защита камеры": "Захист камери",
        "Рифленая текстура": "Рифлена текстура",
        "Функция подставки": "Функція підставки",
        "Покрытие soft touch": "Покриття soft touch",
        "Магнитный": "Магнітний",
        "Защита всего корпуса": "Захист всього корпусу",
        "Силикон": "Силікон",
        "Поликарбонат": "Полікарбонат",
        "Термополиуретан": "ТПУ (термопластичний поліуретан)",
        "Арамид": "Арамід",
        "Кевлар": "Арамід",
        "Металл": "Метал",
        "Алюминий": "Метал",
        "Пластик": "Пластик",
        "Стекло": "Скло",
        "Искусственная кожа": "Штучна шкіра",
        "Кожа": "Штучна шкіра"
    }

    for item in new_items:
        # Цикл проходить по ВСІХ тегах товару
        for child in item["el"].iter():
            tag_name = child.tag.split('}')[-1]
            
            # Обробка характеристик (для різних форматів файлу)
            if tag_name in ['param', 'attribute_name', 'attribute_value', 'color']:
                
                # Відрізаємо сміттєві ID постачальника, щоб Пром не бракував поле
                if 'paramid' in child.attrib:
                    del child.attrib['paramid']
                if 'valueid' in child.attrib:
                    del child.attrib['valueid']
                
                # Переклад НАЗВ характеристик
                if tag_name in ['param', 'attribute_name']:
                    name = child.get("name") or child.text
                    if name:
                        clean_name = name.strip().lower()
                        if clean_name in param_mapping:
                            if tag_name == 'param':
                                child.set("name", param_mapping[clean_name])
                            else:
                                child.text = param_mapping[clean_name]
                
                # Переклад ЗНАЧЕНЬ характеристик
                if tag_name in ['param', 'attribute_value', 'color'] and child.text:
                    val = child.text.strip()
                    if val.lower() in val_exact:
                        child.text = val_exact[val.lower()]
                    else:
                        new_val = child.text
                        for rus, ukr in val_replace.items():
                            new_val = new_val.replace(rus, ukr)
                        child.text = new_val

    # Збірка фінального файлу
    if fmt == "gmc":
        rss = ET.Element("rss", {"xmlns:g": "http://base.google.com/ns/1.0", "version": "2.0"})
        ch = ET.SubElement(rss, "channel")
        oc = orig_root.find("channel")
        if oc is not None:
            for tag in ["title", "link"]:
                el = oc.find(tag)
                if el is not None:
                    ne = ET.SubElement(ch, tag)
                    ne.text = el.text
        for item in new_items:
            ch.append(item["el"])
        out_root = rss
    else:
        out_root = ET.Element("yml_catalog", {"date": orig_root.get("date", "")})
        shop = ET.SubElement(out_root, "shop")
        os_  = orig_root.find("shop")
        if os_ is not None:
            for tag in ["name", "company", "url", "currencies", "categories"]:
                el = os_.find(tag)
                if el is not None:
                    shop.append(el)
        offers_el = ET.SubElement(shop, "offers")
        for item in new_items:
            offers_el.append(item["el"])

    raw    = ET.tostring(out_root, encoding="unicode")
    parsed = minidom.parseString(raw)
    return parsed.toprettyxml(indent="  ", encoding="utf-8")

# ──────────────────────────────────────────────
#  ГОЛОВНИЙ КЛАС ДОДАТКУ
# ──────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Profetch — Prom Filter")
        self.geometry("680x720")
        self.minsize(620, 640)
        self.configure(bg=BG)
        self.resizable(True, True)

        self.feed_path  = tk.StringVar(value="")
        self.prom_token = tk.StringVar(value="")
        self.out_path   = tk.StringVar(value="")
        self._running   = False

        self._build_ui()
        self._load_settings()

    def _build_ui(self):
        hdr = tk.Frame(self, bg=BG, pady=18)
        hdr.pack(fill="x", padx=24)

        tk.Label(hdr, text="⬡", font=("Segoe UI", 20), fg=ACCENT, bg=BG).pack(side="left")
        tk.Label(hdr, text="  PROFETCH", font=FONT_TITLE, fg=TEXT, bg=BG).pack(side="left")
        tk.Label(hdr, text="  Prom Filter", font=("Segoe UI", 10), fg=TEXT_DIM, bg=BG).pack(side="left", pady=3)

        self._sep()

        self._section("1  XML-фід (ваш каталог)")
        feed_row = tk.Frame(self, bg=BG2, bd=0)
        feed_row.pack(fill="x", padx=24, pady=(0, 4))

        self.feed_entry = self._entry(feed_row, self.feed_path, width=52)
        self.feed_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self._btn_small(feed_row, "Огляд…", self._pick_feed).pack(side="left")

        tk.Label(self, text="або вставте URL-посилання на фід", font=FONT_SMALL, fg=TEXT_DIM, bg=BG).pack(anchor="w", padx=28)

        self._section("2  API-токен Пром.юа")
        tok_row = tk.Frame(self, bg=BG)
        tok_row.pack(fill="x", padx=24, pady=(0, 4))

        self.tok_entry = self._entry(tok_row, self.prom_token, width=52, show="•")
        self.tok_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self._btn_small(tok_row, "👁", self._toggle_token).pack(side="left")

        tk.Label(self, text="Пром → Налаштування → Управління API-токенами", font=FONT_SMALL, fg=TEXT_DIM, bg=BG).pack(anchor="w", padx=28)

        self._section("3  Зберегти результат як")
        out_row = tk.Frame(self, bg=BG)
        out_row.pack(fill="x", padx=24, pady=(0, 4))

        self.out_entry = self._entry(out_row, self.out_path, width=52)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self._btn_small(out_row, "Огляд…", self._pick_out).pack(side="left")

        self._sep()

        self.run_btn = tk.Button(
            self, text="▶   ЗНАЙТИ НОВІ ТОВАРИ", font=FONT_BTN, bg=ACCENT, fg="#ffffff",
            activebackground=ACCENT2, activeforeground="#ffffff", bd=0, cursor="hand2", padx=28, pady=12,
            command=self._start
        )
        self.run_btn.pack(fill="x", padx=24, pady=(12, 4))

        self.pb_frame = tk.Frame(self, bg=BG, height=6)
        self.pb_frame.pack(fill="x", padx=24, pady=(0, 8))
        self.pb_frame.pack_propagate(False)

        self.pb_bg = tk.Frame(self.pb_frame, bg=BG3, height=6)
        self.pb_bg.pack(fill="both", expand=True)

        self.pb_fill = tk.Frame(self.pb_bg, bg=ACCENT, height=6, width=0)
        self.pb_fill.place(x=0, y=0, height=6)
        self._pb_anim = 0
        self._pb_dir  = 1

        log_hdr = tk.Frame(self, bg=BG)
        log_hdr.pack(fill="x", padx=24)
        tk.Label(log_hdr, text="Журнал", font=FONT_SMALL, fg=TEXT_DIM, bg=BG).pack(side="left")

        self.log_box = tk.Text(
            self, font=FONT_MONO, bg=BG2, fg=TEXT, insertbackground=TEXT, bd=0, pady=10, padx=12,
            relief="flat", state="disabled", height=9, selectbackground=BG3
        )
        self.log_box.pack(fill="both", expand=True, padx=24, pady=(4, 0))

        stat_frame = tk.Frame(self, bg=BG3, pady=14)
        stat_frame.pack(fill="x", padx=24, pady=8)

        self.stat_total = self._stat_cell(stat_frame, "У фіді",  "—")
        self.stat_prom  = self._stat_cell(stat_frame, "На Проме", "—")
        self.stat_new   = self._stat_cell(stat_frame, "НОВИХ",   "—", color=SUCCESS)

        for col in range(3):
            stat_frame.columnconfigure(col, weight=1)

        self.open_btn = tk.Button(
            self, text="📂  Відкрити папку з результатом", font=FONT_SMALL, bg=BG2, fg=TEXT_DIM,
            activebackground=BG3, activeforeground=TEXT, bd=0, cursor="hand2", padx=16, pady=8,
            command=self._open_folder, state="disabled"
        )
        self.open_btn.pack(pady=(0, 16))

    def _sep(self):
        tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=24, pady=8)

    def _section(self, text):
        tk.Label(self, text=text, font=("Segoe UI", 9, "bold"), fg=ACCENT, bg=BG).pack(anchor="w", padx=24, pady=(12, 4))

    def _entry(self, parent, var, width=40, show=None):
        kw = dict(textvariable=var, font=FONT_LABEL, bg=BG3, fg=TEXT, insertbackground=TEXT, bd=0, relief="flat", highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT, width=width)
        if show:
            kw["show"] = show
        e = tk.Entry(parent, **kw)
        return e

    def _btn_small(self, parent, text, cmd):
        return tk.Button(parent, text=text, font=FONT_SMALL, bg=BG3, fg=TEXT_DIM, activebackground=BORDER, activeforeground=TEXT, bd=0, cursor="hand2", padx=10, pady=6, relief="flat", command=cmd, highlightthickness=1, highlightbackground=BORDER)

    def _stat_cell(self, parent, label, value, color=TEXT):
        frame = tk.Frame(parent, bg=BG3)
        frame.grid(row=0, column=len(parent.winfo_children()) - 1, sticky="nsew", padx=6)
        lbl = tk.Label(frame, text=label, font=FONT_SMALL, fg=TEXT_DIM, bg=BG3)
        lbl.pack(pady=(8, 0))
        val = tk.Label(frame, text=value, font=FONT_BIG, fg=color, bg=BG3)
        val.pack(pady=(0, 8))
        return val

    def _toggle_token(self):
        cur = self.tok_entry.cget("show")
        self.tok_entry.configure(show="" if cur else "•")

    def _pick_feed(self):
        path = filedialog.askopenfilename(title="Виберіть XML-фід", filetypes=[("XML файли", "*.xml *.yml"), ("Всі файли", "*.*")])
        if path:
            self.feed_path.set(path)
            base = os.path.splitext(path)[0]
            self.out_path.set(base + "-new.xml")

    def _pick_out(self):
        path = filedialog.asksaveasfilename(title="Зберегти як", defaultextension=".xml", filetypes=[("XML файли", "*.xml"), ("Всі файли", "*.*")])
        if path:
            self.out_path.set(path)

    def _open_folder(self):
        path = self.out_path.get()
        if path and os.path.exists(path):
            folder = os.path.dirname(path)
            os.startfile(folder) if sys.platform == "win32" else None

    def _log(self, text, color=None):
        self.log_box.configure(state="normal")
        tag = None
        if color:
            tag = f"color_{color}"
            self.log_box.tag_configure(tag, foreground=color)
        self.log_box.insert("end", text + "\n", tag or "")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _log_clear(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _pb_start(self):
        self._pb_anim = 0
        self._pb_dir  = 1
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
            self._pb_dir = 1
            self._pb_anim = 0
        self.pb_fill.place(x=self._pb_anim, y=0, height=6, width=bar_w)
        self.after(20, self._pb_tick)

    def _settings_path(self):
        return os.path.join(os.path.expanduser("~"), ".profetch_filter.ini")

    def _save_settings(self):
        try:
            with open(self._settings_path(), "w", encoding="utf-8") as f:
                f.write(f"token={self.prom_token.get()}\n")
                f.write(f"feed={self.feed_path.get()}\n")
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
                        self.feed_path.set(v)
                    elif k == "out":
                        self.out_path.set(v)
        except Exception:
            pass

    def _start(self):
        feed  = self.feed_path.get().strip()
        token = self.prom_token.get().strip()
        out   = self.out_path.get().strip()

        if not feed:
            messagebox.showwarning("Увага", "Виберіть або вставте URL XML-фіду")
            return
        if not token:
            messagebox.showwarning("Увага", "Введіть API-токен Пром.юа")
            return
        if not out:
            messagebox.showwarning("Увага", "Вкажіть шлях для збереження результату")
            return

        self._save_settings()
        self._log_clear()
        self._running = True
        self.run_btn.configure(state="disabled", text="⏳  Обробка...")
        self.open_btn.configure(state="disabled")
        self.stat_total.configure(text="…", fg=TEXT_DIM)
        self.stat_prom.configure(text="…",  fg=TEXT_DIM)
        self.stat_new.configure(text="…",   fg=TEXT_DIM)
        self._pb_start()

        threading.Thread(target=self._worker,
                         args=(feed, token, out), daemon=True).start()

    def _worker(self, feed, token, out):
        def log(msg, color=None):
            self.after(0, lambda: self._log(msg, color))

        try:
            log("⟳  Підключення до Пром.юа API...")
            prom_skus, prom_ext_ids = get_prom_products(token, lambda m: log(f"   {m}"))
            log(f"✓  Пром: {len(prom_skus)} унікальних артикулів", SUCCESS)

            log("")
            log("⟳  Читання XML-фіду...")
            feed_items, orig_root, fmt = parse_feed(feed, lambda m: log(f"   {m}"))
            total = len(feed_items)
            log(f"✓  Фід: {total} товарів (формат: {fmt.upper()})", SUCCESS)

            log("")
            log("⟳  Порівняння...")
            new_items = []
            for item in feed_items:
                on_prom = (
                    item["sku"]  in prom_skus    or
                    item["id"]   in prom_ext_ids or
                    item["sku"]  in prom_ext_ids
                )
                if not on_prom:
                    new_items.append(item)

            already = total - len(new_items)
            log(f"✓  Порівняння завершено", SUCCESS)
            log(f"   У фіді:    {total}")
            log(f"   На Пром:   {already}")
            log(f"   НОВИХ:     {len(new_items)}", SUCCESS if new_items else TEXT_DIM)

            self.after(0, lambda: [
                self.stat_total.configure(text=str(total), fg=TEXT),
                self.stat_prom.configure(text=str(already), fg=TEXT),
                self.stat_new.configure(
                    text=str(len(new_items)),
                    fg=SUCCESS if new_items else TEXT_DIM
                )
            ])

            if new_items:
                log("")
                log("⟳  Генерація XML...")
                xml_bytes = build_output_xml(new_items, orig_root, fmt)
                os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
                with open(out, "wb") as f:
                    f.write(xml_bytes.encode('utf-8') if isinstance(xml_bytes, str) else xml_bytes)
                size_kb = os.path.getsize(out) / 1024
                log(f"✓  Збережено: {os.path.basename(out)} ({size_kb:.1f} КБ)", SUCCESS)
                log("")
                log("─" * 42, TEXT_DIM)
                log("ЩО РОБИТИ ДАЛІ:", ACCENT)
                log("  1. Пром → Товари → Імпорт", TEXT)
                log("  2. Завантажте: " + os.path.basename(out), TEXT)
                log("  3. БЕЗ галочки «Тільки оновлення»", WARNING)
                log("  4. Галочки: Ціна + Наявність", TEXT)
                log("  5. Натисніть «Почати імпорт»", TEXT)
                log("─" * 42, TEXT_DIM)
                self.after(0, lambda: self.open_btn.configure(state="normal"))
            else:
                log("")
                log("✓  Нових товарів немає — все вже на Пром!", SUCCESS)

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
        finally:
            self._running = False
            self.after(0, lambda: self.run_btn.configure(
                state="normal", text="▶   ЗНАЙТИ НОВІ ТОВАРИ"
            ))

if __name__ == "__main__":
    app = App()
    app.mainloop()
