#!/usr/bin/env python3
"""
Profetch — Prom New Products Filter
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
BG = "#0f1117"
BG2 = "#1a1d27"
BG3 = "#22263a"
ACCENT = "#4f8ef7"
ACCENT2 = "#7c5cfc"
SUCCESS = "#3ecf8e"
WARNING = "#f7b731"
DANGER = "#fc5c65"
TEXT = "#e8eaf6"
TEXT_DIM = "#6b7280"
BORDER = "#2d3250"

FONT_TITLE = ("Segoe UI", 13, "bold")
FONT_LABEL = ("Segoe UI", 9)
FONT_SMALL = ("Segoe UI", 8)
FONT_MONO = ("Consolas", 9)
FONT_BIG = ("Segoe UI", 22, "bold")
FONT_BTN = ("Segoe UI", 10, "bold")

# ──────────────────────────────────────────────
#  ЛОГІКА ПОРІВНЯННЯ ТА ПЕРЕКЛАДУ
# ──────────────────────────────────────────────

def get_prom_products(token, on_progress):
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    prom_skus, prom_ext_ids = set(), set()
    limit, last_id, page = 100, None, 1
    while True:
        params = {"limit": limit, "status": "on_display,draft,deleted"}
        if last_id: params["last_id"] = last_id
        resp = requests.get("https://my.prom.ua/api/v1/products/list", headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        products = data.get("products", [])
        if not products: break
        for p in products:
            sku, ext_id = str(p.get("sku", "")).strip(), str(p.get("external_id", "")).strip()
            if sku: prom_skus.add(sku)
            if ext_id and ext_id not in ("None", ""): prom_ext_ids.add(ext_id)
        on_progress(f"Пром: завантажено сторінку {page} ({len(prom_skus)} артикулів)...")
        page += 1
        if len(products) < limit: break
    return prom_skus, prom_ext_ids

def detect_feed_format(root):
    if root.tag == "rss" or root.find(".//item") is not None: return "gmc"
    if root.tag in ("yml_catalog",) or root.find(".//offer") is not None: return "yml"
    if "base.google.com" in ET.tostring(root, encoding="unicode")[:500]: return "gmc"
    return "yml"

def parse_feed(xml_path_or_url, on_progress):
    on_progress("Читання XML фіду...")
    ua = {"User-Agent": "Mozilla/5.0"}
    if xml_path_or_url.startswith("http"):
        resp = requests.get(xml_path_or_url, headers=ua, timeout=60)
        resp.raise_for_status()
        content = resp.content
    else:
        with open(xml_path_or_url, "rb") as f: content = f.read()

    root = ET.fromstring(content)
    fmt = detect_feed_format(root)
    ns = {"g": "http://base.google.com/ns/1.0"}
    items = []

    if fmt == "gmc":
        elements = root.findall(".//item")
        on_progress(f"Формат: Google Merchant, {len(elements)} товарів")
        for el in elements:
            g_id = el.findtext("g:id", namespaces=ns, default="").strip()
            mpn = el.findtext("g:mpn", namespaces=ns, default="").strip()
            gtin = el.findtext("g:gtin", namespaces=ns, default="").strip()
            title = el.findtext("g:title", namespaces=ns, default="").strip()
            sku = mpn or gtin or g_id
            items.append({"id": g_id, "sku": sku, "title": title, "el": el, "fmt": "gmc"})
    else:
        elements = root.findall(".//offer")
        on_progress(f"Формат: YML / Пром, {len(elements)} товарів")
        for el in elements:
            oid = el.get("id", "").strip()
            vc = (el.findtext("vendorCode", default="") or el.findtext("article", default="") or oid).strip()
            title = el.findtext("name", default="").strip()
            items.append({"id": oid, "sku": vc, "title": title, "el": el, "fmt": "yml"})
    return items, root, fmt

def build_output_xml(new_items, orig_root, fmt):
    # 1. СЛОВНИК НАЗВ
    param_mapping = {
        "цвет": "Колір", "материал": "Матеріал", "особенности": "Особливості",
        "форм-фактор": "Форм-фактор", "состояние": "Стан", "гарантия": "Гарантійний термін",
        "совместимая модель": "Модель телефону", "модель телефона": "Модель телефону",
        "модель": "Модель телефону", "совместимость с": "Сумісність з",
        "совместимый бренд": "Сумісність з", "совместимость с apple iphone": "Сумісність з Apple iPhone",
        "поддержка magsafe": "Підтримка MagSafe", "совместимость с беспроводной зарядкой": "Сумісність з бездротовою зарядкою",
        "особенность цвета": "Особливість кольору", "узоры и принты": "Візерунки і принти",
        "тип застежки": "Тип застібки", "назначение": "Призначення", "вид": "Тип", "тип": "Тип",
        "страна-производитель товара": "Країна виробник", "страна производитель": "Країна виробник",
        "комплектация": "Комплектація", "где находится товар": "Де знаходиться товар", 
        "где_находится_товар": "Де знаходиться товар", "код товара": "Код", "артикул": "Код"
    }

    # 2. СЛОВНИК ТОЧНИХ ЗНАЧЕНЬ
    val_exact = {
        # Особливості
        "противоударный": "Протиударний", "рифленая текстура": "Рифлена текстура",
        "покрытие soft touch": "Покриття soft touch", "усиленные борты чехла": "Посилені борти чохла",
        "для занятий спортом": "Для занять спортом", "защита камеры": "Захист камери",
        "с подставкой": "З підставкою", "с кольцом": "З кільцем", "функция подставки": "Функція підставки",
        "магнитный": "Магнітний", "защита всего корпуса": "Захист всього корпусу",
        "возможность использования магнитных держателей": "Можливість використання магнітних тримачів",
        "возможность использовать беспроводное зу": "Можливість використовувати бездротовий ЗП",
        "поддержка бесконтактных платежей": "Підтримка безконтактних платежів",
        
        # Застібки та інше
        "без застежки": "Без застібки", "без узоров и принтов": "Без візерунків і принтів",
        "да": "Так", "нет": "Ні", "для телефона": "Для телефону",
        
        # Форм-фактор
        "матовый": "Матовий", "глянцевый": "Глянсовий", "прозрачный": "Прозорий",
        "панель (накладка на корпус)": "Панель (Накладка на корпус)", "панель": "Панель (Накладка на корпус)",
        "накладка": "Панель (Накладка на корпус)", "чехол-книжка": "Чохол-книжка", "бампер": "Бампер",
        "защитное стекло": "Захисне скло", "защитная пленка": "Захисна плівка",

        # Матеріали
        "тпу (термопластичный полиуретан)": "ТПУ (термопластичний поліуретан)", 
        "силикон + пластик": "Силікон + пластик", "поликарбонат": "Полікарбонат", "силикон": "Силікон",

        # Кольори (вибірка з ваших файлів)
        "matt black": "Чорний", "black": "Чорний", "camo black": "Чорний", "kevlar black": "Чорний", 
        "smoke black": "Чорний", "черный": "Чорний", "ash": "Сірий", "carbon fiber": "Сірий", 
        "titanium": "Сірий", "grey": "Сірий", "gray": "Сірий", "серый": "Сірий",
        "olive": "Зелений", "olive drab": "Оливковий", "olive green": "Зелений", "зеленый": "Зелений",
        "dark orange": "Помаранчевий", "sunset": "Помаранчевий", "brown": "Коричневий",
        "clear": "Прозорий", "transparent": "Прозорий", "white": "Білий", "белый": "Білий",
        "red": "Червоний", "crimson": "Червоний", "красный": "Червоний", "blue": "Синій", 
        "синий": "Синій", "pink": "Рожевий", "розовый": "Рожевий"
    }

    # 3. СЛОВНИК ЗАМІНИ ТЕКСТУ (для складових матеріалів)
    val_replace = {
        "Силикон": "Силікон", "силикон": "силікон",
        "Поликарбонат": "Полікарбонат", "поликарбонат": "полікарбонат",
        "Термополиуретан": "ТПУ", "термополиуретан": "ТПУ",
        "Металл": "Метал", "металл": "метал",
        "Пластик": "Пластик", "пластик": "пластик",
        "Стекло": "Скло", "стекло": "скло",
        "Кожа": "Штучна шкіра", "кожа": "штучна шкіра"
    }

    for item in new_items:
        for child in item["el"].iter():
            tag_name = child.tag.split('}')[-1]
            
            if tag_name in ['param', 'attribute_name', 'attribute_value', 'color']:
                # Видалення ID
                if 'paramid' in child.attrib: del child.attrib['paramid']
                if 'valueid' in child.attrib: del child.attrib['valueid']
                
                # Переклад назв
                if tag_name in ['param', 'attribute_name']:
                    name = child.get("name") or child.text
                    if name:
                        clean_name = name.strip().lower()
                        if clean_name in param_mapping:
                            if tag_name == 'param': child.set("name", param_mapping[clean_name])
                            else: child.text = param_mapping[clean_name]
                
                # ПЕРЕКЛАД ЗНАЧЕНЬ ТА ОБРОБКА СИМВОЛУ "|"
                if tag_name in ['param', 'attribute_value', 'color'] and child.text:
                    original_val = child.text.strip()
                    
                    # Замінюємо "трубу" | на кому, щоб Пром побачив список
                    temp_val = original_val.replace('|', ',')
                    
                    # Розбиваємо за комами, щоб перекласти кожну особливість окремо
                    parts = [p.strip() for p in temp_val.split(',') if p.strip()]
                    
                    translated_parts = []
                    for part in parts:
                        lower_part = part.lower()
                        # Шукаємо точний переклад для "противоударный", "защита камеры" тощо
                        if lower_part in val_exact:
                            translated_parts.append(val_exact[lower_part])
                        else:
                            # Якщо немає точного збігу, робимо заміну букв (наприклад для матеріалів)
                            t = part
                            for r, u in val_replace.items():
                                t = t.replace(r, u)
                            translated_parts.append(t)
                    
                    # Збираємо назад у рядок, розділений КОМАМИ
                    child.text = ", ".join(translated_parts)

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
        self.feed_path = tk.StringVar(value="")
        self.prom_token = tk.StringVar(value="")
        self.out_path = tk.StringVar(value="")
        self._running = False
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

        self._section("2  API-токен Пром.юа")
        tok_row = tk.Frame(self, bg=BG)
        tok_row.pack(fill="x", padx=24, pady=(0, 4))
        self.tok_entry = self._entry(tok_row, self.prom_token, width=52, show="•")
        self.tok_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self._btn_small(tok_row, "👁", self._toggle_token).pack(side="left")

        self._section("3  Зберегти результат як")
        out_row = tk.Frame(self, bg=BG)
        out_row.pack(fill="x", padx=24, pady=(0, 4))
        self.out_entry = self._entry(out_row, self.out_path, width=52)
        self.out_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))
        self._btn_small(out_row, "Огляд…", self._pick_out).pack(side="left")
        self._sep()

        self.run_btn = tk.Button(self, text="▶   ЗНАЙТИ НОВІ ТОВАРИ", font=FONT_BTN, bg=ACCENT, fg="#ffffff", activebackground=ACCENT2, activeforeground="#ffffff", bd=0, cursor="hand2", padx=28, pady=12, command=self._start)
        self.run_btn.pack(fill="x", padx=24, pady=(12, 4))

        self.pb_frame = tk.Frame(self, bg=BG, height=6)
        self.pb_frame.pack(fill="x", padx=24, pady=(0, 8))
        self.pb_frame.pack_propagate(False)
        self.pb_bg = tk.Frame(self.pb_frame, bg=BG3, height=6)
        self.pb_bg.pack(fill="both", expand=True)
        self.pb_fill = tk.Frame(self.pb_bg, bg=ACCENT, height=6, width=0)
        self.pb_fill.place(x=0, y=0, height=6)
        self._pb_anim = 0
        self._pb_dir = 1

        self.log_box = tk.Text(self, font=FONT_MONO, bg=BG2, fg=TEXT, insertbackground=TEXT, bd=0, pady=10, padx=12, relief="flat", state="disabled", height=9, selectbackground=BG3)
        self.log_box.pack(fill="both", expand=True, padx=24, pady=(4, 0))

        stat_frame = tk.Frame(self, bg=BG3, pady=14)
        stat_frame.pack(fill="x", padx=24, pady=8)
        self.stat_total = self._stat_cell(stat_frame, "У фіді",  "—")
        self.stat_prom = self._stat_cell(stat_frame, "На Проме", "—")
        self.stat_new = self._stat_cell(stat_frame, "НОВИХ", "—", color=SUCCESS)
        for col in range(3): stat_frame.columnconfigure(col, weight=1)

        self.open_btn = tk.Button(self, text="📂  Відкрити папку з результатом", font=FONT_SMALL, bg=BG2, fg=TEXT_DIM, activebackground=BG3, activeforeground=TEXT, bd=0, cursor="hand2", padx=16, pady=8, command=self._open_folder, state="disabled")
        self.open_btn.pack(pady=(0, 16))

    def _sep(self): tk.Frame(self, bg=BORDER, height=1).pack(fill="x", padx=24, pady=8)
    def _section(self, text): tk.Label(self, text=text, font=("Segoe UI", 9, "bold"), fg=ACCENT, bg=BG).pack(anchor="w", padx=24, pady=(12, 4))
    def _entry(self, parent, var, width=40, show=None):
        kw = dict(textvariable=var, font=FONT_LABEL, bg=BG3, fg=TEXT, insertbackground=TEXT, bd=0, relief="flat", highlightthickness=1, highlightbackground=BORDER, highlightcolor=ACCENT, width=width)
        if show: kw["show"] = show
        return tk.Entry(parent, **kw)
    def _btn_small(self, parent, text, cmd): return tk.Button(parent, text=text, font=FONT_SMALL, bg=BG3, fg=TEXT_DIM, activebackground=BORDER, activeforeground=TEXT, bd=0, cursor="hand2", padx=10, pady=6, relief="flat", command=cmd, highlightthickness=1, highlightbackground=BORDER)
    def _stat_cell(self, parent, label, value, color=TEXT):
        frame = tk.Frame(parent, bg=BG3)
        frame.grid(row=0, column=len(parent.winfo_children()) - 1, sticky="nsew", padx=6)
        tk.Label(frame, text=label, font=FONT_SMALL, fg=TEXT_DIM, bg=BG3).pack(pady=(8, 0))
        val = tk.Label(frame, text=value, font=FONT_BIG, fg=color, bg=BG3)
        val.pack(pady=(0, 8))
        return val
    def _toggle_token(self):
        cur = self.tok_entry.cget("show")
        self.tok_entry.configure(show="" if cur else "•")
    def _pick_feed(self):
        path = filedialog.askopenfilename(filetypes=[("XML файли", "*.xml *.yml"), ("Всі файли", "*.*")])
        if path:
            self.feed_path.set(path)
            self.out_path.set(os.path.splitext(path)[0] + "-new.xml")
    def _pick_out(self):
        path = filedialog.asksaveasfilename(defaultextension=".xml", filetypes=[("XML файли", "*.xml")])
        if path: self.out_path.set(path)
    def _open_folder(self):
        path = self.out_path.get()
        if path and os.path.exists(path):
            os.startfile(os.path.dirname(path)) if sys.platform == "win32" else None
    def _log(self, text, color=None):
        self.log_box.configure(state="normal")
        tag = f"color_{color}" if color else None
        if color: self.log_box.tag_configure(tag, foreground=color)
        self.log_box.insert("end", text + "\n", tag or "")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
    def _log_clear(self):
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")
    def _pb_start(self): self._pb_anim = 0; self._pb_dir = 1; self._pb_tick()
    def _pb_tick(self):
        if not self._running:
            self.pb_fill.place(x=0, y=0, height=6, width=0); return
        w = self.pb_bg.winfo_width()
        bar_w = max(w // 4, 60)
        self._pb_anim += self._pb_dir * 6
        if self._pb_anim >= w - bar_w: self._pb_dir = -1
        if self._pb_anim <= 0: self._pb_dir = 1; self._pb_anim = 0
        self.pb_fill.place(x=self._pb_anim, y=0, height=6, width=bar_w)
        self.after(20, self._pb_tick)
    def _settings_path(self): return os.path.join(os.path.expanduser("~"), ".profetch_filter.ini")
    def _save_settings(self):
        try:
            with open(self._settings_path(), "w", encoding="utf-8") as f:
                f.write(f"token={self.prom_token.get()}\nfeed={self.feed_path.get()}\nout={self.out_path.get()}\n")
        except: pass
    def _load_settings(self):
        try:
            if not os.path.exists(self._settings_path()): return
            with open(self._settings_path(), encoding="utf-8") as f:
                for line in f:
                    k, _, v = line.strip().partition("=")
                    if k == "token": self.prom_token.set(v)
                    elif k == "feed": self.feed_path.set(v)
                    elif k == "out": self.out_path.set(v)
        except: pass

    def _start(self):
        feed, token, out = self.feed_path.get().strip(), self.prom_token.get().strip(), self.out_path.get().strip()
        if not feed or not token or not out: return messagebox.showwarning("Увага", "Заповніть усі поля")
        self._save_settings(); self._log_clear(); self._running = True
        self.run_btn.configure(state="disabled", text="⏳  Обробка..."); self.open_btn.configure(state="disabled")
        self._pb_start()
        threading.Thread(target=self._worker, args=(feed, token, out), daemon=True).start()

    def _worker(self, feed, token, out):
        def log(msg, color=None): self.after(0, lambda: self._log(msg, color))
        try:
            log("⟳  Підключення до Пром.юа API..."); prom_skus, prom_ext_ids = get_prom_products(token, lambda m: log(f"   {m}"))
            log(f"✓  Пром: {len(prom_skus)} артикулів", SUCCESS)
            log("⟳  Читання XML-фіду...")
            feed_items, orig_root, fmt = parse_feed(feed, lambda m: log(f"   {m}"))
            total = len(feed_items)
            log("⟳  Порівняння...")
            new_items = [item for item in feed_items if not (item["sku"] in prom_skus or item["id"] in prom_ext_ids or item["sku"] in prom_ext_ids)]
            already = total - len(new_items)
            self.after(0, lambda: [
                self.stat_total.configure(text=str(total), fg=TEXT),
                self.stat_prom.configure(text=str(already), fg=TEXT),
                self.stat_new.configure(text=str(len(new_items)), fg=SUCCESS if new_items else TEXT_DIM)
            ])
            if new_items:
                log("⟳  Генерація XML...")
                xml_bytes = build_output_xml(new_items, orig_root, fmt)
                os.makedirs(os.path.dirname(os.path.abspath(out)), exist_ok=True)
                with open(out, "wb") as f: f.write(xml_bytes.encode('utf-8') if isinstance(xml_bytes, str) else xml_bytes)
                log(f"✓  Збережено: {os.path.basename(out)}", SUCCESS)
                self.after(0, lambda: self.open_btn.configure(state="normal"))
            else:
                log("✓  Нових товарів немає!", SUCCESS)
        except Exception as e: log(f"✗  Помилка: {e}", DANGER)
        finally:
            self._running = False
            self.after(0, lambda: self.run_btn.configure(state="normal", text="▶   ЗНАЙТИ НОВІ ТОВАРИ"))

if __name__ == "__main__":
    app = App()
    app.mainloop()
