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
    """Завантажує всі товари з Пром через API."""
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
    """Визначає формат XML: 'gmc' (Google Merchant) або 'yml'."""
    if root.tag == "rss" or root.find(".//item") is not None:
        return "gmc"
    if root.tag in ("yml_catalog",) or root.find(".//offer") is not None:
        return "yml"
    # спроба по namespace
    if "base.google.com" in ET.tostring(root, encoding="unicode")[:500]:
        return "gmc"
    return "yml"


def parse_feed(xml_path_or_url, on_progress):
    """Парсить XML фід (файл або URL) і повертає список товарів та корінь."""
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
    ns_gmc = {"g": "http://base.google.com/ns/1.0"}

    # =========================================================================
    # ВЕЛИКИЙ СЛОВНИК ПЕРЕКЛАДІВ ПРОМ (РОЗШИРЕНИЙ З УРАХУВАННЯМ ФАЙЛУ GMC)
    # =========================================================================
    param_mapping = {
        "цвет": "Колір",
        "материал": "Матеріал",
        "особенности": "Особливості",
        "форм-фактор": "Форм-фактор",
        "состояние": "Стан",
        "гарантия": "Гарантійний термін",
        "комплектация": "Комплектація",
        "в комплекте": "У комплекті",
        "вид": "Вид",
        "тип": "Тип",
        "совместимость с": "Сумісність з",
        "совместимость с брендом": "Сумісність з",
        "совместимая модель": "Модель телефону",
        "модель телефона": "Модель телефону",
        "модели телефона": "Модель телефону",
        "модель": "Модель телефону",
        "мoдель": "Модель телефону", 
        "совместимая серия": "Серія телефону",
        "совместимость с apple iphone": "Сумісність з Apple iPhone",
        "поддержка magsafe": "Підтримка MagSafe",
        "диагональ экрана": "Діагональ екрана",
        "клеевой слой": "Клейовий шар",
        "наличие рамки": "Наявність рамки",
        "выходной разъем": "Вихідний роз'єм",
        "выходной ток": "Вихідний струм",
        "интерфейс подключения": "Інтерфейс під'єднання",
        "тип кабеля": "Тип кабелю",
        "длина кабеля": "Довжина кабелю",
        "застежка": "Застібка",
        "тип застежки": "Тип застібки",
        "узоры и принты": "Візерунки і принти",
        "особенность цвета": "Особливість кольору",
        "вес в упаковке, кг": "Вага в пакованні, кг",
        "высота в упаковке (см)": "Висота в пакованні (см)",
        "глубина в упаковке (см)": "Глибина в пакованні (см)",
        "назначение": "Призначення",
        "страна-производитель товара": "Країна виробник",
        "страна производитель": "Країна виробник",
        "страна регистрации бренда": "Країна реєстрації бренду",
        "совместимость с беспроводной зарядкой": "Сумісність з бездротовою зарядкою"
    }

    val_exact = {
        # Англійські кольори з файлу
        "matt black": "Чорний", "black": "Чорний", "black + transparent": "Чорний",
        "transparent + black": "Чорний", "camo black": "Чорний", "kevlar black": "Чорний",
        "smoke black": "Чорний", "black geo": "Чорний",
        "ash": "Сірий", "ash grey": "Сірий", "ash/titanium": "Сірий", "carbon fiber": "Сірий",
        "carbon": "Сірий", "titanium": "Сірий", "slate": "Сірий",
        "olive": "Зелений", "olive drab": "Оливковий", "olive green": "Зелений", "khaki": "Хакі",
        "dark orange": "Помаранчевий", "sunset": "Помаранчевий", "brown": "Коричневий",
        "clear": "Прозорий", "transparent": "Прозорий", "pink + transparent": "Рожевий",
        "crimson": "Червоний", "magma": "Червоний", "mallard": "Синій", "cobalt": "Синій",
        "spearmint": "М'ятний", "silver": "Сріблястий", "lilac": "Світло-фіолетовий",
        "white": "Білий", "red": "Червоний", "blue": "Синій", "green": "Зелений", "pink": "Рожевий",
        "grey": "Сірий", "gray": "Сірий",
        # Російські кольори
        "черный": "Чорний", "белый": "Білий", "красный": "Червоний", "синий": "Синій",
        "зеленый": "Зелений", "розовый": "Рожевий", "серый": "Сірий", "светло фиолетовый": "Світло-фіолетовий",
        "серебристый": "Сріблястий", "оливковый": "Оливковий", "фиолетовый": "Фіолетовий",
        "оранжевый": "Помаранчевий", "мятный": "М'ятний", "голубой": "Блакитний",
        # Інше (точні збіги)
        "да": "Так", "нет": "Ні",
        "панель": "Панель (Накладка на корпус)",
        "накладка": "Панель (Накладка на корпус)",
        "чехол-книжка": "Чохол-книжка",
        "бампер": "Бампер",
        "защитное стекло": "Захисне скло",
        "защитная пленка": "Захисна плівка",
        "глянцевое": "Глянсове",
        "по всей поверхности": "По всій поверхні",
        "без узоров и принтов": "Без візерунків і принтів",
        "без застежки": "Без застібки",
        "для телефона": "Для телефону",
        "китай": "Китай", "вьетнам": "В'єтнам", "южная корея": "Південна Корея"
    }

    val_replace = {
        "Силикон": "Силікон", "силикон": "силікон",
        "Поликарбонат": "Полікарбонат", "поликарбонат": "полікарбонат",
        "Термополиуретан": "ТПУ (термопластичний поліуретан)", "термополиуретан": "ТПУ (термопластичний поліуретан)",
        "Арамид": "Арамід", "арамид": "арамід",
        "Кевлар": "Арамід", "кевлар": "арамід",
        "Алюминий": "Метал", "алюминий": "метал",
        "Металл": "Метал", "металл": "метал",
        "Пластик": "Пластик", "пластик": "пластик",
        "Стекло": "Скло", "стекло": "скло",
        "Искусственная кожа": "Штучна шкіра", "искусственная кожа": "штучна шкіра",
        "Кожа": "Штучна шкіра", "кожа": "штучна шкіра",
        
        "Противоударный": "Протиударний", "противоударный": "протиударний",
        "С кольцом": "З кільцем", "с кольцом": "з кільцем",
        "С подставкой": "З підставкою", "с подставкой": "з підставкою",
        "Матовый": "Матовий", "матовый": "матовий",
        "Глянцевый": "Глянсовий", "глянцевый": "глянсовий",
        "Прозрачный": "Прозорий", "прозрачный": "прозорий",
        "Для телефона": "Для телефону", "для телефона": "для телефону",
        "Без застежки": "Без застібки", "без застежки": "без застібки",
        "Без узоров и принтов": "Без візерунків і принтів", "без узоров и принтов": "без візерунків і принтів",
        "Для занятий спортом": "Для занять спортом", "для занятий спортом": "для занять спортом",
        "Защита всего корпуса": "Захист всього корпусу", "защита всего корпуса": "захист всього корпусу",
        "Поддержка бесконтактных платежей": "Підтримка безконтактних платежів", "поддержка бесконтактных платежей": "підтримка безконтактних платежів",
        "Возможность использовать беспроводное ЗУ": "Можливість використовувати бездротовий ЗП", "возможность использовать беспроводное зу": "можливість використовувати бездротовий ЗП",
        "Возможность использования магнитных держателей": "Можливість використання магнітних тримачів", "возможность использования магнитных держателей": "можливість використання магнітних тримачів"
    }

    for item in new_items:
        # Використовуємо .iter() для обходу будь-яких тегів та просторів імен
        for child in item["el"].iter():
            tag_name = child.tag.split('}')[-1]
            
            # --- ЛОГІКА ДЛЯ ФОРМАТУ YML (Prom.ua) ---
            if tag_name == 'param':
                # Видалення сміття від Prom
                if 'paramid' in child.attrib:
                    del child.attrib['paramid']
                if 'valueid' in child.attrib:
                    del child.attrib['valueid']
                
                # Переклад назви (атрибут name)
                old_name = child.get("name")
                if old_name:
                    clean_name = old_name.strip().lower()
                    if clean_name in param_mapping:
                        child.set("name", param_mapping[clean_name])
                
                # Переклад значення (вміст тегу)
                if child.text:
                    val = child.text.strip()
                    if val.lower() in val_exact:
                        child.text = val_exact[val.lower()]
                    else:
                        new_val = child.text
                        for rus, ukr in val_replace.items():
                            new_val = new_val.replace(rus, ukr)
                        child.text = new_val

            # --- ЛОГІКА ДЛЯ ФОРМАТУ GOOGLE MERCHANT CENTER (GMC) ---
            elif tag_name == 'attribute_name':
                if child.text:
                    clean_name = child.text.strip().lower()
                    if clean_name in param_mapping:
                        child.text = param_mapping[clean_name]
                        
            elif tag_name == 'attribute_value':
                if child.text:
                    val = child.text.strip()
                    if val.lower() in val_exact:
                        child.text = val_exact[val.lower()]
                    else:
                        new_val = child.text
                        for rus, ukr in val_replace.items():
                            new_val = new_val.replace(rus, ukr)
                        child.text = new_val
                        
            # Переклад системних кольорів у форматі GMC (<g:color>)
            elif tag_name == 'color':
                if child.text:
                    val = child.text.strip()
                    if val.lower() in val_exact:
                        child.text = val_exact[val.lower()]
                    else:
                        new_val = child.text
                        for rus, ukr in val_replace.items():
                            new_val = new_val.replace(rus, ukr)
                        child.text = new_val
    # =========================================================================

    if fmt == "gmc":
        rss = ET.Element("rss", {
            "xmlns:g": "http://base.google.com/ns/1.0", "version": "2.0"
        })
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
        out_root = ET.Element("yml_catalog",
                              {"date": orig_root.get("date", "")})
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

        # Стан
        self.feed_path  = tk.StringVar(value="")
        self.prom_token = tk.StringVar(value="")
        self.out_path   = tk.StringVar(value="")
        self._running   = False

        self._build_ui()
        self._load_settings()

    # ── ПОБУДОВА ІНТЕРФЕЙСУ ──────────────────

    def _build_ui(self):
        # ── ЗАГОЛОВОК
        hdr = tk.Frame(self, bg=BG, pady=18)
        hdr.pack(fill="x", padx=24)

        tk.Label(hdr, text="⬡", font=("Segoe UI", 20), fg=ACCENT, bg=BG).pack(side="left")
        tk.Label(hdr, text="  PROFETCH", font=FONT_TITLE,
                 fg=TEXT, bg=BG).pack(side="left")
        tk.Label(hdr, text="  Prom Filter", font=("Segoe UI", 10),
                 fg=TEXT_DIM, bg=BG).pack(side="left", pady=3)

        self._sep()

        # ── БЛОК 1: XML ФІДУ
        self._section("1  XML-фід (ваш каталог)")

        feed_row = tk.Frame(self, bg=BG2, bd=0)
        feed_row.pack(fill="x", padx=24, pady=(0, 4))

        self.feed_entry = self._entry(feed_row, self.feed_path, width=52)
        self.feed_entry.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self._btn_small(feed_row, "Огляд…", self._pick_feed).pack(side="left")

        tk.Label(self, text
