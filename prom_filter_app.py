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

        tk.Label(hdr, text="⬡", font=("Segoe UI",
