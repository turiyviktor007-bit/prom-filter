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

# ... (Секції з кольорами та функціями get_prom_products, detect_feed_format, parse_feed залишаються без змін) ...

# Оновлена функція з вашою "бібліотекою"
def build_output_xml(new_items, orig_root, fmt):
    """Генерує XML з використанням точного словника характеристик з вашого Export-файлу."""
    
    # БІБЛІОТЕКА ХАРАКТЕРИСТИК (З вашого CSV експорту)
    # Формат: "Назва з XML": ["Назва для Прому", "Значення для Прому (якщо треба)"]
    library = {
        "Материал": "Матеріал",
        "Цвет": "Колір",
        "Состояние": "Стан",
        "Страна производитель": "Країна виробник",
        "Форм-фактор": "Форм-фактор",
        "Назначение": "Призначення",
        "Совместимая модель": "Модель телефону",
        "Совместимость с Apple iPhone": "Сумісність з Apple iPhone",
        "Особенности": "Особливості",
        "Вид": "Тип",
        "Тип застежки": "Тип застібки"
    }

    # Словник значень (щоб не було "помилок" у системі)
    val_map = {
        "Силикон": "Силікон",
        "Поликарбонат": "Полікарбонат",
        "Термополиуретан": "ТПУ (термопластичний поліуретан)",
        "Панель": "Панель (Накладка на корпус)",
        "Противоударный": "Протиударний"
    }

    for item in new_items:
        for child in item["el"].iter():
            tag_name = child.tag.split('}')[-1]
            if tag_name in ['param', 'attribute_name']:
                name = child.get("name") or child.text
                if name in library:
                    if tag_name == 'param':
                        child.set("name", library[name])
                        # Переклад значення, якщо є в val_map
                        if child.text in val_map:
                            child.text = val_map[child.text]
                    else:
                        child.text = library[name]

    # ... (Решта функції збереження XML залишається як була) ...
    raw    = ET.tostring(orig_root, encoding="unicode") # Спрощено для прикладу
    parsed = minidom.parseString(raw)
    return parsed.toprettyxml(indent="  ", encoding="utf-8")

# ... (Інша частина коду App та main без змін) ...
