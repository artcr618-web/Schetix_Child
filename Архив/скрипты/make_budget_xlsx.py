#!/usr/bin/env python3
# -*- coding: utf-8 -*-
from openpyxl import Workbook
from openpyxl.styles import Font, Fill, PatternFill, Alignment, Border, Side, NamedStyle, Protection
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart.label import DataLabelList
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.chart.series import DataPoint
from openpyxl.drawing.fill import PatternFillProperties, ColorChoice
from copy import copy

wb = Workbook()

# colors
NAVY = "1E3A5F"
NAVY2 = "2A5280"
GOLD = "C9A227"
PAPER = "F7F4EE"
ZEBRA = "FAF8F3"
GREEN = "E8F3EA"
INPUT_BG = "FFF8D6"
ORANGE = "F4E6C8"
WHITE = "FFFFFF"
MUTED = "5B6573"

thin = Border(
    left=Side(style="thin", color="D9D2C4"),
    right=Side(style="thin", color="D9D2C4"),
    top=Side(style="thin", color="D9D2C4"),
    bottom=Side(style="thin", color="D9D2C4"),
)
fill_navy = PatternFill("solid", fgColor=NAVY)
fill_head2 = PatternFill("solid", fgColor=NAVY2)
fill_paper = PatternFill("solid", fgColor=PAPER)
fill_input = PatternFill("solid", fgColor=INPUT_BG)
fill_green = PatternFill("solid", fgColor=GREEN)
fill_zebra = PatternFill("solid", fgColor=ZEBRA)
fill_gold = PatternFill("solid", fgColor="F3E7B8")
fill_white = PatternFill("solid", fgColor=WHITE)
fill_sum = PatternFill("solid", fgColor="1E3A5F")

font_h = Font(name="Calibri", bold=True, color=WHITE, size=11)
font_title = Font(name="Calibri", bold=True, color=NAVY, size=16)
font_sub = Font(name="Calibri", color=MUTED, size=10)
font_b = Font(name="Calibri", bold=True, size=11)
font_n = Font(name="Calibri", size=10)
font_sum = Font(name="Calibri", bold=True, color=WHITE, size=11)
font_hint = Font(name="Calibri", italic=True, color=MUTED, size=9)

center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
right = Alignment(horizontal="right", vertical="center")

money_fmt = '#,##0" ₽"'
num_fmt = "0"

def style_header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row, c)
        cell.fill = fill_navy
        cell.font = font_h
        cell.alignment = center
        cell.border = thin
    ws.row_dimensions[row].height = 32
    ws.auto_filter.ref = f"A{row}:{get_column_letter(cols)}{row}"
    ws.freeze_panes = f"A{row+1}"

def style_input(cell):
    cell.fill = fill_input
    cell.border = thin
    cell.font = font_n
    cell.alignment = center

def style_out(cell, money=False):
    cell.border = thin
    cell.font = font_n
    cell.alignment = right if money else center
    if money:
        cell.number_format = money_fmt

def widths(ws, mapping):
    for col, w in mapping.items():
        ws.column_dimensions[col].width = w

# ===================== DATA =====================
# Things: (sheet, name, qty, shop, url, price, term, unit)
THINGS = [
    ("Школа", "Рубашка школьная", "2 шт.", "Chessford, Детский мир", "https://www.detmir.ru/product/index/id/7065128/", 1398, 1, "год"),
    ("Школа", "Брюки школьные", "2 шт.", "Futurino School", "https://www.detmir.ru/product/index/id/7069209/", 2598, 1, "год"),
    ("Школа", "Жилет школьный", "1 шт.", "Chessford", "https://www.detmir.ru/product/index/id/7065103/", 999, 2, "год"),
    ("Школа", "Джемпер школьный", "1 шт.", "Chessford", "https://www.detmir.ru/product/index/id/7065088/", 1399, 2, "год"),
    ("Школа", "Спортивная форма", "1 компл.", "Futurino School", "https://www.detmir.ru/product/index/id/7069266/", 1999, 2, "год"),
    ("Школа", "Рюкзак школьный", "1 шт.", "Ранец Erhaft", "https://www.detmir.ru/product/index/id/6391235/", 3499, 3, "год"),
    ("Школа", "Сумка для обуви", "1 шт.", "Мешок Erhaft", "https://www.detmir.ru/product/index/id/6622893/", 269, 2, "год"),
    ("Школа", "Сумка для формы", "1 шт.", "Мешок Erhaft Sport", "https://www.detmir.ru/product/index/id/6037286/", 319, 2, "год"),
    ("Школа", "Пенал", "1 шт.", "Erhaft", "https://www.detmir.ru/product/index/id/6400444/", 479, 1, "год"),
    ("Школа", "Дневник", "1 шт.", "Феникс+", "https://www.detmir.ru/product/index/id/6462321/", 329, 1, "год"),
    ("Школа", "Тетради в клетку 12 л.", "20 шт.", "Listoff, 13,90 ₽", "https://www.detmir.ru/product/index/id/6683470/", 278, 1, "год"),
    ("Школа", "Тетради в линейку 12 л.", "20 шт.", "Listoff, 13,90 ₽", "https://www.detmir.ru/product/index/id/6683471/", 278, 1, "год"),
    ("Школа", "Обложки на все тетради", "40 шт.", "Юнландия, 2 набора", "https://www.detmir.ru/product/index/id/6724733/", 358, 1, "год"),
    ("Школа", "Папка для тетрадей", "1 шт.", "А5 на молнии", "https://www.detmir.ru/search/results/?qt=%D0%BF%D0%B0%D0%BF%D0%BA%D0%B0%20%D0%B4%D0%BB%D1%8F%20%D1%82%D0%B5%D1%82%D1%80%D0%B0%D0%B4%D0%B5%D0%B9", 199, 1, "год"),
    ("Школа", "Ручки синие", "5 шт.", "Erhaft", "https://www.detmir.ru/product/index/id/6505850/", 249, 1, "год"),
    ("Школа", "Карандаши простые", "5 шт.", "HB", "https://www.detmir.ru/search/results/?qt=%D0%BA%D0%B0%D1%80%D0%B0%D0%BD%D0%B4%D0%B0%D1%88%20HB", 99, 1, "год"),
    ("Школа", "Карандаши цветные 32/36", "1 набор", "Brauberg 36", "https://www.detmir.ru/product/index/id/6016009/", 328, 1, "год"),
    ("Школа", "Фломастеры 32/50", "1 набор", "Erhaft 50", "https://www.detmir.ru/product/index/id/3293711/", 549, 1, "год"),
    ("Школа", "Ластик", "1 шт.", "Мягкий", "https://www.detmir.ru/search/results/?qt=%D0%BB%D0%B0%D1%81%D1%82%D0%B8%D0%BA", 39, 1, "год"),
    ("Школа", "Точилка", "1 шт.", "С контейнером", "https://www.detmir.ru/search/results/?qt=%D1%82%D0%BE%D1%87%D0%B8%D0%BB%D0%BA%D0%B0", 79, 1, "год"),
    ("Школа", "Линейка", "1 шт.", "20 см", "https://www.detmir.ru/search/results/?qt=%D0%BB%D0%B8%D0%BD%D0%B5%D0%B9%D0%BA%D0%B0%2020", 39, 1, "год"),
    ("Школа", "Ножницы", "1 шт.", "Детские", "https://www.detmir.ru/search/results/?qt=%D0%BD%D0%BE%D0%B6%D0%BD%D0%B8%D1%86%D1%8B%20%D0%B4%D0%B5%D1%82%D1%81%D0%BA%D0%B8%D0%B5", 129, 1, "год"),
    ("Школа", "Клей", "1 шт.", "Клей-карандаш", "https://www.detmir.ru/search/results/?qt=%D0%BA%D0%BB%D0%B5%D0%B9-%D0%BA%D0%B0%D1%80%D0%B0%D0%BD%D0%B4%D0%B0%D1%88", 59, 1, "год"),
    ("Школа", "Цветная бумага", "1 набор", "А4", "https://www.detmir.ru/search/results/?qt=%D1%86%D0%B2%D0%B5%D1%82%D0%BD%D0%B0%D1%8F%20%D0%B1%D1%83%D0%BC%D0%B0%D0%B3%D0%B0", 69, 1, "год"),
    ("Школа", "Альбом для рисования", "1 шт.", "А4", "https://www.detmir.ru/search/results/?qt=%D0%B0%D0%BB%D1%8C%D0%B1%D0%BE%D0%BC", 89, 1, "год"),
    ("Школа", "Картон цветной", "1 набор", "А4", "https://www.detmir.ru/search/results/?qt=%D0%BA%D0%B0%D1%80%D1%82%D0%BE%D0%BD%20%D1%86%D0%B2%D0%B5%D1%82%D0%BD%D0%BE%D0%B9", 89, 1, "год"),
    ("Школа", "Гуашь", "1 набор", "Erhaft 12 цв.", "https://www.detmir.ru/product/index/id/6430833/", 239, 1, "год"),
    ("Школа", "Кисти", "1 набор", "3 шт.", "https://www.detmir.ru/search/results/?qt=%D0%BA%D0%B8%D1%81%D1%82%D0%B8", 99, 1, "год"),
    ("Школа", "Непроливайка", "1 шт.", "Стакан", "https://www.detmir.ru/search/results/?qt=%D0%BD%D0%B5%D0%BF%D1%80%D0%BE%D0%BB%D0%B8%D0%B2%D0%B0%D0%B9%D0%BA%D0%B0", 79, 1, "год"),
    ("Школа", "Пластилин", "1 набор", "Erhaft 12 цв.", "https://www.detmir.ru/search/results/?qt=%D0%BF%D0%BB%D0%B0%D1%81%D1%82%D0%B8%D0%BB%D0%B8%D0%BD", 149, 1, "год"),
    ("Школа", "Доска для пластилина", "1 шт.", "Для лепки", "https://www.detmir.ru/search/results/?qt=%D0%B4%D0%BE%D1%81%D0%BA%D0%B0%20%D0%B4%D0%BB%D1%8F%20%D0%BB%D0%B5%D0%BF%D0%BA%D0%B8", 49, 1, "год"),
    ("Школа", "Готовальня", "1 шт.", "Brauberg", "https://www.detmir.ru/catalog/index/name/cyrkuli/tip_cherchenie-gotovalnya/", 349, 1, "год"),
    ("Школа", "Школьные пособия / рабочие тетради", "на год", "ориентир, цена меняется", "https://www.detmir.ru/search/results/?qt=%D1%80%D0%B0%D0%B1%D0%BE%D1%87%D0%B8%D0%B5%20%D1%82%D0%B5%D1%82%D1%80%D0%B0%D0%B4%D0%B8", 1000, 1, "год"),
    ("Школа", "Папка для труда", "1 шт.", "отдельная, с ручками А4", "https://www.detmir.ru/search/results/?qt=%D0%BF%D0%B0%D0%BF%D0%BA%D0%B0%20%D0%B4%D0%BB%D1%8F%20%D1%82%D1%80%D1%83%D0%B4%D0%B0", 349, 1, "год"),
    ("Школа", "Клей ПВА", "1 тюбик", "каждому ребёнку свой", "https://www.detmir.ru/search/results/?qt=%D0%BA%D0%BB%D0%B5%D0%B9%20%D0%9F%D0%92%D0%90", 79, 1, "год"),
    ("Школа", "Корректор", "1 шт.", "лента или ручка", "https://www.detmir.ru/search/results/?qt=%D0%BA%D0%BE%D1%80%D1%80%D0%B5%D0%BA%D1%82%D0%BE%D1%80%20%D1%88%D0%BA%D0%BE%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9", 59, 1, "год"),
    ("Школа", "Букет учителю", "1 шт.", "1 сентября", "https://www.detmir.ru/search/results/?qt=%D0%B1%D1%83%D0%BA%D0%B5%D1%82", 2000, 1, "год"),
    ("Школа", "Выпускной", "1 раз", "праздник школы", "", 10000, 5, "год"),

    ("Обувь", "Туфли школьные", "1 пара", "Futurino School", "https://www.detmir.ru/product/index/id/7077747/", 999, 1, "год"),
    ("Обувь", "Обувь спортивная", "1 пара", "Кроссовки Jomoto", "https://www.detmir.ru/product/index/id/7076432/", 2499, 1, "год"),
    ("Обувь", "Туфли праздничные", "1 пара", "Futurino School", "https://www.detmir.ru/product/index/id/7077779/", 1599, 1, "год"),
    ("Обувь", "Ботинки зимние", "1 пара", "Jomoto / Futurino", "https://www.detmir.ru/catalog/index/name/botinki_m/sezon_oio-zima/", 3999, 1, "год"),
    ("Обувь", "Ботинки осенние утеплённые", "1 пара", "Детский мир", "https://www.detmir.ru/catalog/index/name/botinki_m/sezon_oio-demisezon/", 2999, 1, "год"),
    ("Обувь", "Тапочки домашние", "1 пара", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D1%82%D0%B0%D0%BF%D0%BE%D1%87%D0%BA%D0%B8", 599, 1, "год"),
    ("Обувь", "Сланцы летние", "1 пара", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D1%81%D0%BB%D0%B0%D0%BD%D1%86%D1%8B", 499, 1, "год"),
    ("Обувь", "Сандалии летние", "1 пара", "Futurino", "https://www.detmir.ru/catalog/index/name/shoes_for_boys/brand/4112/", 999, 1, "год"),

    ("Верхняя одежда", "Куртка зимняя", "2 шт.", "Lassie / Futurino", "https://www.detmir.ru/product/index/id/7061634/", 8584, 2, "год"),
    ("Верхняя одежда", "Куртка осенняя утеплённая", "1 шт.", "Futurino", "https://www.detmir.ru/product/index/id/6653929/", 2199, 2, "год"),
    ("Верхняя одежда", "Ветровка / тонкая осенняя", "1 шт.", "Jomoto / Futurino", "https://www.detmir.ru/product/index/id/6022408/", 1299, 2, "год"),
    ("Верхняя одежда", "Штаны утеплённые", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D1%88%D1%82%D0%B0%D0%BD%D1%8B%20%D1%83%D1%82%D0%B5%D0%BF%D0%BB%D0%B5%D0%BD%D0%BD%D1%8B%D0%B5", 2499, 2, "год"),
    ("Верхняя одежда", "Комбинезон зимний", "1 шт.", "GUSTI", "https://www.detmir.ru/catalog/index/name/komplekty_m/sezon_oio-zima/", 9221, 2, "год"),
    ("Верхняя одежда", "Зонт или дождевик", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%B4%D0%BE%D0%B6%D0%B4%D0%B5%D0%B2%D0%B8%D0%BA%20%D0%B4%D0%B5%D1%82%D1%81%D0%BA%D0%B8%D0%B9", 699, 2, "год"),

    ("Бельё", "Майки", "5 шт.", "Futurino", "https://www.detmir.ru/product/index/id/6733145/", 1995, 1, "год"),
    ("Бельё", "Трусы", "5 шт.", "Futurino набор", "https://www.detmir.ru/product/index/id/7072737/", 799, 1, "год"),
    ("Бельё", "Носки обычные", "5 пар", "Futurino набор", "https://www.detmir.ru/product/index/id/6550049/", 399, 1, "год"),
    ("Бельё", "Носки утеплённые зимние", "2 пары", "Futurino / Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BD%D0%BE%D1%81%D0%BA%D0%B8%20%D0%B7%D0%B8%D0%BC%D0%BD%D0%B8%D0%B5%20%D0%B4%D0%B5%D1%82%D1%81%D0%BA%D0%B8%D0%B5", 798, 1, "год"),
    ("Бельё", "Термобельё", "1 компл.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D1%82%D0%B5%D1%80%D0%BC%D0%BE%D0%B1%D0%B5%D0%BB%D1%8C%D0%B5", 1499, 2, "год"),

    ("Одежда", "Шорты летние", "2 шт.", "если не добавляли — двое", "https://www.detmir.ru/search/results/?qt=%D1%88%D0%BE%D1%80%D1%82%D1%8B%20%D0%B4%D0%B5%D1%82%D1%81%D0%BA%D0%B8%D0%B5", 1598, 1, "год"),
    ("Одежда", "Джинсы", "2 шт.", "Futurino", "https://www.detmir.ru/catalog/index/name/detskaya_odezhda/brand/4112/", 1598, 2, "год"),
    ("Одежда", "Футболки", "2 шт.", "Chessford", "https://www.detmir.ru/product/index/id/7065134/", 598, 1, "год"),
    ("Одежда", "Худи", "2 шт.", "Futurino", "https://www.detmir.ru/search/results/?qt=%D1%85%D1%83%D0%B4%D0%B8", 2598, 2, "год"),
    ("Одежда", "Свитер", "2 шт.", "Futurino", "https://www.detmir.ru/search/results/?qt=%D1%81%D0%B2%D0%B8%D1%82%D0%B5%D1%80", 2398, 2, "год"),
    ("Одежда", "Домашний костюм", "2 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%B4%D0%BE%D0%BC%D0%B0%D1%88%D0%BD%D0%B8%D0%B9%20%D0%BA%D0%BE%D1%81%D1%82%D1%8E%D0%BC", 2998, 2, "год"),
    ("Одежда", "Брюки праздничные", "1 шт.", "Futurino School", "https://www.detmir.ru/product/index/id/7069209/", 1299, 2, "год"),
    ("Одежда", "Рубашка праздничная", "1 шт.", "Chessford", "https://www.detmir.ru/product/index/id/7065128/", 699, 1, "год"),
    ("Одежда", "Пиджак праздничный", "1 шт.", "Futurino School", "https://www.detmir.ru/product/index/id/6643629/", 2399, 2, "год"),

    ("Аксессуары", "Солнечные очки", "1 шт.", "PlayToday, покупали", "https://www.detmir.ru/search/results/?qt=PlayToday%20%D0%BE%D1%87%D0%BA%D0%B8", 799, 2, "год"),
    ("Аксессуары", "Бейсболка", "2 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%B1%D0%B5%D0%B9%D1%81%D0%B1%D0%BE%D0%BB%D0%BA%D0%B0", 798, 2, "год"),
    ("Аксессуары", "Шапка осенняя", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D1%88%D0%B0%D0%BF%D0%BA%D0%B0%20%D0%BE%D1%81%D0%B5%D0%BD%D0%BD%D1%8F%D1%8F", 499, 2, "год"),
    ("Аксессуары", "Шапка зимняя", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D1%88%D0%B0%D0%BF%D0%BA%D0%B0%20%D0%B7%D0%B8%D0%BC%D0%BD%D1%8F%D1%8F", 699, 2, "год"),
    ("Аксессуары", "Перчатки осенние", "1 пара", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BF%D0%B5%D1%80%D1%87%D0%B0%D1%82%D0%BA%D0%B8", 399, 2, "год"),
    ("Аксессуары", "Перчатки зимние", "1 пара", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BF%D0%B5%D1%80%D1%87%D0%B0%D1%82%D0%BA%D0%B8%20%D0%B7%D0%B8%D0%BC%D0%BD%D0%B8%D0%B5", 599, 2, "год"),
    ("Аксессуары", "Шарф", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D1%88%D0%B0%D1%80%D1%84", 399, 2, "год"),
    ("Аксессуары", "Галстук", "1 шт.", "Futurino School", "https://www.detmir.ru/catalog/index/name/neckties/", 299, 2, "год"),

    ("Самбо и спорт", "Борцовка (куртка самбо)", "1 шт.", "Green Hill Junior", "https://rocky-shop.ru/ekipirovka-dlya-edinoborstv/kurtka-dlya-sambo-green-hill-junior-rostovki-160-190-sm-10032.html", 3690, 2, "год"),
    ("Самбо и спорт", "Шорты для самбо", "1 шт.", "Green Hill", "https://rocky-shop.ru/ekipirovka-dlya-edinoborstv/detskie-bortscovki-dlya-sambo-green-hill-triumf-odobreny-fias-27458.html", 1990, 2, "год"),
    ("Самбо и спорт", "Перчатки для самбо", "1 пара", "Green Hill", "https://dantesport.ru/product_info.php?products_id=1207", 2990, 2, "год"),
    ("Самбо и спорт", "Защита голени", "1 пара", "Green Hill", "https://karate.ru/shop/catalog/407-perchatki-mma-sinie-kozhzamenitel-combat-sambo-green-hill/", 2190, 2, "год"),
    ("Самбо и спорт", "Сумка для самбо", "1 шт.", "Спортмагазин", "https://www.detmir.ru/search/results/?qt=%D1%81%D0%BF%D0%BE%D1%80%D1%82%D0%B8%D0%B2%D0%BD%D0%B0%D1%8F%20%D1%81%D1%83%D0%BC%D0%BA%D0%B0", 1499, 3, "год"),
    ("Самбо и спорт", "Скакалка", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D1%81%D0%BA%D0%B0%D0%BA%D0%B0%D0%BB%D0%BA%D0%B0", 399, 3, "год"),
    ("Самбо и спорт", "Плавки", "1 шт.", "Детский мир", "https://www.detmir.ru/catalog/index/name/plavki/", 599, 1, "год"),
    ("Самбо и спорт", "Полотенце для плавания", "1 шт.", "маска уже есть", "https://www.detmir.ru/search/results/?qt=%D0%BF%D0%BE%D0%BB%D0%BE%D1%82%D0%B5%D0%BD%D1%86%D0%B5%20%D0%BF%D0%BB%D0%B0%D0%B2%D0%B0%D0%BD%D0%B8%D0%B5", 499, 1, "год"),
    ("Самбо и спорт", "Велосипед Stern 24\"", "1 шт.", "Stern", "https://market.yandex.ru/search?text=%D0%B2%D0%B5%D0%BB%D0%BE%D1%81%D0%B8%D0%BF%D0%B5%D0%B4%20Stern%2024", 27999, 5, "год"),

    ("Гигиена", "Шейкер / бутылка", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%B1%D1%83%D1%82%D1%8B%D0%BB%D0%BA%D0%B0", 499, 2, "год"),
    ("Гигиена", "Мочалка", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BC%D0%BE%D1%87%D0%B0%D0%BB%D0%BA%D0%B0", 199, 1, "год"),
    ("Гигиена", "Полотенце банное", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BF%D0%BE%D0%BB%D0%BE%D1%82%D0%B5%D0%BD%D1%86%D0%B5%20%D0%B1%D0%B0%D0%BD%D0%BD%D0%BE%D0%B5", 699, 1, "год"),
    ("Гигиена", "Полотенце для лица", "2 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BF%D0%BE%D0%BB%D0%BE%D1%82%D0%B5%D0%BD%D1%86%D0%B5%20%D0%B4%D0%BB%D1%8F%20%D0%BB%D0%B8%D1%86%D0%B0", 698, 1, "год"),

    ("Комната", "Кровать-машинка", "1 шт.", "Авито, ваша покупка", "https://www.avito.ru/", 25000, 7, "год"),
    ("Комната", "Матрас 80×180", "1 шт.", "Орматек Kids Classic", "https://ormatek.com/catalog/matrasy/80x180/", 11950, 8, "год"),
    ("Комната", "Подушка", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BF%D0%BE%D0%B4%D1%83%D1%88%D0%BA%D0%B0", 1299, 3, "год"),
    ("Комната", "Одеяло", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BE%D0%B4%D0%B5%D1%8F%D0%BB%D0%BE", 1999, 3, "год"),
    ("Комната", "Постельное бельё", "2 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BF%D0%BE%D1%81%D1%82%D0%B5%D0%BB%D1%8C%D0%BD%D0%BE%D0%B5", 2998, 2, "год"),
    ("Комната", "Стеллаж для книг", "1 шт.", "Детский мир", "https://www.detmir.ru/catalog/index/name/stellsjy/", 4786, 5, "год"),
    ("Комната", "Стол письменный", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D1%81%D1%82%D0%BE%D0%BB%20%D0%BF%D0%B8%D1%81%D1%8C%D0%BC%D0%B5%D0%BD%D0%BD%D1%8B%D0%B9", 4999, 7, "год"),
    ("Комната", "Кресло компьютерное", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BA%D1%80%D0%B5%D1%81%D0%BB%D0%BE", 4499, 5, "год"),
    ("Комната", "Настольная лампа", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BB%D0%B0%D0%BC%D0%BF%D0%B0", 1299, 5, "год"),
    ("Комната", "Ковёр на пол", "1 шт.", "Леруа Мерлен, комната ребёнка, ориентир", "https://leroymerlin.ru/catalogue/kovry/", 4000, 7, "год"),
    ("Комната", "Светильник потолочный / люстра", "1 шт.", "пока нет, строка чтобы не забыть", "https://leroymerlin.ru/catalogue/lyustry/", 2500, 10, "год"),
    ("Комната", "Шторы", "2 шт. × 1 500 ₽", "Леруа Мерлен", "https://leroymerlin.ru/catalogue/shtory/", 3000, 5, "год"),
    ("Комната", "Вуаль / тюль", "1 шт.", "Леруа Мерлен, к шторам", "https://leroymerlin.ru/catalogue/tyul/", 1200, 5, "год"),

    ("Улица и дача", "Мяч надувной", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BC%D1%8F%D1%87%20%D0%BD%D0%B0%D0%B4%D1%83%D0%B2%D0%BD%D0%BE%D0%B9", 399, 2, "год"),
    ("Улица и дача", "Мяч футбольный", "1 шт.", "Детский мир / спорт", "https://www.detmir.ru/search/results/?qt=%D0%BC%D1%8F%D1%87%20%D1%84%D1%83%D1%82%D0%B1%D0%BE%D0%BB%D1%8C%D0%BD%D1%8B%D0%B9", 799, 2, "год"),
    ("Улица и дача", "Надувной матрас для купания", "1 шт.", "Детский мир", "https://www.detmir.ru/search/results/?qt=%D0%BD%D0%B0%D0%B4%D1%83%D0%B2%D0%BD%D0%BE%D0%B9%20%D0%BC%D0%B0%D1%82%D1%80%D0%B0%D1%81", 999, 3, "год"),
    ("Улица и дача", "Надувной круг", "1 шт.", "свой у каждого ребёнка", "https://www.detmir.ru/search/results/?qt=%D0%BD%D0%B0%D0%B4%D1%83%D0%B2%D0%BD%D0%BE%D0%B9%20%D0%BA%D1%80%D1%83%D0%B3", 599, 2, "год"),
    ("Улица и дача", "Стул для кемпинга", "1 шт.", "свой у каждого ребёнка", "https://leroymerlin.ru/search/?q=%D1%81%D1%82%D1%83%D0%BB%20%D0%BA%D0%B5%D0%BC%D0%BF%D0%B8%D0%BD%D0%B3", 999, 5, "год"),

    ("Техника", "Системный блок", "1 шт.", "Авито", "https://www.avito.ru/", 23000, 4, "год"),
    ("Техника", "Клавиатура", "1 шт.", "DNS / Авито", "https://www.dns-shop.ru/catalog/17a8a40216404e77/klaviatury/", 1499, 3, "год"),
    ("Техника", "Мышь", "1 шт.", "DNS / Авито", "https://www.dns-shop.ru/catalog/17a8a4d816404e77/myshi/", 799, 2, "год"),
    ("Техника", "Монитор", "1 шт.", "DNS / Авито", "https://www.dns-shop.ru/catalog/17a8943716404e77/monitory/", 8999, 5, "год"),
    ("Техника", "Коврик для мыши", "1 шт.", "DNS", "https://www.dns-shop.ru/catalog/17a9f9e316404e77/kovriki-dlya-myshi/", 299, 2, "год"),
    ("Техника", "Колонки", "1 шт.", "DNS", "https://www.dns-shop.ru/catalog/17a8914916404e77/kompyuternaya-akustika/", 1499, 4, "год"),
    ("Техника", "Микрофон", "1 шт.", "DNS", "https://www.dns-shop.ru/catalog/17a89bdd16404e77/mikrofoni/", 1299, 4, "год"),
    ("Техника", "Наушники", "1 шт.", "DNS", "https://www.detmir.ru/search/results/?qt=%D0%BD%D0%B0%D1%83%D1%88%D0%BD%D0%B8%D0%BA%D0%B8", 1999, 3, "год"),
    ("Техника", "Смартфон", "1 шт.", "с рук до 5 000", "https://www.avito.ru/", 5000, 1, "год"),
]

THING_SHEETS = [
    "Школа", "Обувь", "Верхняя одежда", "Бельё", "Одежда",
    "Аксессуары", "Самбо и спорт", "Улица и дача", "Гигиена", "Комната", "Техника",
]

# ===================== HOW TO =====================
ws = wb.active
ws.title = "Как пользоваться"
ws.sheet_properties.tabColor = GOLD
ws["A1"] = "Как пользоваться этой книгой"
ws["A1"].font = font_title
ws.merge_cells("A1:B1")
ws["A2"] = "Жёлтые ячейки можно менять. Серые с формулами не трогайте — они сами пересчитаются."
ws["A2"].font = font_sub
ws.merge_cells("A2:B2")

howto = [
    ("Листы «Школа» … «Техника»", "Вещи. Меняйте цену (F), срок (G) и единицу (H: год или месяц)."),
    ("Срок + единица", "2 и «год» → в год = цена÷2, в месяц = цена÷24.  6 и «месяц» → в месяц = цена÷6."),
    ("Лист «Регулярные ребёнок»", "То, что целиком на этого ребёнка: секции, карманные, день рождения."),
    ("Лист «Общие ÷3»", "Квартира, коммуналка по статьям, стирка, озеро. Семья целиком и доля ребёнка."),
    ("Лист «Выходные»", "Кино и зоопарк: свой билет + половина маминого (двое детей)."),
    ("День рождения", "Подарок 5 000 + стол 5 000 = 10 000, всё на ребёнка (праздник под него)."),
    ("Новый год", "Подарок ребёнку 5 000 — на него. Стол — общий праздник, делим на 3."),
    ("Коммуналка", "7 000 ₽ разобраны по строкам. Воду и отопление добавила — их часто забывают. Поправьте под свою квитанцию."),
    ("Озеро", "5 000 ₽ поездка, 5 раз в месяц, 3 тёплых месяца. Семья 75 000 / год, ребёнок ÷3."),
    ("Сводка", "Всё складывается формулами с других листов. Если добавите строку — расширьте диапазон SUM."),
]
ws["A4"] = "Что"
ws["B4"] = "Зачем"
style_header(ws, 4, 2)
for i, (a, b) in enumerate(howto, 5):
    ws.cell(i, 1, a).font = font_b
    ws.cell(i, 2, b).font = font_n
    ws.cell(i, 1).alignment = left
    ws.cell(i, 2).alignment = left
    ws.cell(i, 1).border = thin
    ws.cell(i, 2).border = thin
    if i % 2 == 0:
        ws.cell(i, 1).fill = fill_zebra
        ws.cell(i, 2).fill = fill_zebra
widths(ws, {"A": 34, "B": 100})
ws.row_dimensions[1].height = 24
ws.freeze_panes = "A5"
ws.print_title_rows = "1:4"
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True

# ===================== THING SHEETS =====================
dv_unit = DataValidation(type="list", formula1='"год,месяц"', allow_blank=False)
dv_unit.error = "Выберите: год или месяц"
dv_unit.errorTitle = "Единица срока"
dv_unit.prompt = "год или месяц"
dv_unit.promptTitle = "Единица"

def write_things(sheet_name, rows):
    ws = wb.create_sheet(sheet_name)
    ws.sheet_properties.tabColor = NAVY
    ws["A1"] = sheet_name
    ws["A1"].font = font_title
    ws.merge_cells("A1:J1")
    ws["A2"] = "Жёлтое — ваши цифры. Срок = число, единица = год или месяц. Амортизация в колонках I и J считается сама."
    ws["A2"].font = font_sub
    ws.merge_cells("A2:J2")
    headers = ["№", "Что", "Сколько", "Где / что брать", "Ссылка", "Цена, ₽", "Срок", "Единица", "В год, ₽", "В месяц, ₽"]
    for c, h in enumerate(headers, 1):
        ws.cell(4, c, h)
    style_header(ws, 4, 10)
    ws.auto_filter.ref = f"A4:J{4+len(rows)}"

    for i, row in enumerate(rows, 1):
        r = 4 + i
        _, name, qty, shop, url, price, term, unit = row
        ws.cell(r, 1, i).alignment = center
        ws.cell(r, 1).font = font_n
        ws.cell(r, 1).border = thin
        ws.cell(r, 2, name).alignment = left
        ws.cell(r, 2).font = font_n
        ws.cell(r, 2).border = thin
        ws.cell(r, 3, qty).alignment = center
        ws.cell(r, 3).font = font_n
        ws.cell(r, 3).border = thin
        ws.cell(r, 4, shop).alignment = left
        ws.cell(r, 4).font = font_n
        ws.cell(r, 4).border = thin
        cell_l = ws.cell(r, 5, "открыть" if url else "—")
        cell_l.alignment = center
        cell_l.border = thin
        if url:
            cell_l.hyperlink = url
            cell_l.font = Font(name="Calibri", size=10, color="0563C1", underline="single")
        else:
            cell_l.font = font_n
        c_price = ws.cell(r, 6, price)
        c_price.number_format = money_fmt
        style_input(c_price)
        c_term = ws.cell(r, 7, term)
        style_input(c_term)
        c_unit = ws.cell(r, 8, unit)
        style_input(c_unit)
        # year
        f_y = f'=IF(H{r}="месяц",IF(G{r}=0,0,F{r}*12/G{r}),IF(G{r}=0,0,F{r}/G{r}))'
        c_y = ws.cell(r, 9, f_y)
        c_y.number_format = money_fmt
        style_out(c_y, True)
        f_m = f'=IF(H{r}="месяц",IF(G{r}=0,0,F{r}/G{r}),IF(G{r}=0,0,F{r}/(G{r}*12)))'
        c_m = ws.cell(r, 10, f_m)
        c_m.number_format = money_fmt
        style_out(c_m, True)
        if i % 2 == 0:
            for c in (1, 2, 3, 4, 5):
                if ws.cell(r, c).fill.fgColor is None or ws.cell(r, c).fill.fgColor.rgb in (None, "00000000"):
                    ws.cell(r, c).fill = fill_zebra
        ws.row_dimensions[r].height = 18

    last = 4 + len(rows)
    tot = last + 1
    ws.cell(tot, 1, "").fill = fill_sum
    ws.merge_cells(start_row=tot, start_column=1, end_row=tot, end_column=5)
    ws.cell(tot, 1, f"Итого: {sheet_name}")
    ws.cell(tot, 1).font = font_sum
    ws.cell(tot, 1).fill = fill_sum
    ws.cell(tot, 1).alignment = left
    for c in range(2, 6):
        ws.cell(tot, c).fill = fill_sum
        ws.cell(tot, c).font = font_sum
        ws.cell(tot, c).border = thin
    ws.cell(tot, 6, f"=SUM(F5:F{last})")
    ws.cell(tot, 6).number_format = money_fmt
    ws.cell(tot, 9, f"=SUM(I5:I{last})")
    ws.cell(tot, 9).number_format = money_fmt
    ws.cell(tot, 10, f"=SUM(J5:J{last})")
    ws.cell(tot, 10).number_format = money_fmt
    for c in (6, 7, 8, 9, 10):
        ws.cell(tot, c).fill = fill_sum
        ws.cell(tot, c).font = font_sum
        ws.cell(tot, c).border = thin
        ws.cell(tot, c).alignment = right
    ws.row_dimensions[tot].height = 22

    dv = DataValidation(type="list", formula1='"год,месяц"', allow_blank=False)
    dv.add(f"H5:H{last}")
    ws.add_data_validation(dv)

    widths(ws, {"A": 5, "B": 38, "C": 12, "D": 28, "E": 12, "F": 14, "G": 10, "H": 12, "I": 14, "J": 14})
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_title_rows = "1:4"
    ws.page_setup.horizontalCentered = True
    ws.oddHeader.left.text = sheet_name
    ws.oddFooter.right.text = "Страница &P из &N"
    return last, tot

sheet_ranges = {}
for sh in THING_SHEETS:
    rows = [x for x in THINGS if x[0] == sh]
    last, tot = write_things(sh, rows)
    sheet_ranges[sh] = (last, tot)

# ===================== REGULAR CHILD =====================
# name, note, amount, times, period (месяц/год), months_active, every_n, 
# For simple: monthly_amount OR yearly_amount via period
# Columns: № Что Пояснение Сумма_платежа Сколько_раз Период Месяцев_в_году Каждые_N  В_год  В_месяц
REC_CHILD = [
    ("Називин, 1 флакон", "аптека ~340 ₽, 1 раз в год", 340, 1, "год", 12, 1),
    ("Долфин", "2 упаковки в год × 548 ₽", 548, 2, "год", 12, 1),
    ("Шампунь детский", "1 флакон на 2 месяца ~199 ₽", 199, 1, "месяц", 12, 2),
    ("Зубная щётка", "раз в 3 месяца ~149 ₽", 149, 1, "месяц", 12, 3),
    ("Зубная паста", "тюбик в месяц ~179 ₽", 179, 1, "месяц", 12, 1),
    ("Влажные салфетки маленькие", "4 уп. в месяц × 50 ₽", 50, 4, "месяц", 12, 1),
    ("Питание в школе", "100 ₽ в день × 365", 100, 365, "год", 12, 1),
    ("Карманные расходы", "200 ₽ в день × 365", 200, 365, "год", 12, 1),
    ("Школьные сборы", "3 000 ₽ в год", 3000, 1, "год", 12, 1),
    ("Поездки и мероприятия школы", "500 ₽ в месяц", 500, 1, "месяц", 12, 1),
    ("Секция самбо", "4 000 ₽ в месяц", 4000, 1, "месяц", 12, 1),
    ("Секция плавания", "ориентир Иваново 3 500 ₽/мес", 3500, 1, "месяц", 12, 1),
    ("Оплата мобильного", "300 ₽ в месяц, этот ребёнок", 300, 1, "месяц", 12, 1),
    ("День рождения: подарок", "меньше 5 000 не получается", 5000, 1, "год", 12, 1),
    ("День рождения: стол", "собираем чисто под ребёнка", 5000, 1, "год", 12, 1),
    ("Новый год: подарок ребёнку", "подарок ему, стол — на листе общих", 5000, 1, "год", 12, 1),
    ("Подарок на 23 февраля", "раз в год", 2000, 1, "год", 12, 1),
    ("Хотелочки", "игрушка / мелочь, 2 000 ₽ в месяц", 2000, 1, "месяц", 12, 1),
    ("Доп. курс русский в школе", "платный, 300 ₽ в месяц", 300, 1, "месяц", 12, 1),
]

ws = wb.create_sheet("Регулярные ребёнок")
ws.sheet_properties.tabColor = "2A7D4F"
ws["A1"] = "Регулярные траты — только этот ребёнок (100%)"
ws["A1"].font = font_title
ws.merge_cells("A1:J1")
ws["A2"] = "День рождения: подарок 5 000 + стол 5 000 = 10 000, всё на него. Жёлтое меняйте. Период = месяц или год. «Каждые N» = шампунь раз в 2 месяца поставьте N=2."
ws["A2"].font = font_sub
ws.merge_cells("A2:J2")
headers = ["№", "Что", "Пояснение", "Сумма платежа, ₽", "Сколько раз", "Период", "Месяцев в году", "Каждые N периодов", "В год, ₽", "В месяц, ₽"]
for c, h in enumerate(headers, 1):
    ws.cell(4, c, h)
style_header(ws, 4, 10)

for i, row in enumerate(REC_CHILD, 1):
    r = 4 + i
    name, note, amount, times, period, months, every = row
    ws.cell(r, 1, i).alignment = center
    ws.cell(r, 1).border = thin
    ws.cell(r, 1).font = font_n
    ws.cell(r, 2, name).alignment = left
    ws.cell(r, 2).border = thin
    ws.cell(r, 2).font = font_n
    ws.cell(r, 3, note).alignment = left
    ws.cell(r, 3).border = thin
    ws.cell(r, 3).font = font_n
    c = ws.cell(r, 4, amount)
    c.number_format = money_fmt
    style_input(c)
    style_input(ws.cell(r, 5, times))
    style_input(ws.cell(r, 6, period))
    style_input(ws.cell(r, 7, months))
    style_input(ws.cell(r, 8, every))
    # year = if period year: amount*times; if month: amount*times*(months/every)
    fy = f'=IF(F{r}="год",D{r}*E{r},D{r}*E{r}*(G{r}/IF(H{r}=0,1,H{r})))'
    cy = ws.cell(r, 9, fy)
    cy.number_format = money_fmt
    style_out(cy, True)
    cm = ws.cell(r, 10, f"=IF(I{r}=0,0,I{r}/12)")
    cm.number_format = money_fmt
    style_out(cm, True)
    ws.row_dimensions[r].height = 20

last = 4 + len(REC_CHILD)
tot = last + 1
ws.cell(tot, 1, "Итого на ребёнка")
ws.merge_cells(start_row=tot, start_column=1, end_row=tot, end_column=8)
for c in range(1, 11):
    ws.cell(tot, c).fill = fill_sum
    ws.cell(tot, c).font = font_sum
    ws.cell(tot, c).border = thin
ws.cell(tot, 9, f"=SUM(I5:I{last})")
ws.cell(tot, 9).number_format = money_fmt
ws.cell(tot, 10, f"=SUM(J5:J{last})")
ws.cell(tot, 10).number_format = money_fmt
ws.cell(tot, 9).font = font_sum
ws.cell(tot, 10).font = font_sum

dv = DataValidation(type="list", formula1='"месяц,год"', allow_blank=False)
dv.add(f"F5:F{last}")
ws.add_data_validation(dv)
ws.auto_filter.ref = f"A4:J{last}"
widths(ws, {"A": 5, "B": 36, "C": 42, "D": 16, "E": 12, "F": 12, "G": 16, "H": 18, "I": 14, "J": 14})
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "1:4"
rec_child_tot = tot

# ===================== FOOD CHILD 100% =====================
# name, note, amount (цена единицы), times, period, months, every
PROD_CHILD = [
    ("Молоко Эго, Магнит", "3 бутылки в неделю, этот ребёнок; ориентир 99 ₽/бут.", 99, 3, "неделя", 12, 1),
    ("Пельмени, Магнит", "2 пачки в неделю, этот ребёнок; ориентир 250 ₽", 250, 2, "неделя", 12, 1),
    ("Шоколад Milka", "3 плитки в неделю, этот ребёнок; ориентир 149 ₽", 149, 3, "неделя", 12, 1),
    ("Маффины, пакет", "1 пакет в неделю; ориентир 199 ₽", 199, 1, "неделя", 12, 1),
    ("Сыр Ламбер", "кусочек ~250 г в неделю, Магнит весовой ~1 280 ₽/кг", 320, 1, "неделя", 12, 1),
    ("Сосиски", "1 пачка в неделю, Магнит; ориентир 250 ₽", 250, 1, "неделя", 12, 1),
    ("Картофель", "норма школьника ~200 г/день = 1,4 кг/нед, Магнит ~80 ₽/кг", 80, 1.4, "неделя", 12, 1),
    ("Морковь", "норма ~400 г/нед, Магнит сейчас ~44 ₽/кг", 44, 0.4, "неделя", 12, 1),
    ("Капуста белокочанная", "норма ~500 г/нед, ориентир 55 ₽/кг", 55, 0.5, "неделя", 12, 1),
    ("Лук репчатый", "норма ~300 г/нед, ориентир 55 ₽/кг", 55, 0.3, "неделя", 12, 1),
    ("Свёкла", "норма ~200 г/нед, Магнит ~70 ₽/кг", 70, 0.2, "неделя", 12, 1),
    ("Огурцы", "норма ~350 г/нед, среднее за год ~150 ₽/кг", 150, 0.35, "неделя", 12, 1),
    ("Помидоры", "норма ~350 г/нед, среднее за год ~180 ₽/кг", 180, 0.35, "неделя", 12, 1),
    ("Яблоки", "норма фруктов школьника ~220 г/день ≈ 1,5 кг/нед, ~130 ₽/кг", 130, 1.5, "неделя", 12, 1),
]

ws = wb.create_sheet("Продукты")
ws.sheet_properties.tabColor = "C45C26"
ws["A1"] = "Продукты — этот ребёнок (100%). То, что назвали, плюс овощи/яблоки по норме школьника"
ws["A1"].font = font_title
ws.merge_cells("A1:J1")
ws["A2"] = "Период можно ставить «неделя»: в год = цена × сколько раз × 52. Жёлтое меняйте. Комки «Магнит 15 000» и «Пятёрочка 15 000» убраны — здесь разбивка. Мёд — на листе общих."
ws["A2"].font = font_sub
ws.merge_cells("A2:J2")
headers = ["№", "Что", "Пояснение", "Цена единицы, ₽", "Сколько раз", "Период", "Месяцев в году", "Каждые N", "В год, ₽", "В месяц, ₽"]
for c, h in enumerate(headers, 1):
    ws.cell(4, c, h)
style_header(ws, 4, 10)

for i, row in enumerate(PROD_CHILD, 1):
    r = 4 + i
    name, note, amount, times, period, months, every = row
    ws.cell(r, 1, i).alignment = center
    ws.cell(r, 1).border = thin
    ws.cell(r, 1).font = font_n
    ws.cell(r, 2, name).alignment = left
    ws.cell(r, 2).border = thin
    ws.cell(r, 2).font = font_n
    ws.cell(r, 3, note).alignment = left
    ws.cell(r, 3).border = thin
    ws.cell(r, 3).font = font_n
    c = ws.cell(r, 4, amount)
    c.number_format = money_fmt
    style_input(c)
    style_input(ws.cell(r, 5, times))
    style_input(ws.cell(r, 6, period))
    style_input(ws.cell(r, 7, months))
    style_input(ws.cell(r, 8, every))
    fy = f'=IF(F{r}="год",D{r}*E{r},IF(F{r}="неделя",D{r}*E{r}*52/IF(H{r}=0,1,H{r}),D{r}*E{r}*(G{r}/IF(H{r}=0,1,H{r}))))'
    cy = ws.cell(r, 9, fy)
    cy.number_format = money_fmt
    style_out(cy, True)
    cm = ws.cell(r, 10, f"=IF(I{r}=0,0,I{r}/12)")
    cm.number_format = money_fmt
    style_out(cm, True)
    ws.row_dimensions[r].height = 20

last = 4 + len(PROD_CHILD)
tot = last + 1
ws.cell(tot, 1, "Итого продукты на ребёнка")
ws.merge_cells(start_row=tot, start_column=1, end_row=tot, end_column=8)
for c in range(1, 11):
    ws.cell(tot, c).fill = fill_sum
    ws.cell(tot, c).font = font_sum
    ws.cell(tot, c).border = thin
ws.cell(tot, 9, f"=SUM(I5:I{last})")
ws.cell(tot, 9).number_format = money_fmt
ws.cell(tot, 10, f"=SUM(J5:J{last})")
ws.cell(tot, 10).number_format = money_fmt
ws.cell(tot, 9).font = font_sum
ws.cell(tot, 10).font = font_sum

dv = DataValidation(type="list", formula1='"неделя,месяц,год"', allow_blank=False)
dv.add(f"F5:F{last}")
ws.add_data_validation(dv)
ws.auto_filter.ref = f"A4:J{last}"
widths(ws, {"A": 5, "B": 36, "C": 62, "D": 16, "E": 12, "F": 12, "G": 16, "H": 12, "I": 14, "J": 14})
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "1:4"
prod_child_tot = tot

# ===================== SHARED /3 =====================
# name, note, amount, times, period, months, every, split default 3
# Utilities broken down to ~7000 + rent + lake + NY table
SHARED = [
    ("Аренда квартиры", "15 000 ₽ в месяц, общее", 15000, 1, "месяц", 12, 1, 3),
    ("Электричество", "статья коммуналки, ориентир", 1800, 1, "месяц", 12, 1, 3),
    ("Содержание имущества", "статья коммуналки", 2000, 1, "месяц", 12, 1, 3),
    ("Вода (ХВС + ГВС + сток)", "часто забывают в списке", 1100, 1, "месяц", 12, 1, 3),
    ("Отопление (среднее за год)", "зимой больше, летом 0; здесь среднее", 1000, 1, "месяц", 12, 1, 3),
    ("Газ", "статья коммуналки", 300, 1, "месяц", 12, 1, 3),
    ("Твёрдые отходы (ТКО)", "вывоз мусора", 220, 1, "месяц", 12, 1, 3),
    ("Домофон", "статья коммуналки", 80, 1, "месяц", 12, 1, 3),
    ("Интернет", "часто не в квитанции ЖКХ", 600, 1, "месяц", 12, 1, 3),
    ("Мыло туалетное", "1 уп. в месяц", 120, 1, "месяц", 12, 1, 3),
    ("Ушные капли", "1 флакон в год", 300, 1, "год", 12, 1, 3),
    ("Перекись водорода", "1 флакон в год", 45, 1, "год", 12, 1, 3),
    ("Пластыри", "1 уп. в месяц", 99, 1, "месяц", 12, 1, 3),
    ("Йод", "1 пузырёк в год", 59, 1, "год", 12, 1, 3),
    ("Ватные диски", "1 уп. в месяц", 99, 1, "месяц", 12, 1, 3),
    ("Ватные палочки", "1 уп. в месяц", 89, 1, "месяц", 12, 1, 3),
    ("Ariel Color 6 кг", "1 уп. на 2 месяца ~1 399 ₽", 1399, 1, "месяц", 12, 2, 3),
    ("Lenor ополаскиватель", "1 пузырёк на 2 месяца ~349 ₽", 349, 1, "месяц", 12, 2, 3),
    ("Поездки на озеро летом", "5 000 ₽ поездка (еда+бензин), 5 раз/мес, 3 тёплых месяца", 5000, 5, "месяц", 3, 1, 3),
    ("Новый год: общий стол", "общий праздник — делим на всех", 6000, 1, "год", 12, 1, 3),
    ("Бензин", "15 000 ₽ в месяц на машину", 15000, 1, "месяц", 12, 1, 3),
    ("Масло моторное 5 л", "1 канистра на 3 месяца, ориентир 3 200 ₽", 3200, 1, "месяц", 12, 3, 3),
    ("Незамерзайка", "1 банка в месяц, ориентир 250 ₽", 250, 1, "месяц", 12, 1, 3),
    ("Поездки в другие города", "20 000 ₽ раз в 3 месяца, дорога+жильё+еда+сувениры", 20000, 1, "месяц", 12, 3, 3),
    ("Мёд, 3,5 л", "ездим закупать, 3,5 л в месяц, пасека ~750 ₽/л; общее ÷3", 750, 3.5, "месяц", 12, 1, 3),
    ("Аптека Максавит", "сумму подставьте из выписки", 0, 1, "месяц", 12, 1, 3),
    ("Аптека на Велижской (Мир лекарств)", "дешёвая рядом, сумму из выписки", 0, 1, "месяц", 12, 1, 3),
]

ws = wb.create_sheet("Общие ÷3")
ws.sheet_properties.tabColor = "6B8F71"
ws["A1"] = "Общие расходы семьи — доля этого ребёнка = семья ÷ «На сколько делим»"
ws["A1"].font = font_title
ws.merge_cells("A1:L1")
ws["A2"] = "Коммуналка разобрана по строкам (ориентир под ваши 7 000 ₽ + интернет). Поправьте жёлтое под квитанцию. Озеро: 5 000 × 5 × 3 мес. Новый год стол — общий, не как день рождения."
ws["A2"].font = font_sub
ws.merge_cells("A2:L2")
headers = ["№", "Что", "Пояснение", "Сумма платежа, ₽", "Сколько раз", "Период", "Месяцев в году", "Каждые N", "На сколько делим", "Семья в год, ₽", "Ребёнок в год, ₽", "Ребёнок в месяц, ₽"]
for c, h in enumerate(headers, 1):
    ws.cell(4, c, h)
style_header(ws, 4, 12)

for i, row in enumerate(SHARED, 1):
    r = 4 + i
    name, note, amount, times, period, months, every, split = row
    ws.cell(r, 1, i).alignment = center
    ws.cell(r, 1).border = thin
    ws.cell(r, 1).font = font_n
    ws.cell(r, 2, name).alignment = left
    ws.cell(r, 2).border = thin
    ws.cell(r, 2).font = font_n
    ws.cell(r, 2).fill = fill_green
    ws.cell(r, 3, note).alignment = left
    ws.cell(r, 3).border = thin
    ws.cell(r, 3).font = font_n
    ws.cell(r, 3).fill = fill_green
    c = ws.cell(r, 4, amount)
    c.number_format = money_fmt
    style_input(c)
    style_input(ws.cell(r, 5, times))
    style_input(ws.cell(r, 6, period))
    style_input(ws.cell(r, 7, months))
    style_input(ws.cell(r, 8, every))
    style_input(ws.cell(r, 9, split))
    fy = f'=IF(F{r}="год",D{r}*E{r},D{r}*E{r}*(G{r}/IF(H{r}=0,1,H{r})))'
    cf = ws.cell(r, 10, fy)
    cf.number_format = money_fmt
    style_out(cf, True)
    cc = ws.cell(r, 11, f"=IF(I{r}=0,0,J{r}/I{r})")
    cc.number_format = money_fmt
    style_out(cc, True)
    cm = ws.cell(r, 12, f"=IF(K{r}=0,0,K{r}/12)")
    cm.number_format = money_fmt
    style_out(cm, True)
    ws.row_dimensions[r].height = 20

last = 4 + len(SHARED)
tot = last + 1
ws.cell(tot, 1, "Итого общие (доля ребёнка)")
ws.merge_cells(start_row=tot, start_column=1, end_row=tot, end_column=9)
for c in range(1, 13):
    ws.cell(tot, c).fill = fill_sum
    ws.cell(tot, c).font = font_sum
    ws.cell(tot, c).border = thin
ws.cell(tot, 10, f"=SUM(J5:J{last})")
ws.cell(tot, 11, f"=SUM(K5:K{last})")
ws.cell(tot, 12, f"=SUM(L5:L{last})")
for c in (10, 11, 12):
    ws.cell(tot, c).number_format = money_fmt
    ws.cell(tot, c).font = font_sum

# subtotal utilities note
ws.cell(tot + 2, 2, "Проверка коммуналки (электричество…интернет, без аренды и озера): поправьте жёлтое, чтобы сошлось с квитанцией.")
ws.cell(tot + 2, 2).font = font_hint
ws.merge_cells(start_row=tot + 2, start_column=2, end_row=tot + 2, end_column=6)
# rows 2-9 of shared are utilities (indices 2-9 → excel rows 6-13) rent is row 5
ws.cell(tot + 3, 2, "Семья в месяц, коммуналка+интернет (строки 2–9):")
ws.cell(tot + 3, 2).font = font_n
ws.cell(tot + 3, 4, f"=SUM(J6:J13)/12")
ws.cell(tot + 3, 4).number_format = money_fmt
ws.cell(tot + 3, 4).fill = fill_gold
ws.cell(tot + 3, 4).border = thin

dv = DataValidation(type="list", formula1='"месяц,год"', allow_blank=False)
dv.add(f"F5:F{last}")
ws.add_data_validation(dv)
ws.auto_filter.ref = f"A4:L{last}"
widths(ws, {"A": 5, "B": 32, "C": 52, "D": 16, "E": 12, "F": 12, "G": 16, "H": 12, "I": 16, "J": 16, "K": 16, "L": 18})
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 0
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "1:4"
shared_tot = tot

# ===================== WEEKENDS =====================
# Cinema and zoo with mom/2
ws = wb.create_sheet("Выходные")
ws.sheet_properties.tabColor = "C9A227"
ws["A1"] = "Кино, цирк, кафе — мамин чек делим на двоих детей"
ws["A1"].font = font_title
ws.merge_cells("A1:H1")
ws["A2"] = "Вы идёте из-за детей. На ребёнка = его билет/заказ + ваш÷2. Кафе — раз в неделю (52 раза). Цирк — шапито «Виват» Иваново, билеты 1 000–2 500 ₽, поставила 1 500."
ws["A2"].font = font_sub
ws.merge_cells("A2:H2")

headers = ["№", "Что", "Чек ребёнка, ₽", "Чек взрослого, ₽", "Раз в год", "Семья (2 детей+вы) в год, ₽", "Этот ребёнок в год, ₽", "Этот ребёнок в месяц, ₽"]
for c, h in enumerate(headers, 1):
    ws.cell(4, c, h)
style_header(ws, 4, 8)

weekends = [
    (5, 1, "Кино А113 Иваново, раз в месяц", 380, 380, 12),
    (6, 2, "Зоопарк Иваново, раз в год (ivzoopark.ru с 01.01.2026)", 200, 350, 1),
    (7, 3, "Цирк, раз в год (шапито Виват Иваново ~1 000–2 500)", 1500, 1500, 1),
    (8, 4, "Бургер Кинг / кафе, раз в неделю", 700, 800, 52),
]
for r, n, name, ch, ad, times in weekends:
    ws.cell(r, 1, n).alignment = center
    ws.cell(r, 1).border = thin
    ws.cell(r, 2, name).alignment = left
    ws.cell(r, 2).border = thin
    ws.cell(r, 2).fill = fill_green
    ws.cell(r, 3, ch)
    ws.cell(r, 4, ad)
    ws.cell(r, 5, times)
    ws.cell(r, 6, f"=(C{r}*2+D{r})*E{r}")
    ws.cell(r, 7, f"=(C{r}+D{r}/2)*E{r}")
    ws.cell(r, 8, f"=G{r}/12")
    for c in (3, 4, 5):
        style_input(ws.cell(r, c))
        if c != 5:
            ws.cell(r, c).number_format = money_fmt
    for c in (6, 7, 8):
        ws.cell(r, c).number_format = money_fmt
        style_out(ws.cell(r, c), True)

ws["A9"] = ""
ws["B9"] = "Итого выходные на этого ребёнка"
ws.merge_cells("B9:E9")
ws["F9"] = "=SUM(F5:F8)"
ws["G9"] = "=SUM(G5:G8)"
ws["H9"] = "=SUM(H5:H8)"
for c in range(1, 9):
    ws.cell(9, c).fill = fill_sum
    ws.cell(9, c).font = font_sum
    ws.cell(9, c).border = thin
for c in (6, 7, 8):
    ws.cell(9, c).number_format = money_fmt

ws["A11"] = "Подсказка"
ws["A11"].font = font_b
ws["B11"] = "Кафе: C8 — средний заказ ребёнка, D8 — ваш. Если ходите реже, чем каждую неделю, уменьшите E8 (например 24 = раз в две недели)."
ws["B11"].font = font_hint
ws.merge_cells("B11:H11")

widths(ws, {"A": 5, "B": 58, "C": 18, "D": 20, "E": 12, "F": 28, "G": 22, "H": 24})
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "1:4"

# ===================== LARGE SHARED ASSETS =====================
# name, note, buy, residual, years, split
ASSETS = [
    ("Автомобиль", "стоит ~500 000; если продадите — впишите сумму в «Можно продать»", 500000, 0, 10, 3),
    ("Ёлка искусственная", "общая, 25 000 ₽", 25000, 0, 10, 3),
    ("Украшения на ёлку", "общие, 15 000 ₽", 15000, 0, 10, 3),
    ("Телевизор маленький", "купили за 10 000 ₽", 10000, 0, 7, 3),
    ("Колонка Sony", "2 000 ₽", 2000, 0, 5, 3),
]

ws = wb.create_sheet("Имущество общее")
ws.sheet_properties.tabColor = "8B6914"
ws["A1"] = "Крупные общие вещи: амортизация = (купили − можно продать) ÷ срок, потом ÷ на семью"
ws["A1"].font = font_title
ws.merge_cells("A1:I1")
ws["A2"] = "Машина: если через 10 лет продадите, например, за 150 000 — впишите 150 000 в колонку E, износ станет 350 000. Жёлтое меняйте."
ws["A2"].font = font_sub
ws.merge_cells("A2:I2")
headers = ["№", "Что", "Пояснение", "Купили за, ₽", "Можно продать за, ₽", "Срок, лет", "На сколько делим", "Семья в год, ₽", "Ребёнок в год, ₽", "Ребёнок в месяц, ₽"]
for c, h in enumerate(headers, 1):
    ws.cell(4, c, h)
style_header(ws, 4, 10)

for i, row in enumerate(ASSETS, 1):
    r = 4 + i
    name, note, buy, residual, years, split = row
    ws.cell(r, 1, i).alignment = center
    ws.cell(r, 1).border = thin
    ws.cell(r, 2, name).alignment = left
    ws.cell(r, 2).border = thin
    ws.cell(r, 2).fill = fill_green
    ws.cell(r, 2).font = font_n
    ws.cell(r, 3, note).alignment = left
    ws.cell(r, 3).border = thin
    ws.cell(r, 3).font = font_n
    c = ws.cell(r, 4, buy)
    c.number_format = money_fmt
    style_input(c)
    c = ws.cell(r, 5, residual)
    c.number_format = money_fmt
    style_input(c)
    style_input(ws.cell(r, 6, years))
    style_input(ws.cell(r, 7, split))
    # family year = max(buy-residual,0)/years
    cf = ws.cell(r, 8, f"=IF(F{r}=0,0,MAX(D{r}-E{r},0)/F{r})")
    cf.number_format = money_fmt
    style_out(cf, True)
    cc = ws.cell(r, 9, f"=IF(G{r}=0,0,H{r}/G{r})")
    cc.number_format = money_fmt
    style_out(cc, True)
    cm = ws.cell(r, 10, f"=I{r}/12")
    cm.number_format = money_fmt
    style_out(cm, True)
    ws.row_dimensions[r].height = 22

last_a = 4 + len(ASSETS)
tot_a = last_a + 1
ws.cell(tot_a, 1, "Итого имущество (доля ребёнка)")
ws.merge_cells(start_row=tot_a, start_column=1, end_row=tot_a, end_column=7)
for c in range(1, 11):
    ws.cell(tot_a, c).fill = fill_sum
    ws.cell(tot_a, c).font = font_sum
    ws.cell(tot_a, c).border = thin
ws.cell(tot_a, 8, f"=SUM(H5:H{last_a})")
ws.cell(tot_a, 9, f"=SUM(I5:I{last_a})")
ws.cell(tot_a, 10, f"=SUM(J5:J{last_a})")
for c in (8, 9, 10):
    ws.cell(tot_a, c).number_format = money_fmt
    ws.cell(tot_a, c).font = font_sum

ws.cell(tot_a + 2, 2, "Машина без продажи: 500 000 ÷ 10 лет = 50 000 ₽/год семья, на ребёнка ÷3 ≈ 1 389 ₽/мес. Если продадите — износ меньше.")
ws.cell(tot_a + 2, 2).font = font_hint
ws.merge_cells(start_row=tot_a + 2, start_column=2, end_row=tot_a + 2, end_column=8)

widths(ws, {"A": 5, "B": 28, "C": 62, "D": 16, "E": 20, "F": 12, "G": 16, "H": 16, "I": 18, "J": 20})
ws.page_setup.orientation = "landscape"
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "1:4"
assets_tot = tot_a

# ===================== SUMMARY =====================
ws = wb.create_sheet("Сводка", 1)
ws.sheet_properties.tabColor = GOLD
ws["A1"] = "Сводка бюджета одного ребёнка"
ws["A1"].font = Font(name="Calibri", bold=True, color=NAVY, size=18)
ws.merge_cells("A1:D1")
ws["A2"] = "Все суммы ниже — формулы с других листов. Правите исходные жёлтые ячейки, сводка обновится."
ws["A2"].font = font_sub
ws.merge_cells("A2:D2")

ws["A4"] = "Вещи по категориям"
ws["A4"].font = font_b
headers = ["Категория", "Купить сразу, ₽", "Амортизация в год, ₽", "Амортизация в месяц, ₽"]
for c, h in enumerate(headers, 1):
    ws.cell(5, c, h)
style_header(ws, 5, 4)
ws.auto_filter.ref = None  # don't filter summary

for i, sh in enumerate(THING_SHEETS):
    r = 6 + i
    last, tot = sheet_ranges[sh]
    ws.cell(r, 1, sh).font = font_n
    ws.cell(r, 1).border = thin
    ws.cell(r, 1).alignment = left
    ws.cell(r, 2, f"='{sh}'!F{tot}")
    ws.cell(r, 3, f"='{sh}'!I{tot}")
    ws.cell(r, 4, f"='{sh}'!J{tot}")
    for c in (2, 3, 4):
        ws.cell(r, c).number_format = money_fmt
        style_out(ws.cell(r, c), True)
    if i % 2:
        ws.cell(r, 1).fill = fill_zebra

r0 = 6
r1 = 5 + len(THING_SHEETS)
rt = r1 + 1
ws.cell(rt, 1, "Итого вещи")
ws.cell(rt, 1).font = font_sum
ws.cell(rt, 1).fill = fill_sum
ws.cell(rt, 2, f"=SUM(B{r0}:B{r1})")
ws.cell(rt, 3, f"=SUM(C{r0}:C{r1})")
ws.cell(rt, 4, f"=SUM(D{r0}:D{r1})")
for c in range(1, 5):
    ws.cell(rt, c).fill = fill_sum
    ws.cell(rt, c).font = font_sum
    ws.cell(rt, c).border = thin
    if c > 1:
        ws.cell(rt, c).number_format = money_fmt

# regular blocks
ws.cell(rt + 2, 1, "Регулярные и общие")
ws.cell(rt + 2, 1).font = font_b
headers2 = ["Блок", "Семья / год, ₽", "Ребёнок / год, ₽", "Ребёнок / месяц, ₽"]
for c, h in enumerate(headers2, 1):
    ws.cell(rt + 3, c, h)
    ws.cell(rt + 3, c).fill = fill_navy
    ws.cell(rt + 3, c).font = font_h
    ws.cell(rt + 3, c).alignment = center
    ws.cell(rt + 3, c).border = thin

ws.cell(rt + 4, 1, "Регулярные только ребёнок")
ws.cell(rt + 4, 2, "—")
ws.cell(rt + 4, 3, f"='Регулярные ребёнок'!I{rec_child_tot}")
ws.cell(rt + 4, 4, f"='Регулярные ребёнок'!J{rec_child_tot}")

ws.cell(rt + 5, 1, "Продукты этот ребёнок (Магнит + нормы овощей)")
ws.cell(rt + 5, 2, "—")
ws.cell(rt + 5, 3, f"='Продукты'!I{prod_child_tot}")
ws.cell(rt + 5, 4, f"='Продукты'!J{prod_child_tot}")

ws.cell(rt + 6, 1, "Общие ÷3 (квартира, коммуналка, мёд, озеро, машина, НГ стол)")
ws.cell(rt + 6, 2, f"='Общие ÷3'!J{shared_tot}")
ws.cell(rt + 6, 3, f"='Общие ÷3'!K{shared_tot}")
ws.cell(rt + 6, 4, f"='Общие ÷3'!L{shared_tot}")

ws.cell(rt + 7, 1, "Кино, цирк, кафе, зоопарк (мама ÷ 2)")
ws.cell(rt + 7, 2, "=Выходные!F9")
ws.cell(rt + 7, 3, "=Выходные!G9")
ws.cell(rt + 7, 4, "=Выходные!H9")

ws.cell(rt + 8, 1, "Имущество общее (машина, ёлка, ТВ) ÷3")
ws.cell(rt + 8, 2, "—")
ws.cell(rt + 8, 3, f"='Имущество общее'!I{assets_tot}")
ws.cell(rt + 8, 4, f"='Имущество общее'!J{assets_tot}")

for r in range(rt + 4, rt + 9):
    ws.cell(r, 1).border = thin
    ws.cell(r, 1).font = font_n
    for c in (2, 3, 4):
        if ws.cell(r, c).value != "—":
            ws.cell(r, c).number_format = money_fmt
        style_out(ws.cell(r, c), True)
    ws.cell(r, 1).fill = fill_green

ws.cell(rt + 9, 1, "Итого регулярные + продукты + имущество на ребёнка")
ws.cell(rt + 9, 3, f"=C{rt+4}+C{rt+5}+C{rt+6}+C{rt+7}+C{rt+8}")
ws.cell(rt + 9, 4, f"=D{rt+4}+D{rt+5}+D{rt+6}+D{rt+7}+D{rt+8}")
for c in range(1, 5):
    ws.cell(rt + 9, c).fill = fill_sum
    ws.cell(rt + 9, c).font = font_sum
    ws.cell(rt + 9, c).border = thin
ws.cell(rt + 9, 3).number_format = money_fmt
ws.cell(rt + 9, 4).number_format = money_fmt

# grand
ws.cell(rt + 11, 1, "СОДЕРЖАНИЕ РЕБЁНКА")
ws.cell(rt + 11, 1).font = Font(name="Calibri", bold=True, color=WHITE, size=14)
ws.merge_cells(start_row=rt + 11, start_column=1, end_row=rt + 11, end_column=2)
ws.cell(rt + 11, 3, f"=C{rt}+C{rt+9}")
ws.cell(rt + 11, 4, f"=D{rt}+D{rt+9}")
for c in range(1, 5):
    ws.cell(rt + 11, c).fill = fill_sum
    ws.cell(rt + 11, c).font = Font(name="Calibri", bold=True, color=WHITE, size=13)
    ws.cell(rt + 11, c).border = thin
ws.cell(rt + 11, 3).number_format = money_fmt
ws.cell(rt + 11, 4).number_format = money_fmt
ws.cell(rt + 12, 3, "в год")
ws.cell(rt + 12, 4, "в месяц")
ws.cell(rt + 12, 3).font = font_hint
ws.cell(rt + 12, 4).font = font_hint

ws.cell(rt + 13, 1, "День рождения целиком на ребёнка (подарок + стол). Новый год: подарок ему, стол — в «Общие ÷3».")
ws.cell(rt + 13, 1).font = font_hint
ws.merge_cells(start_row=rt + 13, start_column=1, end_row=rt + 13, end_column=4)
ws.cell(rt + 14, 1, "Коммуналка: сумма строк электричество…интернет должна быть около 7 000 ₽/мес на семью — смотрите проверку на листе «Общие ÷3».")
ws.cell(rt + 14, 1).font = font_hint
ws.merge_cells(start_row=rt + 14, start_column=1, end_row=rt + 14, end_column=4)

widths(ws, {"A": 52, "B": 20, "C": 24, "D": 26})
ws.page_setup.orientation = "portrait"
ws.page_setup.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.print_title_rows = "1:5"
ws.row_dimensions[1].height = 26
ws.row_dimensions[rt + 11].height = 28

# print area / freeze
ws.freeze_panes = "A6"

# bar chart by section — year spend (amort + we'll use column C things only for things chart)
chart = BarChart()
chart.type = "bar"
chart.style = 10
chart.title = "Что на сколько: вещи, амортизация в год"
chart.y_axis.title = None
chart.x_axis.title = None
data = Reference(ws, min_col=3, min_row=5, max_row=r1)
cats = Reference(ws, min_col=1, min_row=6, max_row=r1)
chart.add_data(data, titles_from_data=True)
chart.set_categories(cats)
chart.shape = 4
chart.legend = None
chart.height = 10
chart.width = 18
ws.add_chart(chart, "F4")

# ===================== finish =====================
# reorder: Как пользоваться, Сводка, then rest already created
# already inserted Сводка at index 1

out = "/home/user/Byudzhet-rebenka.xlsx"
wb.save(out)
print("saved", out)
print("sheets", wb.sheetnames)

# ===================== INTERACTIVE HTML (same data) =====================
import json
from pathlib import Path

WEEKENDS_DATA = [
    ["Кино А113 Иваново, раз в месяц", 380, 380, 12],
    ["Зоопарк Иваново, раз в год (ivzoopark.ru с 01.01.2026)", 200, 350, 1],
    ["Цирк, раз в год (шапито Виват Иваново ~1 000–2 500)", 1500, 1500, 1],
    ["Бургер Кинг / кафе, раз в неделю", 700, 800, 52],
]

payload = {
    "THING_SHEETS": THING_SHEETS,
    "THINGS": [list(x) for x in THINGS],
    "REC_CHILD": [list(x) for x in REC_CHILD],
    "PROD_CHILD": [list(x) for x in PROD_CHILD],
    "SHARED": [list(x) for x in SHARED],
    "ASSETS": [list(x) for x in ASSETS],
    "WEEKENDS": WEEKENDS_DATA,
}

def rec_year(amount, times, period, months, every):
    every = every or 1
    if period == "год":
        return amount * times
    if period == "неделя":
        return amount * times * (52 / every)
    return amount * times * (months / every)

def amort_html(price, term, unit):
    term = max(float(term) or 1, 0.01)
    price = float(price) or 0
    if unit == "месяц":
        m = price / term
        return price, m * 12, m
    y = price / term
    return price, y, y / 12

things_ay = 0
for row in THINGS:
    things_ay += amort_html(row[5], row[6], row[7])[1]
rc_y = sum(rec_year(*r[2:]) for r in REC_CHILD)
pr_y = sum(rec_year(*r[2:]) for r in PROD_CHILD)
sh_cy = sum(rec_year(*r[2:7]) / r[7] for r in SHARED)
w_cy = sum((ch + ad / 2) * times for _, ch, ad, times in WEEKENDS_DATA)
as_cy = sum(max(buy - residual, 0) / years / split for _, _, buy, residual, years, split in ASSETS)
grand_m = (things_ay + rc_y + pr_y + sh_cy + w_cy + as_cy) / 12

tpl_paths = [
    Path("/home/user/budget_interactive.html"),
    Path("/tmp/budget_interactive.html"),
]
tpl = next((p.read_text(encoding="utf-8") for p in tpl_paths if p.exists()), None)
if not tpl or "__DATA__" not in tpl:
    raise SystemExit("нет шаблона budget_interactive.html")
html = tpl.replace("__DATA__", json.dumps(payload, ensure_ascii=False))
html_path = "/home/user/spisok-v-shkolu-detskiy-mir.html"
Path(html_path).write_text(html, encoding="utf-8")
print("html", html_path, "grand_m", round(grand_m), "interactive")
raise SystemExit(0)

# ===================== OLD STATIC HTML (unused) =====================
def money_html(n):
    return f"{int(round(n)):,}".replace(",", " ") + " ₽"

def amort_html(price, term, unit):
    term = max(float(term) or 1, 0.01)
    price = float(price) or 0
    if unit == "месяц":
        m = price / term
        return price, m * 12, m
    y = price / term
    return price, y, y / 12

def rec_year(amount, times, period, months, every):
    every = every or 1
    if period == "год":
        return amount * times
    return amount * times * (months / every)

cat_totals = []
for sh in THING_SHEETS:
    rows = [x for x in THINGS if x[0] == sh]
    buy = ay = am = 0
    items = []
    for row in rows:
        _, name, qty, shop, url, price, term, unit = row
        p, y, m = amort_html(price, term, unit)
        buy += p; ay += y; am += m
        items.append((name, qty, shop, url, p, term, unit, y, m))
    cat_totals.append((sh, buy, ay, am, items))

rec_child_rows = []
rc_y = rc_m = 0
for name, note, amount, times, period, months, every in REC_CHILD:
    y = rec_year(amount, times, period, months, every)
    rec_child_rows.append((name, note, amount, times, period, months, every, y, y/12))
    rc_y += y; rc_m += y/12

sh_rows = []
sh_fy = sh_cy = 0
for name, note, amount, times, period, months, every, split in SHARED:
    fy = rec_year(amount, times, period, months, every)
    cy = fy / split
    sh_rows.append((name, note, amount, times, period, months, every, split, fy, cy, cy/12))
    sh_fy += fy; sh_cy += cy

week_data = [
    ("Кино А113 Иваново, раз в месяц", 380, 380, 12),
    ("Зоопарк Иваново, раз в год", 200, 350, 1),
    ("Цирк, раз в год (Виват Иваново)", 1500, 1500, 1),
    ("Бургер Кинг / кафе, раз в неделю", 700, 800, 52),
]
w_fy = w_cy = 0
week_rows = []
for name, ch, ad, times in week_data:
    fy = (ch*2 + ad) * times
    cy = (ch + ad/2) * times
    week_rows.append((name, ch, ad, times, fy, cy, cy/12))
    w_fy += fy; w_cy += cy

as_rows = []
as_fy = as_cy = 0
for name, note, buy, residual, years, split in ASSETS:
    fy = max(buy - residual, 0) / years
    cy = fy / split
    as_rows.append((name, note, buy, residual, years, split, fy, cy, cy/12))
    as_fy += fy; as_cy += cy

things_buy = sum(t[1] for t in cat_totals)
things_ay = sum(t[2] for t in cat_totals)
things_am = sum(t[3] for t in cat_totals)
grand_y = things_ay + rc_y + sh_cy + w_cy + as_cy
grand_m = grand_y / 12
max_ay = max(t[2] for t in cat_totals) or 1

# combined chart data: categories + regular blocks
chart_blocks = [(n, ay) for n, _, ay, _, _ in cat_totals]
chart_blocks += [
    ("Регулярные ребёнок", rc_y),
    ("Общие ÷3", sh_cy),
    ("Выходные", w_cy),
    ("Имущество общее", as_cy),
]
max_block = max(b[1] for b in chart_blocks) or 1

bars = ""
for name, val in chart_blocks:
    pct = 100 * val / max_block
    bars += f'''<div class="barrow"><div class="blab">{name}</div>
      <div class="btrack"><div class="bfill" style="width:{pct:.1f}%"></div></div>
      <div class="bval">{money_html(val)}</div></div>\n'''

def table_things(items):
    rows = ""
    for i, (name, qty, shop, url, p, term, unit, y, m) in enumerate(items, 1):
        link = f'<a href="{url}">открыть</a>' if url else "—"
        rows += f"<tr><td>{i}</td><td>{name}</td><td>{qty}</td><td>{shop}</td><td>{link}</td><td class='n'>{money_html(p)}</td><td class='n'>{term}</td><td>{unit}</td><td class='n'>{money_html(y)}</td><td class='n'>{money_html(m)}</td></tr>\n"
    return rows

html = f'''<!DOCTYPE html>
<html lang="ru">
<head>
<meta charset="UTF-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>Бюджет ребёнка — печать</title>
<style>
:root {{ --navy:#1e3a5f; --line:#e0d9cc; }}
* {{ box-sizing:border-box; }}
body {{ margin:0; background:#f3efe6; color:#1b1b1b; font-family:"Segoe UI",Calibri,Arial,sans-serif; }}
.page {{ max-width:1100px; margin:0 auto; padding:18px 14px 48px; }}
h1 {{ margin:0 0 4px; font-size:24px; color:var(--navy); }}
.sub {{ color:#555; font-size:13px; margin:0 0 14px; }}
.cards {{ display:grid; grid-template-columns:repeat(3,1fr); gap:8px; margin:10px 0 16px; }}
.card {{ background:#fff; border:1px solid var(--line); border-radius:10px; padding:10px 12px; }}
.card b {{ display:block; font-size:18px; color:var(--navy); margin-top:2px; }}
.card span {{ font-size:12px; color:#666; }}
.card.dark {{ background:var(--navy); color:#fff; }}
.card.dark b {{ color:#fff; }}
.card.dark span {{ color:#c9d6e5; }}
h2 {{ margin:22px 0 8px; font-size:16px; color:var(--navy); border-bottom:2px solid var(--navy); padding-bottom:4px; page-break-after:avoid; }}
.chart {{ background:#fff; border:1px solid var(--line); border-radius:12px; padding:12px 14px 8px; margin:8px 0 18px; }}
.chart h2 {{ margin:0 0 10px; border:0; }}
.barrow {{ display:grid; grid-template-columns:160px 1fr 110px; gap:8px; align-items:center; margin:5px 0; }}
.blab {{ font-size:12px; }}
.btrack {{ background:#ece7dc; height:14px; border-radius:7px; overflow:hidden; }}
.bfill {{ height:100%; background:var(--navy); border-radius:7px; }}
.bval {{ font-size:12px; text-align:right; font-variant-numeric:tabular-nums; }}
table {{ width:100%; border-collapse:collapse; background:#fff; font-size:12px; }}
th,td {{ padding:5px 6px; border-bottom:1px solid var(--line); vertical-align:top; }}
th {{ background:var(--navy); color:#fff; text-align:left; font-size:11px; }}
td.n {{ text-align:right; white-space:nowrap; font-variant-numeric:tabular-nums; }}
tr:nth-child(even) td {{ background:#faf8f3; }}
.itog {{ background:var(--navy); color:#fff; display:flex; justify-content:space-between; padding:8px 10px; font-weight:700; font-size:13px; }}
a {{ color:#0563C1; }}
.note {{ color:#666; font-size:12px; margin-top:14px; }}
@media print {{
  body {{ background:#fff; }}
  .page {{ padding:0; max-width:none; }}
  .card, .chart, table {{ break-inside:avoid; }}
  h2 {{ break-after:avoid; }}
  a {{ color:#000; text-decoration:none; }}
  .noprint {{ display:none; }}
}}
</style>
</head>
<body>
<div class="page">
<h1>Бюджет одного ребёнка</h1>
<p class="sub">Печатная копия. Цифры те же, что в Excel Byudzhet-rebenka.xlsx. Жёлтые правки в Excel сюда сами не попадают — после правок в таблице попросите обновить HTML.</p>
<div class="cards">
  <div class="card"><span>Вещи купить сразу</span><b>{money_html(things_buy)}</b></div>
  <div class="card"><span>Амортизация вещей / мес</span><b>{money_html(things_am)}</b></div>
  <div class="card"><span>Регулярные ребёнка / мес</span><b>{money_html(rc_m)}</b></div>
  <div class="card"><span>Общие ÷3 / мес</span><b>{money_html(sh_cy/12)}</b></div>
  <div class="card"><span>Выходные + имущество / мес</span><b>{money_html((w_cy+as_cy)/12)}</b></div>
  <div class="card dark"><span>Содержание в месяц</span><b>{money_html(grand_m)}</b></div>
</div>

<div class="chart">
<h2>Что на сколько тратим (в год, на ребёнка)</h2>
{bars}
</div>
'''

for sh, buy, ay, am, items in cat_totals:
    html += f'<h2>{sh}</h2><table><thead><tr><th>№</th><th>Что</th><th>Сколько</th><th>Где</th><th></th><th>Цена</th><th>Срок</th><th>Ед.</th><th>В год</th><th>В месяц</th></tr></thead><tbody>'
    html += table_things(items)
    html += f"</tbody></table><div class='itog'><span>Итого: {sh}</span><span>покупка {money_html(buy)} · год {money_html(ay)} · месяц {money_html(am)}</span></div>\n"

html += '<h2>Регулярные — только ребёнок</h2><table><thead><tr><th>№</th><th>Что</th><th>Пояснение</th><th>Платёж</th><th>Раз</th><th>Период</th><th>Мес. в году</th><th>Каждые N</th><th>В год</th><th>В месяц</th></tr></thead><tbody>'
for i, row in enumerate(rec_child_rows, 1):
    name, note, amount, times, period, months, every, y, m = row
    html += f"<tr><td>{i}</td><td>{name}</td><td>{note}</td><td class='n'>{money_html(amount)}</td><td class='n'>{times}</td><td>{period}</td><td class='n'>{months}</td><td class='n'>{every}</td><td class='n'>{money_html(y)}</td><td class='n'>{money_html(m)}</td></tr>"
html += f"</tbody></table><div class='itog'><span>Итого регулярные ребёнок</span><span>год {money_html(rc_y)} · месяц {money_html(rc_m)}</span></div>"

html += '<h2>Общие ÷3</h2><table><thead><tr><th>№</th><th>Что</th><th>Пояснение</th><th>Платёж</th><th>Раз</th><th>Период</th><th>Мес.</th><th>N</th><th>Делим</th><th>Семья/год</th><th>Ребёнок/год</th><th>Ребёнок/мес</th></tr></thead><tbody>'
for i, row in enumerate(sh_rows, 1):
    name, note, amount, times, period, months, every, split, fy, cy, cm = row
    html += f"<tr><td>{i}</td><td>{name}</td><td>{note}</td><td class='n'>{money_html(amount)}</td><td class='n'>{times}</td><td>{period}</td><td class='n'>{months}</td><td class='n'>{every}</td><td class='n'>{split}</td><td class='n'>{money_html(fy)}</td><td class='n'>{money_html(cy)}</td><td class='n'>{money_html(cm)}</td></tr>"
html += f"</tbody></table><div class='itog'><span>Итого общие, доля ребёнка</span><span>семья {money_html(sh_fy)}/год · ребёнок {money_html(sh_cy)}/год · {money_html(sh_cy/12)}/мес</span></div>"

html += '<h2>Выходные — мамин чек ÷ 2</h2><table><thead><tr><th>№</th><th>Что</th><th>Чек ребёнка</th><th>Чек взрослого</th><th>Раз в год</th><th>Семья/год</th><th>Ребёнок/год</th><th>Ребёнок/мес</th></tr></thead><tbody>'
for i, row in enumerate(week_rows, 1):
    name, ch, ad, times, fy, cy, cm = row
    html += f"<tr><td>{i}</td><td>{name}</td><td class='n'>{money_html(ch)}</td><td class='n'>{money_html(ad)}</td><td class='n'>{times}</td><td class='n'>{money_html(fy)}</td><td class='n'>{money_html(cy)}</td><td class='n'>{money_html(cm)}</td></tr>"
html += f"</tbody></table><div class='itog'><span>Итого выходные</span><span>ребёнок {money_html(w_cy)}/год · {money_html(w_cy/12)}/мес</span></div>"

html += '<h2>Имущество общее (купили − продажа) ÷ срок ÷ 3</h2><table><thead><tr><th>№</th><th>Что</th><th>Купили</th><th>Продать</th><th>Срок, лет</th><th>Делим</th><th>Семья/год</th><th>Ребёнок/год</th><th>Ребёнок/мес</th></tr></thead><tbody>'
for i, row in enumerate(as_rows, 1):
    name, note, buy, residual, years, split, fy, cy, cm = row
    html += f"<tr><td>{i}</td><td>{name}<br><span style='color:#666;font-size:11px'>{note}</span></td><td class='n'>{money_html(buy)}</td><td class='n'>{money_html(residual)}</td><td class='n'>{years}</td><td class='n'>{split}</td><td class='n'>{money_html(fy)}</td><td class='n'>{money_html(cy)}</td><td class='n'>{money_html(cm)}</td></tr>"
html += f"</tbody></table><div class='itog'><span>Итого имущество</span><span>ребёнок {money_html(as_cy)}/год · {money_html(as_cy/12)}/мес</span></div>"

html += f'''<p class="note"><b>Содержание: {money_html(grand_m)} в месяц</b> ({money_html(grand_y)} в год).
Хотелочки 2 000 ₽/мес. Курс русский 300 ₽/мес. Поездки в города 20 000 раз в 3 месяца, общее ÷3.
День рождения 10 000 ему, новогодний стол — на всех.</p>
</div></body></html>'''

html_path = "/home/user/spisok-v-shkolu-detskiy-mir.html"
open(html_path, "w", encoding="utf-8").write(html)
print("html", html_path, "grand_m", round(grand_m))

