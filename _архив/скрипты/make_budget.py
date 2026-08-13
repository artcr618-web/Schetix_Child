#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import json
from pathlib import Path

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

from datetime import datetime
import shutil

from budget_blocks import BLOCKS, TILES, CHART_GROUPS, PARTS, row_year

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
GOTOV = ROOT / "gotovo"
ARXIV = ROOT / "arxiv"

NAVY = "1E3A5F"
GOLD = "C9A227"
MUTED = "5B6573"
INPUT_BG = "FFF8D6"

thin = Border(
    left=Side(style="thin", color="D9D2C4"),
    right=Side(style="thin", color="D9D2C4"),
    top=Side(style="thin", color="D9D2C4"),
    bottom=Side(style="thin", color="D9D2C4"),
)
fill_navy = PatternFill("solid", fgColor=NAVY)
fill_input = PatternFill("solid", fgColor=INPUT_BG)
fill_sum = PatternFill("solid", fgColor=NAVY)
fill_zebra = PatternFill("solid", fgColor="FAF8F3")

font_h = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
font_title = Font(name="Calibri", bold=True, color=NAVY, size=16)
font_sub = Font(name="Calibri", color=MUTED, size=10)
font_b = Font(name="Calibri", bold=True, size=11)
font_n = Font(name="Calibri", size=10)
font_sum = Font(name="Calibri", bold=True, color="FFFFFF", size=11)
center = Alignment(horizontal="center", vertical="center", wrap_text=True)
left = Alignment(horizontal="left", vertical="center", wrap_text=True)
right = Alignment(horizontal="right", vertical="center")
money_fmt = '#,##0" ₽"'


def style_header(ws, row, cols):
    for c in range(1, cols + 1):
        cell = ws.cell(row, c)
        cell.fill = fill_navy
        cell.font = font_h
        cell.alignment = center
        cell.border = thin
    ws.row_dimensions[row].height = 28
    ws.freeze_panes = f"A{row + 1}"


def style_input(cell):
    cell.fill = fill_input
    cell.border = thin
    cell.font = font_n
    cell.alignment = center


def col_letter(n):
    s = ""
    while n:
        n, rem = divmod(n - 1, 26)
        s = chr(65 + rem) + s
    return s


def layout(block):
    compact = bool(block.get("compact"))
    links = bool(block.get("links"))
    items = block.get("items") or []
    has_years = any(float(it.get("years") or 0) for it in items)
    has_period = any(it.get("period") for it in items)
    if has_years and has_period:
        term_h = "Срок / период"
    elif has_years:
        term_h = "Срок службы (лет)"
    else:
        term_h = "Период"
    headers = ["№", "Наименование"]
    cmap = {"num": 1, "name": 2}
    c = 3
    if not compact:
        headers.append("Количество")
        cmap["pcs"] = c
        c += 1
    if links:
        headers.append("Ссылка")
        cmap["url"] = c
        c += 1
    headers.append("Стоимость")
    cmap["price"] = c
    c += 1
    headers.append(term_h)
    cmap["term"] = c
    c += 1
    headers.append("В год, ₽")
    cmap["year"] = c
    c += 1
    headers.append("В месяц, ₽")
    cmap["month"] = c
    return headers, cmap


def year_formula(cmap, r):
    p = f"{col_letter(cmap['price'])}{r}"
    t = f"{col_letter(cmap['term'])}{r}"
    if "pcs" in cmap:
        q = f"IF({col_letter(cmap['pcs'])}{r}=0,1,{col_letter(cmap['pcs'])}{r})"
    else:
        q = "1"
    return (
        f"=IF(ISNUMBER({t}),IF({t}=0,0,{p}*{q}/{t}),"
        f'IF({t}="неделя",{p}*{q}*52,'
        f'IF({t}="месяц",{p}*{q}*12,{p}*{q})))'
    )


def write_rows(ws, start, block):
    items = block.get("items") or []
    headers, cmap = layout(block)
    for c, h in enumerate(headers, 1):
        cell = ws.cell(start, c, h)
        cell.fill = fill_navy
        cell.font = font_h
        cell.alignment = center
        cell.border = thin
    link_font = Font(name="Calibri", size=10, color="0563C1", underline="single")
    rows_out = []
    for i, item in enumerate(items, 1):
        r = start + i
        ws.cell(r, 1, i).alignment = center
        ws.cell(r, 1).border = thin
        ws.cell(r, 2, item.get("name") or "").alignment = left
        ws.cell(r, 2).border = thin
        if "pcs" in cmap:
            style_input(ws.cell(r, cmap["pcs"], item.get("pcs") or 1))
        if "url" in cmap:
            url = item.get("url") or ""
            cell = ws.cell(r, cmap["url"], url)
            style_input(cell)
            cell.alignment = left
            if url.startswith("http"):
                cell.hyperlink = url
                cell.font = link_font
                cell.fill = fill_input
        price = ws.cell(r, cmap["price"], item.get("price") or 0)
        price.number_format = money_fmt
        style_input(price)
        if float(item.get("years") or 0):
            term_val = item.get("years") or 1
        else:
            term_val = item.get("period") or "месяц"
        style_input(ws.cell(r, cmap["term"], term_val))
        cy = ws.cell(r, cmap["year"], year_formula(cmap, r))
        cy.number_format = money_fmt
        cy.border = thin
        cy.alignment = right
        yletter = col_letter(cmap["year"])
        cm = ws.cell(r, cmap["month"], f"=IF({yletter}{r}=0,0,{yletter}{r}/12)")
        cm.number_format = money_fmt
        cm.border = thin
        cm.alignment = right
        ws.row_dimensions[r].height = 20
        rows_out.append(r)
    last = start + max(len(items), 0)
    return last, rows_out, cmap


def sheet_suffix(block):
    sec = str(block.get("section") or "")
    if sec in ("health", "school"):
        return ""
    if sec.endswith("buy"):
        return " покупки"
    return " расходы"


def write_block_sheet(wb, block):
    raw = (block["title"] + sheet_suffix(block)).replace(":", " —").replace("/", "-")
    base = raw[:31]
    title = base
    n = 2
    while title in wb.sheetnames:
        title = (base[:28] + f" {n}")[:31]
        n += 1
    ws = wb.create_sheet(title)
    ws.sheet_properties.tabColor = NAVY
    headers, cmap = layout(block)
    last_col = col_letter(len(headers))
    ws["A1"] = block["title"]
    ws["A1"].font = font_title
    ws.merge_cells(f"A1:{last_col}1")
    ws["A2"] = "Покупки — ссылка на товар и срок службы. Частое — период. Платежи — без ссылки."
    ws["A2"].font = font_sub
    ws.merge_cells(f"A2:{last_col}2")

    last, rows, cmap = write_rows(ws, 4, block)
    yletter = col_letter(cmap["year"])
    mletter = col_letter(cmap["month"])
    year_cells = [f"{yletter}{x}" for x in rows]
    ncols = len(headers)
    merge_end = cmap["term"]
    items = block.get("items") or []
    has_years = any(float(it.get("years") or 0) for it in items)
    has_period = any(it.get("period") for it in items)
    if items and has_period and not has_years:
        dv = DataValidation(type="list", formula1='"неделя,месяц,год"', allow_blank=False)
        tcol = col_letter(cmap["term"])
        dv.add(f"{tcol}5:{tcol}{last}")
        ws.add_data_validation(dv)

    tot = last + 2
    shared = str(block.get("section") or "").startswith("shared")
    ws.cell(tot, 1, "Итого семья" if shared else f"Итого: {block['title']}")
    ws.merge_cells(start_row=tot, start_column=1, end_row=tot, end_column=merge_end)
    for c in range(1, ncols + 1):
        ws.cell(tot, c).fill = fill_sum
        ws.cell(tot, c).font = font_sum
        ws.cell(tot, c).border = thin
    formula = "+".join(year_cells) if year_cells else "0"
    ws.cell(tot, cmap["year"], f"={formula}" if year_cells else 0)
    ws.cell(tot, cmap["month"], f"=IF({yletter}{tot}=0,0,{yletter}{tot}/12)")
    ws.cell(tot, cmap["year"]).number_format = money_fmt
    ws.cell(tot, cmap["month"]).number_format = money_fmt
    if shared:
        tot2 = tot + 1
        ws.cell(tot2, 1, "Итого на ребёнка (÷3)")
        ws.merge_cells(start_row=tot2, start_column=1, end_row=tot2, end_column=merge_end)
        for c in range(1, ncols + 1):
            ws.cell(tot2, c).fill = fill_sum
            ws.cell(tot2, c).font = font_sum
            ws.cell(tot2, c).border = thin
        ws.cell(tot2, cmap["year"], f"={yletter}{tot}/3")
        ws.cell(tot2, cmap["month"], f"={mletter}{tot}/3")
        ws.cell(tot2, cmap["year"]).number_format = money_fmt
        ws.cell(tot2, cmap["month"]).number_format = money_fmt
        tot = tot2

    widths = {col_letter(1): 5, col_letter(2): 40}
    if "pcs" in cmap:
        widths[col_letter(cmap["pcs"])] = 13
    if "url" in cmap:
        widths[col_letter(cmap["url"])] = 36
    widths[col_letter(cmap["price"])] = 14
    widths[col_letter(cmap["term"])] = 16
    widths[col_letter(cmap["year"])] = 14
    widths[col_letter(cmap["month"])] = 14
    for col, w in widths.items():
        ws.column_dimensions[col].width = w
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToPage = True
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    return tot, yletter, mletter, title


def archive_file(path, folder):
    path = Path(path)
    if not path.exists():
        return
    folder = Path(folder)
    folder.mkdir(parents=True, exist_ok=True)
    stamp = datetime.now().strftime("%Y-%m-%d_%H%M%S")
    dest = folder / f"{path.stem}_{stamp}{path.suffix}"
    shutil.copy2(path, dest)
    print("archive", dest)


def build_excel():
    wb = Workbook()
    ws = wb.active
    ws.title = "Как пользоваться"
    ws.sheet_properties.tabColor = GOLD
    ws["A1"] = "Как пользоваться"
    ws["A1"].font = font_title
    ws["A2"] = "Наименование, количество, стоимость. Дальше либо срок службы в годах, либо сколько раз в месяц."
    ws["A2"].font = font_sub
    howto = [
        ("Покупка", "Срок службы (лет). В год = стоимость × количество ÷ срок ÷ делим."),
        ("Продукты и секции", "Период: неделя / месяц / год. В год = стоимость × количество × 52 / 12 / 1."),
        ("Делим = 1", "Целиком на этого ребёнка."),
        ("Делим = 3", "Общее на семью."),
        ("Строки", "Внутри блока по алфавиту."),
    ]
    ws["A4"] = "Что"
    ws["B4"] = "Зачем"
    style_header(ws, 4, 2)
    for i, (a, b) in enumerate(howto, 5):
        ws.cell(i, 1, a).font = font_b
        ws.cell(i, 2, b).font = font_n
        ws.cell(i, 1).border = thin
        ws.cell(i, 2).border = thin
        ws.cell(i, 1).alignment = left
        ws.cell(i, 2).alignment = left
    ws.column_dimensions["A"].width = 24
    ws.column_dimensions["B"].width = 100

    tots = []
    for block in BLOCKS:
        tot, yletter, mletter, sheet_name = write_block_sheet(wb, block)
        tots.append((sheet_name, tot, yletter, mletter))

    gws = wb.create_sheet("Группы диаграммы", 1)
    gws.sheet_properties.tabColor = GOLD
    gws["A1"] = "Группы диаграммы"
    gws["A1"].font = font_title
    gws.merge_cells("A1:B1")
    gws["A2"] = "Расход на диаграмме и блоки, которые в него входят."
    gws["A2"].font = font_sub
    gws.merge_cells("A2:B2")
    gws.cell(4, 1, "Расход")
    gws.cell(4, 2, "Состав блоков")
    style_header(gws, 4, 2)
    by_id = {b["id"]: b for b in BLOCKS}
    grouped_ids = set()
    gr = 5
    for g in CHART_GROUPS:
        names = [by_id[i]["title"] for i in (g.get("blocks") or []) if i in by_id]
        gws.cell(gr, 1, g.get("name") or "").alignment = left
        gws.cell(gr, 2, "; ".join(names)).alignment = left
        style_input(gws.cell(gr, 1))
        style_input(gws.cell(gr, 2))
        gws.cell(gr, 1).alignment = left
        gws.cell(gr, 2).alignment = left
        grouped_ids.update(g.get("blocks") or [])
        gr += 1
    for b in BLOCKS:
        if b["id"] in grouped_ids:
            continue
        gws.cell(gr, 1, b["title"]).alignment = left
        gws.cell(gr, 2, b["title"]).alignment = left
        gws.cell(gr, 1).border = thin
        gws.cell(gr, 2).border = thin
        gr += 1
    gws.column_dimensions["A"].width = 36
    gws.column_dimensions["B"].width = 80

    ws = wb.create_sheet("Сводка", 1)
    ws.sheet_properties.tabColor = GOLD
    ws["A1"] = "Сводка — один ребёнок"
    ws["A1"].font = Font(name="Calibri", bold=True, color=NAVY, size=18)
    headers = ["Блок", "В год, ₽", "В месяц, ₽"]
    for c, h in enumerate(headers, 1):
        ws.cell(4, c, h)
    style_header(ws, 4, 3)
    for i, (name, tot, yletter, mletter) in enumerate(tots):
        r = 5 + i
        sheet = name[:31]
        ws.cell(r, 1, name).border = thin
        ws.cell(r, 2, f"='{sheet}'!{yletter}{tot}")
        ws.cell(r, 3, f"='{sheet}'!{mletter}{tot}")
        for c in (2, 3):
            ws.cell(r, c).number_format = money_fmt
            ws.cell(r, c).border = thin
            ws.cell(r, c).alignment = right
        if i % 2:
            ws.cell(r, 1).fill = fill_zebra
    last = 4 + len(tots)
    gt = last + 1
    ws.cell(gt, 1, "СОДЕРЖАНИЕ РЕБЁНКА")
    ws.cell(gt, 2, f"=SUM(B5:B{last})")
    ws.cell(gt, 3, f"=SUM(C5:C{last})")
    for c in range(1, 4):
        ws.cell(gt, c).fill = fill_sum
        ws.cell(gt, c).font = font_sum
        ws.cell(gt, c).border = thin
        if c > 1:
            ws.cell(gt, c).number_format = money_fmt
    ws.column_dimensions["A"].width = 40
    ws.column_dimensions["B"].width = 16
    ws.column_dimensions["C"].width = 16
    chart = BarChart()
    chart.type = "bar"
    chart.style = 10
    chart.title = "Блоки, в год"
    data = Reference(ws, min_col=2, min_row=4, max_row=last)
    cats = Reference(ws, min_col=1, min_row=5, max_row=last)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    chart.legend = None
    chart.height = 12
    chart.width = 18
    ws.add_chart(chart, "E4")

    GOTOV.mkdir(parents=True, exist_ok=True)
    out = GOTOV / "Byudzhet-rebenka.xlsx"
    archive_file(out, ARXIV / "excel")
    wb.save(out)
    print("saved", out)


def build_html():
    payload = {"blocks": BLOCKS, "tiles": TILES, "chart_groups": CHART_GROUPS, "parts": PARTS}
    tpl = (HERE / "budget_interactive.html").read_text(encoding="utf-8")
    html = tpl.replace("__DATA__", json.dumps(payload, ensure_ascii=False))
    GOTOV.mkdir(parents=True, exist_ok=True)
    out = GOTOV / "spisok-v-shkolu-detskiy-mir.html"
    archive_file(out, ARXIV / "html")
    out.write_text(html, encoding="utf-8")
    grand = sum(row_year(r) for b in BLOCKS for r in b["items"])
    school = sum(row_year(r) for b in BLOCKS if b.get("id") == "school_must" or b.get("school") for r in b["items"])
    print("html grand_m", round(grand / 12), "school_y", round(school))
    for b in BLOCKS:
        y = sum(row_year(r) for r in b["items"])
        print(f"  {b['title']:40} {y/12:8.0f}")


if __name__ == "__main__":
    build_excel()
    build_html()
