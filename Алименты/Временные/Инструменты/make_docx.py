# -*- coding: utf-8 -*-
"""Утилита: сборка docx-документов из простых блоков.
Использование:
    from make_docx import make_doc
    make_doc('путь.docx', blocks)
blocks = список кортежей (kind, text), kind:
    'center' - жирный по центру; 'right' - справа; 'b' - жирный по ширине;
    'plain' - обычный по ширине; 'gap' - пустая строка.
Шрифт: Times New Roman 14.
"""
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

def make_doc(path, blocks):
    doc = Document()
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(14)
    for kind, text in blocks:
        p = doc.add_paragraph()
        if kind == 'center':
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(text); r.bold = True
        elif kind == 'right':
            p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p.add_run(text)
        elif kind == 'b':
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            r = p.add_run(text); r.bold = True
        elif kind == 'plain':
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.add_run(text)
        elif kind == 'gap':
            p.add_run('')
    doc.save(path)
    return path
