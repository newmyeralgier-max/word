"""Форматирование таблиц: граница, центрирование на странице, ширина,
шрифт 12pt, единичный интервал, повторяющаяся шапка."""
from copy import deepcopy

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ROW_HEIGHT_RULE
from docx.shared import Pt, Cm, Mm, Emu
from docx.oxml.ns import qn, nsdecls
from docx.oxml import OxmlElement, parse_xml

from . import config as cfg
from .utils import apply_p_format, set_run_font, has_math


def _set_table_borders(table):
    tbl = table._element
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    # Удаляем старые tblBorders
    for b in tblPr.findall(qn('w:tblBorders')):
        tblPr.remove(b)
    borders_xml = f'''<w:tblBorders {nsdecls("w")}>
        <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>
        <w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>
    </w:tblBorders>'''
    tblPr.append(parse_xml(borders_xml))


def _center_table(table):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl = table._element
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = OxmlElement('w:tblPr')
        tbl.insert(0, tblPr)
    # Обновляем/ставим jc=center
    for jc in tblPr.findall(qn('w:jc')):
        tblPr.remove(jc)
    jc = OxmlElement('w:jc')
    jc.set(qn('w:val'), 'center')
    tblPr.append(jc)
    # Сбрасываем tblInd (отступ таблицы), иначе из-за него может смещаться
    for ind in tblPr.findall(qn('w:tblInd')):
        tblPr.remove(ind)
    ind = OxmlElement('w:tblInd')
    ind.set(qn('w:w'), '0')
    ind.set(qn('w:type'), 'dxa')
    tblPr.append(ind)


def _shrink_table_width(table):
    """Если суммарная ширина колонок превышает ширину текста (16.5 см),
    пересчитываем ширины колонок пропорционально, чтобы влезли."""
    total_emu = 0
    widths = []
    tblGrid = table._element.find(qn('w:tblGrid'))
    if tblGrid is None:
        return
    for g in tblGrid.findall(qn('w:gridCol')):
        w = g.get(qn('w:w'))
        if w is None:
            widths.append(0)
            continue
        widths.append(int(w))
        total_emu += int(w)
    if total_emu <= 0:
        return
    # Word хранит ширину в twips (1/20 pt) — 1 см = 567 twips
    max_twips = int(cfg.TEXT_WIDTH.emu / 635)  # 1 emu = 1/635 twip приблизительно
    # Правильный пересчёт: Cm(16.5) в twips = 16.5 * 567 ≈ 9355
    max_twips = int(16.5 * 567)
    if total_emu <= max_twips:
        return
    scale = max_twips / total_emu
    new_widths = [max(600, int(w * scale)) for w in widths]
    for g, nw in zip(tblGrid.findall(qn('w:gridCol')), new_widths):
        g.set(qn('w:w'), str(nw))
    # Также пройтись по ячейкам и обновить их tcW
    for row in table.rows:
        for cell, nw in zip(row.cells, new_widths):
            tcPr = cell._element.get_or_add_tcPr()
            for tcW in tcPr.findall(qn('w:tcW')):
                tcPr.remove(tcW)
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:w'), str(nw))
            tcW.set(qn('w:type'), 'dxa')
            tcPr.append(tcW)
    # И у tblW (общая ширина) выставить auto
    tblPr = table._element.find(qn('w:tblPr'))
    if tblPr is not None:
        for tblW in tblPr.findall(qn('w:tblW')):
            tblPr.remove(tblW)
        tblW = OxmlElement('w:tblW')
        tblW.set(qn('w:w'), str(sum(new_widths)))
        tblW.set(qn('w:type'), 'dxa')
        tblPr.append(tblW)


def _repeat_header(table):
    if not table.rows:
        return
    tr = table.rows[0]._tr
    trPr = tr.find(qn('w:trPr'))
    if trPr is None:
        trPr = OxmlElement('w:trPr')
        tr.insert(0, trPr)
    # Убрать старый tblHeader
    for th in trPr.findall(qn('w:tblHeader')):
        trPr.remove(th)
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))


def format_table(table):
    _set_table_borders(table)
    _center_table(table)
    _shrink_table_width(table)
    _repeat_header(table)
    for row_idx, row in enumerate(table.rows):
        for cell in row.cells:
            for p in cell.paragraphs:
                if not (p.text.strip() or has_math(p)):
                    # пустой абзац в ячейке
                    apply_p_format(p, first_line_ind=Cm(0),
                                   line_spacing=cfg.LINE_SPACING_TIGHT)
                    continue
                apply_p_format(
                    p,
                    align=WD_ALIGN_PARAGRAPH.CENTER if row_idx == 0 else WD_ALIGN_PARAGRAPH.LEFT,
                    first_line_ind=Cm(0),
                    left_ind=Cm(0),
                    right_ind=Cm(0),
                    space_before=Pt(0),
                    space_after=Pt(0),
                    line_spacing=cfg.LINE_SPACING_TIGHT,
                )
                for r in p.runs:
                    set_run_font(r, font_name=cfg.FONT_NAME,
                                 font_size=cfg.FONT_SIZE_TABLE,
                                 color=cfg.COLOR_BLACK, force_size=True,
                                 bold=(True if row_idx == 0 else None))


def format_table_caption(p):
    """Подпись таблицы: «Таблица X.Y — Название», слева, без отступа."""
    apply_p_format(
        p,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_ind=Cm(0),
        left_ind=Cm(0),
        right_ind=Cm(0),
        space_before=Pt(6),
        space_after=Pt(0),
        line_spacing=cfg.LINE_SPACING,
        keep_next=True,
    )
    for r in p.runs:
        set_run_font(r, font_name=cfg.FONT_NAME, font_size=cfg.FONT_SIZE_MAIN,
                     color=cfg.COLOR_BLACK, bold=False, force_size=True)
