"""Форматирование заголовков H1/H2/H3 + структурных элементов ГОСТ."""
import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm
from docx.oxml.ns import qn

from . import config as cfg
from . import detect
from .utils import (apply_p_format, set_run_font, p_text, insert_page_break_before,
                     strip_trailing_dot)


def _size_for(level: int):
    return {1: cfg.FONT_SIZE_H1, 2: cfg.FONT_SIZE_H2, 3: cfg.FONT_SIZE_H3}.get(level, cfg.FONT_SIZE_H2)


def _bold_for(level: int):
    return {1: cfg.BOLD_H1, 2: cfg.BOLD_H2, 3: cfg.BOLD_H3}.get(level, False)


def format_heading(p, level: int, *, page_break: bool = False):
    apply_p_format(
        p,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_ind=Cm(0),
        space_before=Pt(18) if level == 1 else Pt(12),
        space_after=Pt(12) if level == 1 else Pt(6),
        line_spacing=cfg.LINE_SPACING,
        keep_next=True,
        page_break_before=page_break or None,
    )
    size = _size_for(level)
    bold = _bold_for(level)
    for r in p.runs:
        set_run_font(r, font_name=cfg.FONT_NAME, font_size=size, bold=bold,
                     italic=False, color=cfg.COLOR_BLACK, force_size=True)
    # Удаляем завершающую точку на уровне runs (если вдруг была)
    for r in reversed(p.runs):
        if r.text and r.text.rstrip():
            new = strip_trailing_dot(r.text)
            if new != r.text:
                r.text = new
            break


def format_structural(p):
    """ВВЕДЕНИЕ / ЗАКЛЮЧЕНИЕ / СПИСОК ЛИТЕРАТУРЫ ...

    Капсом, по центру, 16pt, со своей страницы, без нумерации, без точки.
    """
    t = p_text(p)
    t_up = strip_trailing_dot(t).upper()
    # Переписываем единственным run-ом, капсом
    # Но только если реально есть содержимое
    if not t_up:
        return
    # Сохраняем исходный rPr для шрифта
    from copy import deepcopy
    first_rPr = None
    for r in p.runs:
        rPr = r._element.find(qn('w:rPr'))
        if rPr is not None:
            first_rPr = deepcopy(rPr)
            break
    for r in list(p.runs):
        r._element.getparent().remove(r._element)
    new_run = p.add_run(t_up)
    if first_rPr is not None:
        new_run._element.insert(0, first_rPr)

    apply_p_format(
        p,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_ind=Cm(0),
        space_before=Pt(0),
        space_after=Pt(18),
        line_spacing=cfg.LINE_SPACING,
        keep_next=True,
        page_break_before=True,
    )
    for r in p.runs:
        set_run_font(r, font_name=cfg.FONT_NAME, font_size=cfg.FONT_SIZE_H1,
                     bold=False, italic=False, color=cfg.COLOR_BLACK,
                     force_size=True)
