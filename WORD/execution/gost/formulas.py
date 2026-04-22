"""Форматирование формул: центрирование, tab-stop для номера."""
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.shared import Pt, Cm, Mm

from . import config as cfg
from . import detect
from .utils import apply_p_format, set_run_font


def format_formula(p):
    apply_p_format(
        p,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_ind=Cm(0),
        left_ind=Cm(0),
        right_ind=Cm(0),
        space_before=Pt(6),
        space_after=Pt(6),
        line_spacing=cfg.LINE_SPACING,
        keep_next=False,
    )
    tab_stops = p.paragraph_format.tab_stops
    tab_stops.clear_all()
    right = cfg.TEXT_WIDTH
    tab_stops.add_tab_stop(right, WD_TAB_ALIGNMENT.RIGHT)
    for r in p.runs:
        set_run_font(r, font_name=cfg.FONT_NAME, font_size=cfg.FONT_SIZE_MAIN,
                     color=cfg.COLOR_BLACK, force_size=True)


def format_where(p):
    apply_p_format(
        p,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_ind=Cm(0),
        left_ind=Cm(0),
        right_ind=Cm(0),
        space_before=Pt(0),
        space_after=Pt(0),
        line_spacing=cfg.LINE_SPACING,
    )
    for r in p.runs:
        set_run_font(r, font_name=cfg.FONT_NAME, font_size=cfg.FONT_SIZE_MAIN,
                     color=cfg.COLOR_BLACK, force_size=True)
