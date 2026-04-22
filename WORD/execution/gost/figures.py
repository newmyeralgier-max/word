"""Форматирование рисунков и подписей к ним."""
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm

from . import config as cfg
from . import detect
from .utils import apply_p_format, set_run_font, has_image


def format_figure_paragraph(p):
    """Параграф с рисунком — центрировать, без отступа."""
    apply_p_format(
        p,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_ind=Cm(0),
        left_ind=Cm(0),
        right_ind=Cm(0),
        space_before=Pt(6),
        space_after=Pt(0),
        line_spacing=cfg.LINE_SPACING,
        keep_next=True,
    )


def format_figure_caption(p):
    apply_p_format(
        p,
        align=WD_ALIGN_PARAGRAPH.CENTER,
        first_line_ind=Cm(0),
        left_ind=Cm(0),
        right_ind=Cm(0),
        space_before=Pt(0),
        space_after=Pt(12),
        line_spacing=cfg.LINE_SPACING,
    )
    for r in p.runs:
        set_run_font(r, font_name=cfg.FONT_NAME, font_size=cfg.FONT_SIZE_CAP,
                     color=cfg.COLOR_BLACK, bold=False, force_size=True)
