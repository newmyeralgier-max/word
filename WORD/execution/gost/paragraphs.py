"""Форматирование обычных параграфов основного текста."""
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm

from . import config as cfg
from . import detect
from .utils import apply_p_format, set_run_font, is_empty


def format_body(p):
    apply_p_format(
        p,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        left_ind=Cm(0),
        right_ind=Cm(0),
        first_line_ind=cfg.FIRST_LINE_INDENT,
        space_before=Pt(0),
        space_after=Pt(0),
        line_spacing=cfg.LINE_SPACING,
        keep_next=False,
    )
    for r in p.runs:
        set_run_font(r, font_name=cfg.FONT_NAME, font_size=cfg.FONT_SIZE_MAIN,
                     color=cfg.COLOR_BLACK, force_size=True)


def format_empty(p):
    """Пустой абзац — поджимаем межстрочный до 1.0, чтобы не раздувал страницу."""
    apply_p_format(
        p,
        align=WD_ALIGN_PARAGRAPH.LEFT,
        first_line_ind=Cm(0),
        line_spacing=cfg.LINE_SPACING_TIGHT,
    )


def format_caption(p, *, align=WD_ALIGN_PARAGRAPH.CENTER):
    apply_p_format(
        p,
        align=align,
        first_line_ind=Cm(0),
        left_ind=Cm(0),
        right_ind=Cm(0),
        space_before=Pt(6),
        space_after=Pt(6),
        line_spacing=cfg.LINE_SPACING,
        keep_next=False,
    )
    for r in p.runs:
        set_run_font(r, font_name=cfg.FONT_NAME, font_size=cfg.FONT_SIZE_CAP,
                     color=cfg.COLOR_BLACK, bold=False, force_size=True)
