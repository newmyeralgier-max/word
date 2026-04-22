"""Нормализация маркированных/нумерованных списков."""
import re

from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, Cm

from . import config as cfg
from . import detect
from .utils import apply_p_format, set_run_font, p_text

_RE_DASH_PREFIX = re.compile(r'^\s*[-•]\s*')
_RE_EMDASH_PREFIX = re.compile(r'^\s*—\s*')


def format_list_item(p):
    # Нормализуем маркер в начале первого run
    first_r = next(iter(p.runs), None)
    if first_r and first_r.text:
        # «- » / «• » → «– »
        first_r.text = _RE_DASH_PREFIX.sub('– ', first_r.text)
        first_r.text = _RE_EMDASH_PREFIX.sub('– ', first_r.text)
    apply_p_format(
        p,
        align=WD_ALIGN_PARAGRAPH.JUSTIFY,
        first_line_ind=cfg.FIRST_LINE_INDENT,
        left_ind=Cm(0),
        right_ind=Cm(0),
        space_before=Pt(0),
        space_after=Pt(0),
        line_spacing=cfg.LINE_SPACING,
    )
    for r in p.runs:
        set_run_font(r, font_name=cfg.FONT_NAME, font_size=cfg.FONT_SIZE_MAIN,
                     color=cfg.COLOR_BLACK, force_size=True)
