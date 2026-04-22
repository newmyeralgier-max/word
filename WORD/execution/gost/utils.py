"""Общие утилиты для пайплайна gost."""
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from . import config as cfg


# --- Безопасный текст и стиль -----------------------------------------------

def p_text(p) -> str:
    try:
        return (p.text or '').strip()
    except Exception:
        return ''


def p_style_name(p) -> str:
    try:
        return p.style.name or ''
    except Exception:
        return ''


def has_math(paragraph) -> bool:
    return len(paragraph._element.xpath('.//m:oMath')) > 0


def has_image(paragraph) -> bool:
    return bool(paragraph._element.xpath('.//w:drawing')) or \
           bool(paragraph._element.xpath('.//w:pict'))


def is_empty(paragraph) -> bool:
    """Параграф считается пустым, если в нём нет ни текста, ни формулы,
    ни изображения."""
    if p_text(paragraph):
        return False
    if has_math(paragraph) or has_image(paragraph):
        return False
    return True


# --- Работа с runs ----------------------------------------------------------

MONO_FONTS = {'consolas', 'courier new', 'courier', 'monaco', 'menlo'}


def set_run_font(run, *, font_name=None, font_size=None, bold=None,
                 italic=None, color=None, force_size=False):
    """Аккуратная установка параметров run.

    * `font_size` применяется ТОЛЬКО если текущий размер уже задан, либо
      явно передан `force_size=True` — чтобы не ломать наследование от стиля.
    * моноширинный шрифт (код) сохраняется.
    * `color` всегда можно форсить — цветные остатки AI-подсветки сводим к
      чёрному.
    """
    if run is None:
        return
    current_name = (run.font.name or '').lower()
    is_code = current_name in MONO_FONTS

    if font_name and not is_code:
        run.font.name = font_name
        # eastAsia/cs — чтобы Word не переключал на Calibri
        rFonts = run._element.xpath('.//w:rFonts')
        for rf in rFonts:
            rf.set(qn('w:eastAsia'), font_name)
            rf.set(qn('w:cs'), font_name)
            rf.set(qn('w:hAnsi'), font_name)
    if font_size is not None:
        if force_size or run.font.size is not None:
            run.font.size = font_size
    if bold is not None:
        run.bold = bold
    if italic is not None:
        run.italic = italic
    if color is not None:
        run.font.color.rgb = color


def apply_p_format(p, *, align=None, left_ind=Cm(0), right_ind=Cm(0),
                   first_line_ind=Cm(0), space_before=Pt(0), space_after=Pt(0),
                   line_spacing=cfg.LINE_SPACING, keep_next=False,
                   widow_control=True, page_break_before=None):
    pf = p.paragraph_format
    if align is not None:
        p.alignment = align
    pf.left_indent = left_ind
    pf.right_indent = right_ind
    pf.first_line_indent = first_line_ind
    pf.space_before = space_before
    pf.space_after = space_after
    pf.line_spacing = line_spacing
    pf.keep_with_next = keep_next
    pf.widow_control = widow_control
    if page_break_before is not None:
        pf.page_break_before = page_break_before


# --- Page-break helpers -----------------------------------------------------

def has_page_break_before(p) -> bool:
    pPr = p._element.find(qn('w:pPr'))
    if pPr is not None and pPr.find(qn('w:pageBreakBefore')) is not None:
        return True
    prev = p._element.getprevious()
    while prev is not None:
        for br in prev.findall('.//' + qn('w:br')):
            if br.get(qn('w:type')) == 'page':
                return True
        texts = prev.findall('.//' + qn('w:t'))
        if ''.join((t.text or '') for t in texts).strip():
            break
        prev = prev.getprevious()
    return False


def insert_page_break_before(p):
    if not has_page_break_before(p):
        p.paragraph_format.page_break_before = True


def iter_body_paragraphs(doc):
    """Генератор всех параграфов документа в порядке появления."""
    for p in doc.paragraphs:
        yield p


def paragraph_has_text_upper(p, needle: str) -> bool:
    return needle.strip().upper() in p_text(p).upper()


def strip_trailing_dot(text: str) -> str:
    text = text.rstrip()
    while text.endswith(('.', '．')):
        text = text[:-1].rstrip()
    return text
