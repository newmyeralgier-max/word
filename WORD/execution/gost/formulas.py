"""Форматирование формул по ГОСТ.

Главная проблема, которую решает модуль, — «уехавший вправо-вниз» номер
формулы вида `(2.3)`. В исходном БР такие формулы лежат в одном абзаце,
в последовательности: [OMML/OLE-объект формулы] [запятая] [несколько
`<w:tab/>`] [`(2.3)`]. Если такой абзац центрировать, tab-stop-ов нет, и
длинный OMML + запятая не влезают в строку — номер переносится вниз.

Что делаем:
* таб-стопы переводим в строго два положения:  центр (TEXT_WIDTH/2) и правый
  (TEXT_WIDTH). Абзац — `jc=left`, чтобы tab-ы работали.
* ПЕРЕД формулой вставляем ведущий tab (чтобы формула оказалась у центра-
  tab-а), ПОСЛЕ формулы — один tab к правому tab-у.
* все лишние `<w:tab/>` удаляем.
* если абзац — «голая» формула без номера, просто центрируем.
"""
import re

from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from docx.shared import Pt, Cm

from . import config as cfg
from .utils import apply_p_format, set_run_font, p_text


_RE_TAIL_NUMBER = re.compile(r'\(\s*(\d+(?:\.\d+)*)\s*\)\s*\.?\s*$')


def _set_tab_stops(p, positions):
    """Полностью переопределить таб-стопы в pPr."""
    pPr = p._element.get_or_add_pPr()
    for old in list(pPr.findall(qn('w:tabs'))):
        pPr.remove(old)
    tabs = OxmlElement('w:tabs')
    for pos_emu, align in positions:
        tab = OxmlElement('w:tab')
        tab.set(qn('w:val'), align)
        tab.set(qn('w:pos'), str(int(pos_emu / 635)))  # EMU -> twips (EMU=635 per twip)
        tabs.append(tab)
    pPr.append(tabs)


def format_formula(p):
    """Форматировать абзац-формулу.

    Если в конце текста обнаружен номер `(X.Y)`, выставляем разметку с тремя
    зонами (лев-центр-право). Иначе просто центрируем.
    """
    text = p_text(p)
    m_tail = _RE_TAIL_NUMBER.search(text)
    has_number = bool(m_tail)

    if has_number:
        apply_p_format(
            p,
            align=WD_ALIGN_PARAGRAPH.LEFT,
            first_line_ind=Cm(0),
            left_ind=Cm(0),
            right_ind=Cm(0),
            space_before=Pt(6),
            space_after=Pt(6),
            line_spacing=cfg.LINE_SPACING,
        )
        # Таб-стопы: центр (8,25 см) и правый (16,5 см)
        center_pos = cfg.TEXT_WIDTH // 2
        right_pos = cfg.TEXT_WIDTH
        # python-docx на Mm() возвращает EMU. Наш помощник ждёт EMU.
        _set_tab_stops(p, [
            (int(center_pos), 'center'),
            (int(right_pos), 'right'),
        ])
        _rebuild_tabs_for_numbered_formula(p)
    else:
        apply_p_format(
            p,
            align=WD_ALIGN_PARAGRAPH.CENTER,
            first_line_ind=Cm(0),
            left_ind=Cm(0),
            right_ind=Cm(0),
            space_before=Pt(6),
            space_after=Pt(6),
            line_spacing=cfg.LINE_SPACING,
        )
        # Удалим лишние <w:tab/>, чтобы они не добавляли хаотичные отступы
        _drop_all_tabs(p)

    for r in p.runs:
        set_run_font(r, font_name=cfg.FONT_NAME, font_size=cfg.FONT_SIZE_MAIN,
                     color=cfg.COLOR_BLACK, force_size=False)


def _drop_all_tabs(p):
    for tab in list(p._element.iter(qn('w:tab'))):
        # оставить только <w:tab> внутри <w:tabs> (это таб-стопы, а не символы)
        parent = tab.getparent()
        if parent.tag == qn('w:tabs'):
            continue
        parent.remove(tab)


def _rebuild_tabs_for_numbered_formula(p):
    """В абзаце с формулой и номером вида `(X.Y)` выставить один leading-tab
    перед формулой и один trailing-tab перед номером.

    Мы не пытаемся угадывать, где в run-ах именно номер: вместо этого
    *нормализуем* последовательность `w:tab` символов внутри run-ов:
    удаляем все, затем вставляем  один `w:tab` как самый первый дочерний
    элемент первого run-а, и ещё один — прямо перед run-ом, содержащим
    `(N.M)`.
    """
    p_el = p._element
    # 1) Удалить все символьные <w:tab/>
    for tab in list(p_el.iter(qn('w:tab'))):
        parent = tab.getparent()
        if parent.tag == qn('w:tabs'):
            continue  # это из pPr
        parent.remove(tab)

    # 2) Найти run с '(N.M)' в конце
    runs = list(p_el.findall(qn('w:r')))
    if not runs:
        return
    number_run = None
    for r in reversed(runs):
        t_elements = r.findall(qn('w:t'))
        if not t_elements:
            continue
        joined = ''.join((t.text or '') for t in t_elements)
        if _RE_TAIL_NUMBER.search(joined):
            number_run = r
            break

    # 3) Вставить <w:tab/> внутрь первого run-а — как самый первый элемент.
    first_run = runs[0]
    tab_lead = OxmlElement('w:tab')
    # Вставим перед <w:t>/<m:oMath>/<w:object> — т.е. как первый дочерний
    # после <w:rPr> (если он есть)
    rPr = first_run.find(qn('w:rPr'))
    if rPr is not None:
        rPr.addnext(tab_lead)
    else:
        first_run.insert(0, tab_lead)

    # 4) Вставить <w:tab/> непосредственно перед `number_run` — как
    # самостоятельный run (так безопаснее по XML).
    if number_run is not None and number_run is not first_run:
        tab_run = OxmlElement('w:r')
        tab_inside = OxmlElement('w:tab')
        tab_run.append(tab_inside)
        number_run.addprevious(tab_run)


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
                     color=cfg.COLOR_BLACK, force_size=False)
