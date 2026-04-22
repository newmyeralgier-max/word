"""Классификаторы параграфов: заголовок, формула, подпись таблицы и т.п."""
import re

from docx.oxml.ns import qn

from . import config as cfg
from .utils import p_text, p_style_name, has_math, has_image, is_empty


_RE_FIGURE = re.compile(
    r'^\s*(Рис\.|Рисунок)\s*(\d+(?:[.,]\d+)*(?:,?[а-яА-Я])?)\s*[.\s—–\-]*\s*(.*)',
    re.I,
)
_RE_TABLE_CAP = re.compile(
    r'^\s*(Табл\.|Таблица)\s*(\d+(?:\.\d+)?)\.?\s*[.\s—–\-]*\s*(.*)',
    re.I,
)
_RE_TABLE_NUM_ONLY = re.compile(
    r'^\s*(Табл\.|Таблица)\s*(\d+(?:\.\d+)?)\s*\.?\s*$', re.I,
)
_RE_FORMULA_NUM = re.compile(r'\(\s*\d+(?:\.\d+)*\s*\)\s*$')
_RE_WHERE = re.compile(r'^где\s+', re.I)
_RE_LIST_DASH = re.compile(r'^\s*[-—–•]\s+')
_RE_LIST_NUM_CYRIL = re.compile(r'^\s*[а-яё]\)\s+')
_RE_LIST_NUM_DIGIT = re.compile(r'^\s*\d+[.)]\s+')


# --- Заголовки ---------------------------------------------------------------

def is_structural_element(p) -> bool:
    """ВВЕДЕНИЕ / ЗАКЛЮЧЕНИЕ / СПИСОК ЛИТЕРАТУРЫ и т.п. — по тексту."""
    t = p_text(p).upper().strip().strip('.').strip()
    if not t:
        return False
    for needle in cfg.STRUCTURAL_ELEMENTS:
        if t == needle or t.startswith(needle + ' ') or t.startswith(needle + '.'):
            return True
    # «ПРИЛОЖЕНИЕ А», «ПРИЛОЖЕНИЕ Б»
    if re.match(r'^ПРИЛОЖЕНИЕ(\s+[А-ЯA-Z])?$', t):
        return True
    return False


def is_style_heading(p) -> int:
    """Уровень заголовка по имени встроенного стиля, 0 если не заголовок."""
    name = p_style_name(p)
    if not name:
        return 0
    low = name.lower()
    if 'heading' in low or 'заголовок' in low:
        m = re.search(r'(\d+)', name)
        return int(m.group(1)) if m else 1
    # Специальные стили, которые в БР часто используют как заголовки
    if low in ('title', 'заголовок'):
        return 1
    return 0


def section_heading_level(text: str) -> int:
    """Возвращает уровень заголовка раздела, 0 — не заголовок."""
    first_line = text.split('\n')[0].strip()
    if not first_line:
        return 0
    # "1.ТЕКСТ" без пробела или "1. Текст" / "1.1 Текст" / "1.1. Текст"
    m = re.match(r'^(\d+(?:\.\d+){0,2})[.\s]+(.+)', first_line)
    if not m:
        m2 = re.match(r'^(\d+)\.([А-ЯЁA-Z].*)', first_line)
        if m2:
            m = m2
        else:
            return 0
    num, rest = m.group(1), m.group(2).strip()
    if num.count('.') >= 3:
        return 0
    if len(rest) < 4:
        return 0
    # Исключаем «1. Iном аппарата» (формулы/параметры)
    if re.match(r'^[A-Za-zА-Яа-я]{1,4}[ _]', rest) and re.search(r'\d', rest[:10]):
        return 0
    if re.search(r'[=≥≤<>]', rest):
        return 0
    # Если большая часть капсом — H1 (или H2 при наличии подномера)
    upper_ratio = sum(1 for c in rest if c.isupper()) / max(len(rest), 1)
    depth = num.count('.')
    if upper_ratio > 0.4:
        return max(1, min(depth + 1, 3))
    # Иначе — это H2/H3 только если есть подномер
    if depth >= 1 and rest[0:1].isupper() and len(rest) < 140:
        return min(depth + 1, 3)
    return 0


# --- Подписи к таблицам / рисункам ------------------------------------------

def match_figure_caption(text):
    return _RE_FIGURE.match(text)


def match_table_caption(text):
    return _RE_TABLE_CAP.match(text)


def match_table_num_only(text):
    return _RE_TABLE_NUM_ONLY.match(text)


# --- Формулы -----------------------------------------------------------------

def is_formula_paragraph(p) -> bool:
    if has_math(p):
        return True
    # OLE-equation (<w:object> с ProgID=Equation.*)
    objs = p._element.findall('.//' + qn_w('object'))
    if objs:
        for obj in objs:
            for ole in obj.iter():
                if ole.tag.endswith('}OLEObject') and (ole.get('ProgID') or '').startswith('Equation'):
                    return True
    # Текстовая формула-одиночка: заканчивается (N.M)
    t = p_text(p)
    if not t:
        return False
    if _RE_FORMULA_NUM.search(t) and len(t) < 220:
        return True
    # Короткий (<160 chars) параграф, похожий на уравнение: есть '=' и мало
    # кириллических «длинных» слов (> 3 букв). Примеры:
    #   "Рвк = Fвк Руд.кв = 1,2·3,4=4,08"
    #   "Руд = Руд.ул L=1,4·7,98 =11,18 кВт"
    if len(t) <= 160 and '=' in t and not t.endswith(':'):
        long_ru = re.findall(r'[а-яёА-ЯЁ]{4,}', t)
        # Слово-исключение «где …» — это не формула
        if re.match(r'^\s*(где|при|если|так)\b', t, re.I):
            return False
        # Заголовок/структурка уже ловится раньше в pipeline; здесь только
        # страхуемся по числам.
        if not re.search(r'\d', t):
            return False
        if len(long_ru) <= 1:
            return True
    return False


def qn_w(local: str) -> str:
    return qn('w:' + local)


def is_where_line(p) -> bool:
    t = p_text(p)
    return bool(_RE_WHERE.match(t))


# --- Списки ------------------------------------------------------------------

def list_marker_kind(text: str) -> str:
    """'dash' / 'alpha' / 'num' / ''."""
    if _RE_LIST_DASH.match(text):
        return 'dash'
    if _RE_LIST_NUM_CYRIL.match(text):
        return 'alpha'
    if _RE_LIST_NUM_DIGIT.match(text):
        return 'num'
    return ''


# --- Классификация высшего уровня ------------------------------------------

class PType:
    EMPTY = 'empty'
    TITLE_ZONE = 'title_zone'
    STRUCTURAL = 'structural'
    HEADING = 'heading'
    FIGURE_CAP = 'figure_caption'
    FIGURE_IMG = 'figure_image'
    TABLE_CAP = 'table_caption'
    FORMULA = 'formula'
    WHERE = 'where'
    LIST = 'list'
    BODY = 'body'
    TOC = 'toc'


def is_toc_paragraph(p) -> bool:
    name = p_style_name(p).lower()
    return name.startswith('toc') or name.startswith('содержание')
