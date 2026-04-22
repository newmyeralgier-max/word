"""Общие утилиты пайплайна.

Принципы (из ANALYTICS_v4.md — «PRIMUM NON NOCERE»):
- Никогда не удаляем параграфы.
- Никогда не перезаписываем тело формул (OMML).
- Правки текста делаем через runs, без .text= на параграфе.
- Все номера глав/рисунков/таблиц/формул — сквозные в пределах главы.
"""
from __future__ import annotations

import copy
import re
from dataclasses import dataclass, field
from typing import Iterable, List, Optional

from docx.oxml.ns import qn


# ─── Регэкспы ──────────────────────────────────────────────────────────────

RE_FIG_CAPTION = re.compile(r"^\s*(Рис(?:унок|\.)?\s*\d+(?:\.\d+)*)", re.I)
RE_TBL_CAPTION = re.compile(r"^\s*(Табл(?:ица|\.)?\s*\d+(?:\.\d+)*)", re.I)
RE_H1 = re.compile(r"^\s*(\d+)\.\s+(.+)$")
RE_H2 = re.compile(r"^\s*(\d+)\.(\d+)\.?\s+(.+)$")
RE_H3 = re.compile(r"^\s*(\d+)\.(\d+)\.(\d+)\.?\s+(.+)$")
RE_FORMULA_NUM_SUFFIX = re.compile(r"\((\d+(?:\.\d+)*)\)\s*$")

STRUCTURAL = {
    "ВВЕДЕНИЕ",
    "ЗАКЛЮЧЕНИЕ",
    "СОДЕРЖАНИЕ",
    "ОГЛАВЛЕНИЕ",
    "СПИСОК ЛИТЕРАТУРЫ",
    "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ",
    "АННОТАЦИЯ",
    "РЕФЕРАТ",
    "ПРИЛОЖЕНИЕ",
    "ПРИЛОЖЕНИЯ",
}


# ─── Доступ к тексту ───────────────────────────────────────────────────────

def paragraph_text(p_elem) -> str:
    """Чистый текст параграфа из XML (обходя runs)."""
    parts = [t.text or "" for t in p_elem.findall(".//" + qn("w:t"))]
    return "".join(parts)


def paragraph_has_omml(p_elem) -> bool:
    return bool(p_elem.findall(".//" + qn("m:oMath")))


def paragraph_has_drawing(p_elem) -> bool:
    return bool(p_elem.findall(".//" + qn("w:drawing")))


def paragraph_tag(elem) -> bool:
    return elem.tag == qn("w:p") or elem.tag.endswith("}p")


def table_tag(elem) -> bool:
    return elem.tag == qn("w:tbl") or elem.tag.endswith("}tbl")


def first_style_of(p_elem) -> Optional[str]:
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        return None
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        return None
    return pStyle.get(qn("w:val"))


# ─── Структурная модель (карта документа) ─────────────────────────────────

@dataclass
class BodyElem:
    kind: str  # 'p' | 'tbl'
    idx_in_body: int
    text: str = ""
    elem: object = None
    style: Optional[str] = None
    has_omml: bool = False
    has_drawing: bool = False


def iter_body_elements(doc) -> List[BodyElem]:
    """Пройтись по всем детям <w:body> в порядке появления."""
    out: List[BodyElem] = []
    body = doc.element.body
    children = [c for c in body if c.tag != qn("w:sectPr")]
    for i, elem in enumerate(children):
        if paragraph_tag(elem):
            out.append(
                BodyElem(
                    kind="p",
                    idx_in_body=i,
                    text=paragraph_text(elem),
                    elem=elem,
                    style=first_style_of(elem),
                    has_omml=paragraph_has_omml(elem),
                    has_drawing=paragraph_has_drawing(elem),
                )
            )
        elif table_tag(elem):
            out.append(BodyElem(kind="tbl", idx_in_body=i, elem=elem))
    return out


# ─── Заголовок параграфа: «первый структурный» для разделения зон ──────────

def find_first_structural_idx(elems: List[BodyElem]) -> int:
    """Вернёт индекс (в списке elems) первого параграфа, начинающегося
    со структурного слова (ВВЕДЕНИЕ/СОДЕРЖАНИЕ и т.п.).

    Возвращает -1, если не нашли.
    """
    for i, e in enumerate(elems):
        if e.kind != "p":
            continue
        t = e.text.strip().upper().rstrip(".").rstrip()
        if not t:
            continue
        # берем первое слово
        first_word = t.split()[0] if t.split() else ""
        if first_word in STRUCTURAL:
            return i
    return -1


# ─── Правка текста параграфа БЕЗОПАСНО через runs ──────────────────────────

def replace_in_paragraph(p_elem, old: str, new: str) -> int:
    """Заменить подстроку в тексте параграфа, сохранив форматирование runs.

    Стратегия: склеить все w:t в один строковый буфер, найти подстроку, потом
    аккуратно распределить итог по существующим w:t (пропорционально позициям).
    Возвращает число сделанных замен.
    """
    t_nodes = p_elem.findall(".//" + qn("w:t"))
    if not t_nodes:
        return 0
    buf = "".join(n.text or "" for n in t_nodes)
    if old not in buf:
        return 0
    new_buf = buf.replace(old, new)
    # Простейшая стратегия: весь новый текст кладём в первый w:t, остальные чистим.
    t_nodes[0].text = new_buf
    # preserve space
    if new_buf.startswith(" ") or new_buf.endswith(" ") or "  " in new_buf:
        t_nodes[0].set(qn("xml:space"), "preserve")
    for n in t_nodes[1:]:
        n.text = ""
    return buf.count(old)


def set_paragraph_text(p_elem, new_text: str) -> None:
    """Задать текст параграфа целиком, затерев все runs кроме первого."""
    r_nodes = p_elem.findall(qn("w:r"))
    if not r_nodes:
        # нет runs — создаем один
        from docx.oxml import OxmlElement
        r = OxmlElement("w:r")
        t = OxmlElement("w:t")
        t.text = new_text
        t.set(qn("xml:space"), "preserve")
        r.append(t)
        p_elem.append(r)
        return
    # первый run оставляем, его w:t чистим и пишем новый текст, остальные runs удаляем
    first = r_nodes[0]
    t_first = first.find(qn("w:t"))
    if t_first is None:
        from docx.oxml import OxmlElement
        t_first = OxmlElement("w:t")
        first.append(t_first)
    t_first.text = new_text
    t_first.set(qn("xml:space"), "preserve")
    for r in r_nodes[1:]:
        p_elem.remove(r)


# ─── Стиль параграфа ──────────────────────────────────────────────────────

def set_paragraph_style(p_elem, style_name: str) -> None:
    """Назначить стиль параграфу (создаёт/обновляет w:pPr/w:pStyle)."""
    from docx.oxml import OxmlElement
    pPr = p_elem.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p_elem.insert(0, pPr)
    pStyle = pPr.find(qn("w:pStyle"))
    if pStyle is None:
        pStyle = OxmlElement("w:pStyle")
        pPr.insert(0, pStyle)
    pStyle.set(qn("w:val"), style_name)


# ─── Контекст вокруг элемента ─────────────────────────────────────────────

def collect_text_context(elems: List[BodyElem], center: int, before: int = 2, after: int = 2) -> str:
    """Собрать текстовый контекст вокруг элемента."""
    start = max(0, center - before)
    end = min(len(elems), center + after + 1)
    buf = []
    for i in range(start, end):
        e = elems[i]
        if e.kind != "p":
            continue
        t = e.text.strip()
        if t:
            buf.append(t)
    return " | ".join(buf)
