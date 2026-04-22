"""Этап финальной сборки: body_03_gost.docx + титульник (впереди) → final.docx.

Стратегия: берём BODY как основу (в нём уже все картинки/формулы/таблицы/rels),
и ПРИПЕНДИМ в начало его <w:body> все дочерние элементы титульника.

Титульник у нас чисто текстовый (без картинок/таблиц/формул — проверено
в 01_extract_title.py), так что просто перенос XML-узлов работает корректно.

Использование:
    python stitch_title.py <body_gost.docx> <title.docx> <output.docx>
"""
from __future__ import annotations

import copy
import sys

from docx import Document
from docx.oxml.ns import qn


def stitch(body_path: str, title_path: str, out_path: str) -> None:
    body_doc = Document(body_path)
    title_doc = Document(title_path)
    bbody = body_doc.element.body
    tbody = title_doc.element.body

    # Первый ребёнок body_doc — чтобы перед ним вставить титульник
    first_child = None
    for c in bbody:
        if c.tag != qn("w:sectPr"):
            first_child = c
            break

    # Титульник: берём все элементы кроме sectPr
    title_children = [c for c in tbody if c.tag != qn("w:sectPr")]

    inserted = 0
    if first_child is not None:
        for t_child in title_children:
            new = copy.deepcopy(t_child)
            first_child.addprevious(new)
            inserted += 1
    else:
        # body пустой — просто добавляем
        sectPr = bbody.find(qn("w:sectPr"))
        for t_child in title_children:
            new = copy.deepcopy(t_child)
            if sectPr is not None:
                sectPr.addprevious(new)
            else:
                bbody.append(new)
            inserted += 1

    # Добавим разрыв страницы после последнего элемента титульника
    # (чтобы основной текст начинался с новой страницы)
    # Создаём параграф с page-break-before
    from docx.oxml import OxmlElement
    page_break_p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    p_before = OxmlElement("w:pageBreakBefore")
    pPr.append(p_before)
    page_break_p.append(pPr)
    if first_child is not None:
        first_child.addprevious(page_break_p)

    body_doc.save(out_path)
    print(f"[OK] Склейка сохранена: {out_path} (титульных элементов: {inserted})")


if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python stitch_title.py <body_gost.docx> <title.docx> <output.docx>")
        sys.exit(1)
    stitch(sys.argv[1], sys.argv[2], sys.argv[3])
