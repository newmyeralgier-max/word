"""
align_formulas.py — ГОСТ-выравнивание формул и их номеров.

Правило: формула по центру строки, номер (N.M) у правого края на той же
строке. Реализация: параграф выравнивается по левому краю, в pPr прописаны
две табуляции — центр (~4844 twips) и правый край (~9689 twips). В параграф
вставляется последовательность `<TAB>[formula]<TAB>(N.M)`.

Что обрабатываем:
  1. Параграф содержит <m:oMath> и в тексте его же рана есть "(N.M)"
     → переставляем ран с tab-ами так, чтобы tab был ДО oMath, один tab
       между oMath и номером, номер в финальном t-узле.
  2. Параграф содержит <m:oMath> без номера, а следующий параграф состоит
     только из "(N.M)" — сливаем в один параграф, добиваем tab-ами.

Формулы без номера не трогаем (бывают промежуточные выкладки).
"""
from __future__ import annotations
import re
from copy import deepcopy
from typing import Dict

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap

_M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
_M_OMATH = "{%s}oMath" % _M_NS
_NUM_RE = re.compile(r"\(\s*(\d+)\s*[\.,]\s*(\d+)\s*\)")

_TAB_CENTER_POS = "4844"
_TAB_RIGHT_POS = "9689"


def _text(p_elem) -> str:
    return "".join((t.text or "") for t in p_elem.findall(".//" + qn("w:t")))


def _ensure_pPr(p):
    pPr = p.find(qn("w:pPr"))
    if pPr is None:
        pPr = OxmlElement("w:pPr")
        p.insert(0, pPr)
    return pPr


def _set_tabs_and_jc_left(p):
    """Поставить табуляции center/right и выравнивание по левому краю."""
    pPr = _ensure_pPr(p)

    # убираем старые w:tabs
    old_tabs = pPr.find(qn("w:tabs"))
    if old_tabs is not None:
        pPr.remove(old_tabs)

    tabs = OxmlElement("w:tabs")
    t1 = OxmlElement("w:tab")
    t1.set(qn("w:val"), "center")
    t1.set(qn("w:pos"), _TAB_CENTER_POS)
    t2 = OxmlElement("w:tab")
    t2.set(qn("w:val"), "right")
    t2.set(qn("w:pos"), _TAB_RIGHT_POS)
    tabs.append(t1)
    tabs.append(t2)

    # w:tabs должен идти после keepNext/widowControl, но достаточно вставить
    # его в pPr — Word обычно порядок валидирует мягко. Ставим сразу после
    # первого элемента или в конец.
    pPr.append(tabs)

    # выравнивание по левому краю
    jc = pPr.find(qn("w:jc"))
    if jc is None:
        jc = OxmlElement("w:jc")
        pPr.append(jc)
    jc.set(qn("w:val"), "left")

    # убираем первую строку-отступ, чтобы tab не «съедался» красной строкой
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        ind = OxmlElement("w:ind")
        pPr.append(ind)
    ind.set(qn("w:firstLine"), "0")
    ind.set(qn("w:left"), "0")
    ind.set(qn("w:right"), "0")


def _make_tab_run(rpr_template=None):
    r = OxmlElement("w:r")
    if rpr_template is not None:
        r.append(deepcopy(rpr_template))
    tab = OxmlElement("w:tab")
    r.append(tab)
    return r


def _make_text_run(text: str, rpr_template=None):
    r = OxmlElement("w:r")
    if rpr_template is not None:
        r.append(deepcopy(rpr_template))
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    return r


def _first_rpr_in_paragraph(p):
    """Берём w:rPr первого w:r для переиспользования шрифта/размера."""
    for r in p.findall(qn("w:r")):
        rpr = r.find(qn("w:rPr"))
        if rpr is not None:
            return rpr
    return None


def _remove_all_tabs_and_find_number(p) -> str | None:
    """Удалить все w:tab-элементы и вернуть текст номера, если он был."""
    number = None
    for r in list(p.findall(qn("w:r"))):
        # собрать текст
        texts = r.findall(qn("w:t"))
        joined = "".join((t.text or "") for t in texts)
        m = _NUM_RE.search(joined)
        if m:
            number = m.group(0)
        # если ран содержит только tabs и число — удалим его
        tabs = r.findall(qn("w:tab"))
        breaks = r.findall(qn("w:br"))
        if tabs and not texts:
            p.remove(r)
            continue
        if m and re.fullmatch(r"\s*\(\s*\d+\s*[.,]\s*\d+\s*\)\s*", joined):
            # ран содержит только номер → удалим его (добавим заново позже)
            p.remove(r)
            continue
        # если ран содержит и текст и tabs — чистим tabs
        for t in tabs:
            r.remove(t)
        for b in breaks:
            r.remove(b)
    return number


def _insert_tab_before_math(p, rpr_template):
    """Вставить ран с <w:tab/> перед первым m:oMath/m:oMathPara."""
    # find first oMath element
    first_math = None
    for child in p.iter():
        if child.tag == _M_OMATH or child.tag == "{%s}oMathPara" % _M_NS:
            first_math = child
            break
    if first_math is None:
        return False
    # oMath может лежать прямо в p; найдём ancestor-ребёнка p
    target = first_math
    while target.getparent() is not p and target.getparent() is not None:
        target = target.getparent()
    if target.getparent() is not p:
        return False
    tab_run = _make_tab_run(rpr_template)
    p.insert(list(p).index(target), tab_run)
    return True


def _append_tab_and_number(p, number: str, rpr_template):
    """Добавить в конец параграфа: ран с tab + ран с '(N.M)'."""
    p.append(_make_tab_run(rpr_template))
    p.append(_make_text_run(number, rpr_template))


def align(body_in: str, body_out: str) -> Dict[str, int]:
    doc = Document(body_in)
    body = doc.element.body

    paras = list(body.iter(qn("w:p")))
    stats = {
        "formulas_total": 0,
        "same_para_number": 0,
        "merged_next_para": 0,
        "skipped_no_number": 0,
    }

    i = 0
    while i < len(paras):
        p = paras[i]
        if p.find(".//" + _M_OMATH) is None:
            i += 1
            continue
        stats["formulas_total"] += 1

        own_text = _text(p)
        own_num = _NUM_RE.search(own_text)

        number = None
        if own_num:
            number = own_num.group(0)
            case = "own"
        else:
            # смотрим следующий параграф: состоит ли ОН ЦЕЛИКОМ из "(N.M)"
            # и НЕ содержит собственной формулы (иначе удаление съест oMath)
            if i + 1 < len(paras):
                nxt = paras[i + 1]
                nxt_text = _text(nxt).strip()
                nxt_has_math = nxt.find(".//" + _M_OMATH) is not None
                if _NUM_RE.fullmatch(nxt_text) and not nxt_has_math:
                    number = nxt_text
                    case = "next"
                else:
                    case = None
            else:
                case = None

        if number is None:
            stats["skipped_no_number"] += 1
            i += 1
            continue

        rpr_template = _first_rpr_in_paragraph(p)

        _set_tabs_and_jc_left(p)

        if case == "own":
            _remove_all_tabs_and_find_number(p)
            _insert_tab_before_math(p, rpr_template)
            _append_tab_and_number(p, number, rpr_template)
            stats["same_para_number"] += 1
        else:  # case == "next"
            _remove_all_tabs_and_find_number(p)
            _insert_tab_before_math(p, rpr_template)
            _append_tab_and_number(p, number, rpr_template)
            # удалить следующий параграф с номером
            nxt = paras[i + 1]
            nxt.getparent().remove(nxt)
            stats["merged_next_para"] += 1

        i += 1

    doc.save(body_out)
    return stats


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: align_formulas.py IN.docx OUT.docx")
        sys.exit(1)
    print(align(sys.argv[1], sys.argv[2]))
