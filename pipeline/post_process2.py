# -*- coding: utf-8 -*-
"""Post-process V2 (session 7): правки после ревью PR #4.

1. Перенос TOC из самого начала документа вниз — перед «ВВЕДЕНИЕ»
   (там где у пользователя оставлены «свободные страницы»).
2. Обратная переклейка: Таблица 3.1, 5.4, 6.1, 6.2, 7.1, 8.1, 9.1 —
   на самом деле рисунки → убираем caption «Таблица», добавляем
   «Рисунок N.M — …». Таблицы после них перенумеровываются, рисунки
   получают новый сквозной номер в главе.
3. Удаление мусора между Таблицей 5.8 и Таблицей 5.9 (дубль-caption
   «Таблица 4.» + «Погрешности опыта короткого замыкания…»).
4. Расширенная типография — добивка слипшегося текста.

Скрипт идемпотентный. Работает на текущем GOST-docx на диске.
"""
from __future__ import annotations

import re
import sys

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from common import replace_in_paragraph, set_paragraph_style, set_paragraph_text


# ─── 1. Перенос TOC к «ВВЕДЕНИЕ» ─────────────────────────────────────────

def _is_page_break_only(p) -> bool:
    """True, если параграф содержит только page-break и ничего больше."""
    brs = [b for b in p.findall(".//" + qn("w:br")) if b.get(qn("w:type")) == "page"]
    if not brs:
        return False
    txt = "".join((x.text or "") for x in p.findall(".//" + qn("w:t"))).strip()
    return txt == ""


def _is_toc_field_para(p) -> bool:
    for it in p.findall(".//" + qn("w:instrText")):
        if it.text and "TOC" in it.text:
            return True
    return False


def _is_toc_heading(p) -> bool:
    t = "".join((x.text or "") for x in p.findall(".//" + qn("w:t"))).strip()
    return t == "СОДЕРЖАНИЕ"


def move_toc_before_intro(body) -> dict:
    """Находит блок TOC (heading «СОДЕРЖАНИЕ» + TOC-field + опциональный
    page break) в ~первых 5 параграфах и переносит его прямо перед
    параграфом «ВВЕДЕНИЕ».
    Если вариант блока не найден — ничего не делаем.
    """
    stats = {"moved": False, "removed_leading_pgbr": False}
    children = list(body)
    paras = [c for c in children if c.tag == qn("w:p")]

    # найти TOC-блок в первых 10 параграфах: page-break + TOC-field + heading
    toc_block: list = []
    for i, p in enumerate(paras[:10]):
        t = "".join((x.text or "") for x in p.findall(".//" + qn("w:t"))).strip()
        if _is_toc_field_para(p):
            # бэк: включаем предыдущий page-break и эту строку TOC
            block_start = i
            # если до TOC-field стоит heading «СОДЕРЖАНИЕ» — его в блок не включаем
            # (оно ПОСЛЕ field в нашей вставке было: pg-break, TOC-field, heading?
            # посмотрим дамп: #0 pgbr, #1 TOC, #2 heading — heading ПОСЛЕ.
            # Значит блок = p[i-1]?(pg-break), p[i] (TOC-field), p[i+1] (heading).
            # + также page break в КОНЦЕ блока если был
            toc_block = []
            # leading page break (if i>=1 is a page-break only para)
            if block_start >= 1 and _is_page_break_only(paras[block_start - 1]):
                toc_block.append(paras[block_start - 1])
            toc_block.append(paras[block_start])
            # heading "СОДЕРЖАНИЕ"
            if block_start + 1 < len(paras) and _is_toc_heading(paras[block_start + 1]):
                toc_block.append(paras[block_start + 1])
            # trailing page break
            nxt_idx = block_start + 1 + (1 if len(toc_block) > 2 else 0)
            if nxt_idx < len(paras) and _is_page_break_only(paras[nxt_idx]):
                toc_block.append(paras[nxt_idx])
            break

    if not toc_block:
        return stats

    # найти параграф «ВВЕДЕНИЕ»
    intro_para = None
    for p in paras:
        t = "".join((x.text or "") for x in p.findall(".//" + qn("w:t"))).strip()
        if t == "ВВЕДЕНИЕ":
            intro_para = p
            break
    if intro_para is None:
        return stats

    # убираем TOC-блок из текущей позиции
    for el in toc_block:
        el.getparent().remove(el)

    # Строим новый TOC-блок в правильном порядке:
    # [page-break] [heading СОДЕРЖАНИЕ (центр)] [TOC-field] [page-break]
    new_block = _build_toc_block()

    # вставляем перед ВВЕДЕНИЕ (но после предшествующего page-break, если он есть,
    # чтобы ВВЕДЕНИЕ всё равно начиналось со своей страницы)
    # Найдём, где ВВЕДЕНИЕ в родителе
    parent = intro_para.getparent()
    idx = list(parent).index(intro_para)
    # Если над ВВЕДЕНИЕ уже стоит page-break-only — вставляем ПЕРЕД ним,
    # чтобы: [старый pg-break → СОДЕРЖАНИЕ → TOC → pg-break из нашего блока → ВВЕДЕНИЕ]
    insert_at = idx
    if idx > 0 and _is_page_break_only(list(parent)[idx - 1]):
        insert_at = idx  # после существующего pg-break
    for el in new_block:
        parent.insert(insert_at, el)
        insert_at += 1

    stats["moved"] = True
    return stats


def _build_toc_block() -> list:
    """[heading СОДЕРЖАНИЕ (center, Heading1)][TOC-field][page-break]"""
    out = []

    # heading
    p_head = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pStyle = OxmlElement("w:pStyle")
    pStyle.set(qn("w:val"), "1")  # наш H1 стиль в доке имеет val="1"
    pPr.append(pStyle)
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), "center")
    pPr.append(jc)
    p_head.append(pPr)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = "СОДЕРЖАНИЕ"
    r.append(t)
    p_head.append(r)
    out.append(p_head)

    # TOC field
    p_toc = OxmlElement("w:p")
    r1 = OxmlElement("w:r")
    fc_begin = OxmlElement("w:fldChar")
    fc_begin.set(qn("w:fldCharType"), "begin")
    r1.append(fc_begin)
    p_toc.append(r1)
    r2 = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = r' TOC \o "1-3" \h \z \u '
    r2.append(instr)
    p_toc.append(r2)
    r3 = OxmlElement("w:r")
    fc_sep = OxmlElement("w:fldChar")
    fc_sep.set(qn("w:fldCharType"), "separate")
    r3.append(fc_sep)
    p_toc.append(r3)
    r4 = OxmlElement("w:r")
    t4 = OxmlElement("w:t")
    t4.text = "Обновите оглавление: правый клик → «Обновить поле» → «Обновить целиком»."
    r4.append(t4)
    p_toc.append(r4)
    r5 = OxmlElement("w:r")
    fc_end = OxmlElement("w:fldChar")
    fc_end.set(qn("w:fldCharType"), "end")
    r5.append(fc_end)
    p_toc.append(r5)
    out.append(p_toc)

    # page break после TOC
    p_pb = OxmlElement("w:p")
    r_pb = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    r_pb.append(br)
    p_pb.append(r_pb)
    out.append(p_pb)

    return out


# ─── 2. Обратная переклейка: Таблица → Рисунок ──────────────────────────

_FIG_RE = re.compile(r"^\s*Рисунок\s+(\d+\.\d+)\s*[—\-–]?\s*(.*)$")
_TBL_RE = re.compile(r"^\s*Таблица\s+(\d+\.\d+)\s*[—\-–]?\s*(.*)$")


def reverse_reclassify(body, target_table_numbers: set) -> dict:
    """Для caption «Таблица N.M» с N.M из target_table_numbers:
    - меняем текст на «Рисунок ??? — ???...» (номер подставится на
      этапе перенумерации).
    """
    stats = {"converted_to_figure": 0,
             "renumbered_figures": 0, "renumbered_tables": 0}
    for p in body.iter(qn("w:p")):
        t = "".join((x.text or "") for x in p.findall(".//" + qn("w:t"))).strip()
        m = _TBL_RE.match(t)
        if not m:
            continue
        if m.group(1) in target_table_numbers:
            rest = m.group(2).strip() or "???"
            set_paragraph_text(p, f"Рисунок ??? — {rest}")
            set_paragraph_style(p, "Caption")
            stats["converted_to_figure"] += 1
    return stats


# ─── 2a. Удаление известного мусора ──────────────────────────────────────

def remove_stray_table4(body) -> int:
    """Удалить параграфы:
        «Таблица 4.»
        «Погрешности опыта короткого замыкания силового трансформатора»
    если они встречаются между «Таблица 5.8» и «Таблица 5.9».
    """
    paras = list(body.iter(qn("w:p")))
    removed = 0
    i = 0
    while i < len(paras):
        t = "".join((x.text or "") for x in paras[i].findall(".//" + qn("w:t"))).strip()
        if t == "Таблица 4.":
            # удаляем сам caption
            paras[i].getparent().remove(paras[i])
            removed += 1
            # и следующий параграф если начинается с «Погрешности опыта…»
            if i + 1 < len(paras):
                nxt_text = "".join(
                    (x.text or "") for x in paras[i + 1].findall(".//" + qn("w:t"))
                ).strip()
                if "Погрешности опыта короткого замыкания" in nxt_text:
                    paras[i + 1].getparent().remove(paras[i + 1])
                    removed += 1
        i += 1
    return removed


# ─── 3. Перенумерация рисунков и таблиц по главам ────────────────────────

def renumber_all(body) -> dict:
    """Итерируем в порядке появления; главы по Heading1 (style «1» или
    «Heading1»). Рисунки и таблицы получают новый N.M в главе.
    """
    stats = {"renumbered_figures": 0, "renumbered_tables": 0}
    chapter_re = re.compile(r"^\s*(\d{1,2})(?:\s|\.|$)")
    fig_rebuild = re.compile(r"^\s*Рисунок\s+(?:\d+\.\d+|\?\?\?)\s*[—\-–]?\s*")
    tbl_rebuild = re.compile(r"^\s*Таблица\s+(?:\d+\.\d+|\?\?\?)\s*[—\-–]?\s*")

    chapter = 0
    fig_ctr = 0
    tbl_ctr = 0

    for p in body.iter(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        style = ""
        if pPr is not None:
            ps = pPr.find(qn("w:pStyle"))
            if ps is not None:
                style = ps.get(qn("w:val")) or ""
        txt = "".join((x.text or "") for x in p.findall(".//" + qn("w:t"))).strip()

        sn_low = style.lower().replace(" ", "")
        is_h1 = style == "1" or sn_low == "heading1" or sn_low.startswith("heading1,")
        if is_h1:
            # Skip "СОДЕРЖАНИЕ" и «ВВЕДЕНИЕ» — в них нет числа
            m = chapter_re.match(txt)
            if m:
                n = int(m.group(1))
                if 1 <= n <= 50:
                    chapter = n
                    fig_ctr = 0
                    tbl_ctr = 0

        # caption?
        mf = _FIG_RE.match(txt)
        mt = _TBL_RE.match(txt)
        is_fig_placeholder = txt.startswith("Рисунок ???")
        is_tbl_placeholder = txt.startswith("Таблица ???")

        if (mf or is_fig_placeholder) and chapter > 0:
            fig_ctr += 1
            new_num = f"{chapter}.{fig_ctr}"
            rest = fig_rebuild.sub("", txt, count=1).strip() or "???"
            set_paragraph_text(p, f"Рисунок {new_num} — {rest}")
            stats["renumbered_figures"] += 1
        elif (mt or is_tbl_placeholder) and chapter > 0:
            tbl_ctr += 1
            new_num = f"{chapter}.{tbl_ctr}"
            rest = tbl_rebuild.sub("", txt, count=1).strip() or "???"
            set_paragraph_text(p, f"Таблица {new_num} — {rest}")
            stats["renumbered_tables"] += 1

    return stats


# ─── 4. Расширенная типография ──────────────────────────────────────────

_EXTRA_TYPO = [
    # уже были в PR #4 — идемпотентно
    ("Openatsimulationstart", "Open at simulation start"),
    ("Display thefullpath", "Display the full path"),
    ("Number ofinputports", "Number of input ports"),
    ("Time spanoverrunaction", "Time span overrun action"),
    ("Limitdatapointstolast", "Limit data points to last"),
    ("Logdatatoworkspace", "Log data to workspace"),
    ("Parameterблока", "Parameter блока"),
    ("Powerguiвбиблиотеке", "Powergui в библиотеке"),
    ("Switchсоответствует", "Switch соответствует"),
    ("FFTAnalysis", "FFT Analysis"),
    # новые
    ("НАУКИИВЫСШЕГО", "НАУКИ И ВЫСШЕГО"),
    ("АлексеенкоМ.А.", "Алексеенко М.А."),
    ("АфанасьевАлександр", "Афанасьев Александр"),
    ("Захарчук А.С. ,АлексеенкоМ.А.", "Захарчук А.С., Алексеенко М.А."),
    ("ФГБОУ ВО“Луганский", "ФГБОУ ВО «Луганский"),
    (" ”, ", "», "),
    ("Даля ”,", "Даля»,"),
    ("напряжениедо", "напряжение до"),
    ("на 10кВ", "на 10 кВ"),
    ("табл. 4.3 -4,5", "табл. 4.3–4.5"),
    ("Параметр Паспортное", "Параметр — Паспортное"),
    # терминология: «в табл.» «в ф-ле» и т.п. редко склеиваются, но проверим
    ("рисунок1", "рисунок 1"),
    ("рисунке1", "рисунке 1"),
    ("рисунок2", "рисунок 2"),
    ("рисунке2", "рисунке 2"),
    ("таблице1", "таблице 1"),
    ("таблице2", "таблице 2"),
]


def apply_extra_typography(body) -> int:
    total = 0
    for p in body.iter(qn("w:p")):
        for old, new in _EXTRA_TYPO:
            total += replace_in_paragraph(p, old, new)
    return total


# ─── 5. Авто-поиск оставшихся склеек (rus+Eng, eng+Rus) ─────────────────

_GLUE_PATTERNS_DYNAMIC = re.compile(
    # слово, в котором русские и латинские буквы соприкасаются без пробела
    r"(?<![\s\.\(\[\-])([а-яёА-ЯЁ]{2,})([A-Z][a-zA-Z]{2,})|"
    r"(?<![\s\.\(\[\-])([a-zA-Z]{3,})([а-яёА-ЯЁ]{2,})"
)


def scan_glued_words(body) -> list:
    """Вернуть уникальные «подозрительные» слова — кириллица/латиница без
    пробела между ними. Не применяет исправления — только репортинг.
    """
    seen = set()
    # собираем ТОЛЬКО текст вне формул
    omml_tag = qn("m:oMath")
    for p in body.iter(qn("w:p")):
        for tn in p.findall(".//" + qn("w:t")):
            anc = tn.getparent()
            inside_math = False
            while anc is not None and anc is not p:
                if anc.tag == omml_tag:
                    inside_math = True
                    break
                anc = anc.getparent()
            if inside_math:
                continue
            txt = tn.text or ""
            for m in _GLUE_PATTERNS_DYNAMIC.finditer(txt):
                tok = m.group(0)
                if len(tok) < 6:
                    continue
                seen.add(tok)
    return sorted(seen)


# ─── main ────────────────────────────────────────────────────────────────

TARGET_FIGURE_MASK_TABLES = {"3.1", "5.4", "6.1", "6.2", "7.1", "8.1", "9.1"}


def post_process(in_path: str, out_path: str) -> dict:
    doc = Document(in_path)
    body = doc.element.body

    report = {}
    report["move_toc"] = move_toc_before_intro(body)
    report["extra_typo"] = apply_extra_typography(body)
    report["stray_removed"] = remove_stray_table4(body)
    report["reverse_reclass"] = reverse_reclassify(body, TARGET_FIGURE_MASK_TABLES)
    report["renumber"] = renumber_all(body)
    report["glued_scan"] = scan_glued_words(body)

    doc.save(out_path)
    return report


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("usage: post_process2.py IN.docx OUT.docx")
        sys.exit(1)
    r = post_process(sys.argv[1], sys.argv[2])
    for k, v in r.items():
        if k == "glued_scan":
            print(f"glued candidates remaining ({len(v)}):")
            for w in v[:40]:
                print(f"  {w}")
        else:
            print(f"{k}: {v}")
