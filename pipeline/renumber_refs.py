"""Перенумерация «Рисунок N.M» и «Таблица N.M» в соответствии с НОВОЙ
структурой глав, и синхронизация inline-ссылок («рисунок 2,9»,
«см. таблицу 4.3» и т.п.).

Проблема после restructure: заголовки глав переиначены (главы 5, 6, 7, 8
уже не те, что в исходнике), а подписи рисунков и таблиц остались со
старой нумерацией. Отсюда:
  - «Рисунок 2.9» внутри 5-й главы
  - «Таблица 2.9» встречается 3 раза
  - «Рисунок 6.1» — дубликат после предыдущих глав

Алгоритм:
  Проход 1. Идём по документу, в каждом H1-заголовке извлекаем номер главы;
            у каждой встреченной подписи (Рисунок/Таблица) сохраняем
            старый номер и присваиваем новый = (глава, порядковый+1).
  Проход 2. Применяем замены:
            • в подписях — заменяем «Рисунок N.M» на «Рисунок NEW»;
              также чиним «N.M — заголовок-следующего-раздела» на «N.M — ???»
              (если title явно прилетел из соседнего H2).
            • в inline-ссылках — ищем «рисунок N.M/N,M», «табл. N.M»,
              «рисунке N.M», «рисунков N.M и N.M» и заменяем по карте.
"""
from __future__ import annotations

import re
from collections import Counter
from typing import Dict, List, Tuple

from docx import Document
from docx.oxml.ns import qn

from common import set_paragraph_text


# ── подписи ───────────────────────────────────────────────────────────────
_RE_FIG_CAP = re.compile(
    r"^\s*Рисунок\s+(\d+)[.,](\d+)\s*[—–-]?\s*(.*)$",
    re.I,
)
_RE_TBL_CAP = re.compile(
    r"^\s*Таблица\s+(\d+)[.,](\d+)\s*[—–-]?\s*(.*)$",
    re.I,
)

# inline-ссылки: «(рисунок 2,9)», «на рисунке 4.3», «рис. 1.2»
_RE_FIG_REF = re.compile(
    r"(\b(?:р[ии]сун(?:ок|ка|ке|ку|ки|ков)?|рис\.)\s+)(\d+)[.,](\d+)",
    re.I,
)
_RE_TBL_REF = re.compile(
    r"(\b(?:табл[ии]ц(?:а|ы|е|у|ам)?|табл\.)\s+)(\d+)[.,](\d+)",
    re.I,
)

# заголовок главы: «1. ...», «10. ...», «8. ...»
_RE_H1_NUM = re.compile(r"^\s*(\d+)\.\s+")
# заголовок уровня 2/3: «1.5 ...», «5.2.1 ...»
_RE_H2H3_NUM = re.compile(r"^\s*(\d+(?:\.\d+){1,2})\s+")


def _is_heading(p, levels=(1, 2, 3)) -> int:
    sn = (p.style.name or "").lower() if p.style else ""
    for lv in levels:
        if sn == f"heading {lv}":
            return lv
    return 0


def _title_looks_like_next_heading(title: str) -> bool:
    """Подписи, которые случайно захватили текст следующего подраздела.
    Пример: «Рисунок 4.5 — 4.8 Силовые кабели с изоляцией…».
    Признак: title начинается с «N.M » или «N.M.K ».
    """
    return bool(_RE_H2H3_NUM.match(title.strip()))


def _title_looks_corrupt(title: str) -> bool:
    """Прочие явно поломанные title: начинается с фрагмента текста,
    обрывающегося без смысла («В качестве базисного напряжения ()»,
    «9 Значения напряжения …» и т.п.).
    """
    t = title.strip()
    if not t:
        return True
    # одиночное число в начале (отрубленный номер):
    if re.match(r"^\d+\s+[а-яА-Я]", t):
        return True
    # обрывающееся «(…)» или «()» в первой половине
    if t.startswith("()") or "()" in t[:40]:
        return True
    return False


def _iter_all_paragraphs(doc):
    """Все Paragraph в порядке появления, включая вложенные в таблицы.

    Используем python-docx wrapper (чтобы style.name работал через
    styles.xml), проходим по всем <w:p> в document-order.
    """
    from docx.text.paragraph import Paragraph
    body = doc.element.body
    for p_elem in body.iter(qn("w:p")):
        yield Paragraph(p_elem, doc.part)


def _text_of(p_elem):
    parts = [t.text or "" for t in p_elem.findall(".//" + qn("w:t"))]
    return "".join(parts).strip()


def _style_name(p) -> str:
    try:
        return (p.style.name or "").lower() if p.style else ""
    except Exception:
        return ""


def _is_heading_para(p):
    sn = _style_name(p)
    for lv in (1, 2, 3):
        if sn == f"heading {lv}":
            return lv
    return 0


# ── Определение главы по ТЕКСТУ (на случай, если стили ещё не проставлены) ─
# Формат: «N. UPPERCASE TITLE» — короткий, >50% прописных.
_RE_CHAPTER_TEXT = re.compile(r"^\s*(\d+)\.\s+(.+)$")


def _detect_chapter_by_text(text: str) -> int:
    """Возвращает номер главы, если текст похож на заголовок главы H1.

    Пример: «5. СХЕМЫ ЗАМЕЩЕНИЯ ТРАНСФОРМАТОРОВ» → 5
    Не-пример: «5.1 Общие сведения» → 0
    Не-пример: «5. Список пунктов:» (обычный текст) → 0 (<50% caps)
    """
    t = (text or "").strip()
    m = _RE_CHAPTER_TEXT.match(t)
    if not m:
        return 0
    body = m.group(2).strip()
    if len(body) < 3 or len(body) > 140:
        return 0
    letters = [c for c in body if c.isalpha()]
    if not letters:
        return 0
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    if upper_ratio < 0.6:
        return 0
    try:
        return int(m.group(1))
    except ValueError:
        return 0


def renumber(body_in: str, body_out: str) -> dict:
    doc = Document(body_in)
    paras = list(_iter_all_paragraphs(doc))
    para_elems = [p._element for p in paras]

    # ── Проход 1: собираем карты ──────────────────────────────────────────
    current_chapter = 0
    fig_counter: Dict[int, int] = {}
    tbl_counter: Dict[int, int] = {}

    # карты: old_str ("2.9") → new_str ("5.3")
    fig_map: Dict[str, str] = {}
    tbl_map: Dict[str, str] = {}

    # зафиксируем порядок подписей для второй фазы
    cap_fixes: List[Tuple[int, str, str]] = []  # (paragraph_idx, old_text, new_text)
    orphan_count = 0
    corrupt_title_count = 0

    for i, p in enumerate(paras):
        p_el = para_elems[i]
        lv = _is_heading_para(p)
        t = _text_of(p_el)
        # Хотя стили главы к этому моменту ещё могут быть не проставлены
        # (GOST-форматирование идёт после), определяем главу по тексту.
        chap_by_text = _detect_chapter_by_text(t)
        if lv == 1 or chap_by_text:
            if chap_by_text:
                current_chapter = chap_by_text
            else:
                m = _RE_H1_NUM.match(t)
                if m:
                    current_chapter = int(m.group(1))
            fig_counter.setdefault(current_chapter, 0)
            tbl_counter.setdefault(current_chapter, 0)
            continue

        mf = _RE_FIG_CAP.match(t)
        if mf:
            old = f"{mf.group(1)}.{mf.group(2)}"
            title = (mf.group(3) or "").strip()
            if current_chapter == 0:
                # подпись до первой главы — оставим как есть
                continue
            fig_counter[current_chapter] += 1
            new = f"{current_chapter}.{fig_counter[current_chapter]}"
            fig_map[old] = new
            # чиним title если туда попал заголовок следующего раздела
            if _title_looks_like_next_heading(title) or _title_looks_corrupt(title):
                corrupt_title_count += 1
                title = "???"
            if title == "":
                title = "???"
            new_text = f"Рисунок {new} — {title}"
            cap_fixes.append((i, t, new_text))
            continue

        mt = _RE_TBL_CAP.match(t)
        if mt:
            old = f"{mt.group(1)}.{mt.group(2)}"
            title = (mt.group(3) or "").strip()
            if current_chapter == 0:
                continue
            tbl_counter[current_chapter] += 1
            new = f"{current_chapter}.{tbl_counter[current_chapter]}"
            tbl_map[old] = new
            if _title_looks_like_next_heading(title) or _title_looks_corrupt(title):
                corrupt_title_count += 1
                title = "???"
            if title:
                new_text = f"Таблица {new} — {title}"
            else:
                new_text = f"Таблица {new}"
            cap_fixes.append((i, t, new_text))
            continue

    # ── Проход 2: применяем фиксы подписей ────────────────────────────────
    caption_indices = {idx for idx, _, _ in cap_fixes}
    for idx, old_text, new_text in cap_fixes:
        set_paragraph_text(para_elems[idx], new_text)

    # inline-refs: пробегаем все параграфы и делаем текстовую подстановку
    inline_fig = 0
    inline_tbl = 0

    def _sub_fig(m: re.Match) -> str:
        nonlocal inline_fig
        old = f"{m.group(2)}.{m.group(3)}"
        new = fig_map.get(old)
        if not new:
            return m.group(0)
        inline_fig += 1
        return f"{m.group(1)}{new}"

    def _sub_tbl(m: re.Match) -> str:
        nonlocal inline_tbl
        old = f"{m.group(2)}.{m.group(3)}"
        new = tbl_map.get(old)
        if not new:
            return m.group(0)
        inline_tbl += 1
        return f"{m.group(1)}{new}"

    # обрабатываем каждый w:p с накоплением текста по runs; работаем и
    # с вложенными в таблицы параграфами.
    # ВАЖНО: пропускаем сами абзацы-подписи, иначе inline-sub перепишет
    # их ещё раз по fig_map (который dict, и при дубликатах старых номеров
    # хранит последнее значение → создаёт ложные дубли).
    for i, p_el in enumerate(para_elems):
        if i in caption_indices:
            continue
        t = _text_of(p_el)
        if not t:
            continue
        if not re.search(r"(р[ии]сун|рис\.|табл[ии]ц|табл\.)", t, re.I):
            continue
        t_nodes = p_el.findall(".//" + qn("w:t"))
        joined = "".join((tn.text or "") for tn in t_nodes)
        new_joined = _RE_FIG_REF.sub(_sub_fig, joined)
        new_joined = _RE_TBL_REF.sub(_sub_tbl, new_joined)
        if new_joined != joined and t_nodes:
            t_nodes[0].text = new_joined
            t_nodes[0].set(qn("xml:space"), "preserve")
            for tn in t_nodes[1:]:
                tn.text = ""

    doc.save(body_out)
    return {
        "fig_renumbered": len(fig_map),
        "tbl_renumbered": len(tbl_map),
        "inline_fig_refs_fixed": inline_fig,
        "inline_tbl_refs_fixed": inline_tbl,
        "captions_corrupt_titles": corrupt_title_count,
        "fig_map_sample": dict(list(fig_map.items())[:15]),
        "tbl_map_sample": dict(list(tbl_map.items())[:15]),
    }


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 3:
        print("Usage: python renumber_refs.py <in.docx> <out.docx>")
        sys.exit(1)
    r = renumber(sys.argv[1], sys.argv[2])
    for k, v in r.items():
        print(k, "=", v)
