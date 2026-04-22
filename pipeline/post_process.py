# -*- coding: utf-8 -*-
"""Пост-обработка ГОСТ-документа (Session 6+):

1. Вставка поля автособираемого оглавления (TOC) в начало тела — Word обновит
   его автоматически по стилям Heading 1/2/3.
2. Переклейка «рисунков-таблиц»: для указанных номеров рисунков
   (скриншоты таблиц) — либо удаляем caption «Рисунок», либо заменяем его
   на «Таблица N.M — ???» (если над рисунком нет своей таблицы).
3. Разбивка длинных абзацев, где несколько буллетов склеены через « – ».
4. Дополнительная типография (склейки английских имён блоков: Openat-
   simulationstart → Open at simulation start и пр.).
5. Подсветка красным «AI-шных» фрагментов по шаблонам.

Скрипт работает на УЖЕ полностью оформленном GOST-docx — то есть после
полного пайплайна. Выход — новый docx с применёнными правками.
"""
from __future__ import annotations

import re
import sys
from typing import List, Set

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn, nsmap
from lxml import etree

from common import replace_in_paragraph, set_paragraph_style, set_paragraph_text


# ─── 1. Вставка автосборного TOC ──────────────────────────────────────────

def _make_toc_para() -> list:
    """Создать два параграфа: заголовок «СОДЕРЖАНИЕ» и TOC-поле (levels 1-3)."""
    out = []

    # heading "СОДЕРЖАНИЕ"
    p_head = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    pStyle = OxmlElement("w:pStyle")
    pStyle.set(qn("w:val"), "Heading1")
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
    # begin field
    r1 = OxmlElement("w:r")
    fc_begin = OxmlElement("w:fldChar")
    fc_begin.set(qn("w:fldCharType"), "begin")
    r1.append(fc_begin)
    p_toc.append(r1)
    # instruction
    r2 = OxmlElement("w:r")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = r' TOC \o "1-3" \h \z \u '
    r2.append(instr)
    p_toc.append(r2)
    # separate
    r3 = OxmlElement("w:r")
    fc_sep = OxmlElement("w:fldChar")
    fc_sep.set(qn("w:fldCharType"), "separate")
    r3.append(fc_sep)
    p_toc.append(r3)
    # placeholder text that Word will replace upon F9/open
    r4 = OxmlElement("w:r")
    t4 = OxmlElement("w:t")
    t4.text = "Обновите оглавление: правый клик → «Обновить поле» → «Обновить целиком»."
    r4.append(t4)
    p_toc.append(r4)
    # end field
    r5 = OxmlElement("w:r")
    fc_end = OxmlElement("w:fldChar")
    fc_end.set(qn("w:fldCharType"), "end")
    r5.append(fc_end)
    p_toc.append(r5)

    out.append(p_toc)

    # page break after TOC
    p_pb = OxmlElement("w:p")
    rpb = OxmlElement("w:r")
    br = OxmlElement("w:br")
    br.set(qn("w:type"), "page")
    rpb.append(br)
    p_pb.append(rpb)
    out.append(p_pb)

    return out


def insert_auto_toc(body) -> bool:
    """Вставить TOC-поле сразу в начало body (перед первым параграфом).
    Если уже есть старый TOC (поле) — не дублируем.
    """
    # проверяем, нет ли уже TOC-поля
    for instr in body.iter(qn("w:instrText")):
        if instr.text and "TOC" in instr.text:
            return False  # уже есть
    # находим первый параграф
    first_p = None
    for child in body:
        if child.tag == qn("w:p"):
            first_p = child
            break
    if first_p is None:
        return False
    # вставляем перед ним
    toc_elems = _make_toc_para()
    for el in reversed(toc_elems):
        first_p.addprevious(el)
    return True


# ─── 2. Переклейка рисунков-таблиц ────────────────────────────────────────

_FIG_RE = re.compile(r"^\s*Рисунок\s+(\d+\.\d+)\s*[—\-–]?\s*(.*)$")
_TBL_RE = re.compile(r"^\s*Таблица\s+(\d+\.\d+)\s*[—\-–]?\s*(.*)$")


def _para_text(p) -> str:
    return "".join((x.text or "") for x in p.findall(".//" + qn("w:t")))


def _has_drawing(p) -> bool:
    return p.find(".//" + qn("w:drawing")) is not None


def reclassify_image_tables(body, target_figure_numbers: Set[str]) -> dict:
    """Для каждого figure с номером из target_figure_numbers:
    - если в пределах 3 предыдущих параграфов есть caption «Таблица N.M — …» →
      удаляем caption «Рисунок N.M — …» (drawing сам остаётся и относится к
      существующей таблице выше),
    - иначе: заменяем текст caption на «Таблица NEW_N.M — ???», при этом
      NEW_N.M — новый номер таблицы в той же главе (вставляется в общую
      последовательность таблиц).

    Затем перенумеровываем ВСЕ рисунки и ВСЕ таблицы по порядку появления,
    с учётом главы (H1-заголовок «N …»).
    """
    stats = {
        "removed_fig_caption": 0,
        "converted_to_table": 0,
        "renumbered_figures": 0,
        "renumbered_tables": 0,
    }

    # собираем все paragraphs в document order
    paras = list(body.iter(qn("w:p")))

    # индекс параграфов (mutable — будем удалять)
    # шаг A: решить для каждого target figure — удалить или конвертировать
    to_remove = []  # параграфы-рисунки для удаления
    to_convert = []  # параграфы-рисунки для конверсии в «Таблица ??? — ???»

    for i, p in enumerate(paras):
        txt = _para_text(p).strip()
        mf = _FIG_RE.match(txt)
        if not mf:
            continue
        num = mf.group(1)
        if num not in target_figure_numbers:
            continue
        # смотрим 5 предыдущих параграфов — есть ли Таблица-caption
        has_table_above = False
        for j in range(max(0, i - 6), i):
            tt = _para_text(paras[j]).strip()
            if _TBL_RE.match(tt):
                has_table_above = True
                break
        if has_table_above:
            to_remove.append(p)
        else:
            to_convert.append(p)

    # шаг B: удалить или преобразовать
    for p in to_remove:
        parent = p.getparent()
        parent.remove(p)
        stats["removed_fig_caption"] += 1
    for p in to_convert:
        # Меняем текст на «Таблица ??? — ???» (номер переназначим на шаге C)
        set_paragraph_text(p, "Таблица ??? — ???")
        # убеждаемся, что стиль caption
        set_paragraph_style(p, "Caption")
        stats["converted_to_table"] += 1

    # шаг C: перенумеровать рисунки и таблицы в порядке появления, по главам
    # Главы определяем по Heading 1 (style startswith Heading1 или text
    # начинается с "N "  для N=1..50)
    paras = list(body.iter(qn("w:p")))  # пересобрать после удалений
    current_chapter = 0
    fig_ctr = 0
    tbl_ctr = 0
    chapter_re = re.compile(r"^\s*(\d{1,2})(?:\s|\.|$)")

    # Наши регексы — рубить текст до «Рисунок X.Y — …»
    fig_rebuild = re.compile(r"^\s*Рисунок\s+\d+\.\d+\s*[—\-–]?\s*")
    tbl_rebuild = re.compile(r"^\s*Таблица\s+(?:\d+\.\d+|\?\?\?)\s*[—\-–]?\s*")

    for p in paras:
        # проверяем стиль — heading или caption?
        pPr = p.find(qn("w:pPr"))
        style_name = ""
        if pPr is not None:
            pStyle = pPr.find(qn("w:pStyle"))
            if pStyle is not None:
                style_name = pStyle.get(qn("w:val")) or ""
        txt = _para_text(p).strip()

        # Главы — Heading 1. Text обычно «N Название»
        # Стиль может быть «1», «Heading1», «heading 1» и т.п.
        sn_low = style_name.lower().replace(" ", "")
        is_h1 = (
            style_name == "1"
            or sn_low == "heading1"
            or sn_low.startswith("heading1,")
        )
        if is_h1:
            m = chapter_re.match(txt)
            if m:
                new_chap = int(m.group(1))
                if 1 <= new_chap <= 50:
                    current_chapter = new_chap
                    fig_ctr = 0
                    tbl_ctr = 0

        # Captions: ищем «Рисунок …» или «Таблица …»
        mf = _FIG_RE.match(txt)
        mt = _TBL_RE.match(txt)
        is_tbl_placeholder = txt.startswith("Таблица ???")

        if mf and current_chapter > 0:
            fig_ctr += 1
            new_num = f"{current_chapter}.{fig_ctr}"
            # заменить префикс
            rest = fig_rebuild.sub("", txt, count=1)
            if not rest.strip():
                rest = "???"
            new_text = f"Рисунок {new_num} — {rest}"
            # учесть что может уже совпадать — всё равно перепишем чтобы убрать
            # старый номер, но сохранить стиль и выравнивание абзаца
            set_paragraph_text(p, new_text)
            stats["renumbered_figures"] += 1
        elif (mt or is_tbl_placeholder) and current_chapter > 0:
            tbl_ctr += 1
            new_num = f"{current_chapter}.{tbl_ctr}"
            rest = tbl_rebuild.sub("", txt, count=1)
            if not rest.strip():
                rest = "???"
            new_text = f"Таблица {new_num} — {rest}"
            set_paragraph_text(p, new_text)
            stats["renumbered_tables"] += 1

    return stats


# ─── 3. Разбивка склеенных буллет-абзацев ─────────────────────────────────

# Шаблон: текст содержит « – » (en-dash окружённый пробелами) ≥ 2 раз.
# Разбиваем на N параграфов: каждая часть начинается с « – ».

# EN-DASH (U+2013) — разделитель буллетов.
# EM-DASH (U+2014) — разделитель термин/определение ВНУТРИ буллета: НЕ трогаем.
_BULLET_SPLIT = re.compile(r"\s–\s")


def split_glued_bullets(body) -> int:
    """Если в одном параграфе через « – » склеены ≥ 2 буллета — разбиваем."""
    count = 0
    paras = list(body.iter(qn("w:p")))
    for p in paras:
        txt = _para_text(p)
        # считаем вхождения « – »
        splits = _BULLET_SPLIT.findall(txt)
        # нужно минимум 2 вхождения И текст не-пустой до первого дэша
        # (иначе это нормальный абзац с одним дэшем посередине)
        if len(splits) < 2:
            continue
        # не трогаем caption'ы
        if txt.lstrip().startswith(("Рисунок ", "Таблица ", "Формула ")):
            continue
        parts = _BULLET_SPLIT.split(txt)
        # первая часть — заголовок/преамбула, далее — буллеты
        # Пропускаем если первая часть слишком мала и похожа на буллет → тогда
        # все части — буллеты.
        lead = parts[0].rstrip()
        bullets = parts[1:]
        if len(bullets) < 2:
            continue

        parent = p.getparent()
        idx = list(parent).index(p)
        # Обновляем текущий параграф: в нём остаётся только «lead:»
        # если он кончается двоеточием — ок, иначе добавим.
        if lead:
            if not lead.endswith((":", ".")):
                lead = lead + ":"
            set_paragraph_text(p, lead)
        else:
            # пустая лидовая часть — этот параграф превращаем в первый буллет
            set_paragraph_text(p, "— " + bullets[0].lstrip())
            bullets = bullets[1:]

        # Вставляем после p новые параграфы-буллеты
        insert_at = idx + 1
        for b in bullets:
            bb = b.strip()
            if not bb:
                continue
            new_p = OxmlElement("w:p")
            pPr = OxmlElement("w:pPr")
            new_p.append(pPr)
            r = OxmlElement("w:r")
            t = OxmlElement("w:t")
            t.text = "— " + bb
            t.set(qn("xml:space"), "preserve")
            r.append(t)
            new_p.append(r)
            parent.insert(insert_at, new_p)
            insert_at += 1
            count += 1
    return count


# ─── 4. Доп. типография (склейки в англ. именах блоков) ──────────────────

_EXTRA_TYPO = [
    ("Openatsimulationstart", "Open at simulation start"),
    ("Display thefullpath", "Display the full path"),
    ("Number ofinputports", "Number of input ports"),
    ("Time spanoverrunaction", "Time span overrun action"),
    ("Time spanoverrun", "Time span overrun"),
    ("Limitdatapointstolast", "Limit data points to last"),
    ("Logdatatoworkspace", "Log data to workspace"),
    ("Parameterблока", "Parameter блока"),
    ("Powerguiвбиблиотеке", "Powergui в библиотеке"),
    ("Switchсоответствует", "Switch соответствует"),
    ("FFTAnalysis", "FFT Analysis"),
    ("Maxstepsize", "Max step size"),
    ("Stoptime", "Stop time"),
    ("Starttime", "Start time"),
    # Session 6 доп. — из bullets
    ("Sampletime", "Sample time"),
    ("Inputprocessing", "Input processing"),
    ("Maximizeaxes", "Maximize axes"),
    ("Axesscaling", "Axes scaling"),
    ("Elements aschannels", "Elements as channels"),
    ("Columnsaschannels", "Columns as channels"),
    ("Time displayoffset", "Time display offset"),
    ("Time-axislabels", "Time-axis labels"),
    ("Bottomdisplaysonly", "Bottom displays only"),
    ("Datahistory", "Data history"),
    ("Decimation", "Decimation"),
    ("Autoscale", "Autoscale"),
    # частые склейки — термин-название + рус. слово
    ("Librarybrowser", "Library browser"),
    ("Newmodel", "New model"),
    ("Simulinklibrary", "Simulink library"),
]


def apply_extra_typography(body) -> int:
    total = 0
    for p in body.iter(qn("w:p")):
        for old, new in _EXTRA_TYPO:
            if old == new:
                continue
            n = replace_in_paragraph(p, old, new)
            total += n
    return total


# ─── 5. Подсветка AI-шных абзацев красным ────────────────────────────────

# Специфичные AI-маркеры. Стараемся избежать ложных срабатываний на
# нейтральных словах («удобно», «полезно», «рассмотрим» — нет!)
_AI_MARKERS = [
    # прямые обращения к читателю (очень характерно для ChatGPT)
    "Приветствую", "Уважаемый пользователь", "Уважаемый читатель",
    "Как эксперт", "Как специалист", "я с удовольствием помогу",
    "Как вы и просили", "Как вы просили", "в запрошенном формате",
    "в запрошенном вами", "ориентируясь на предоставленные",
    "предоставленные скриншоты", "предоставленные вами",
    "ваш запрос", "Вашему вниманию", "Ваше пособие", "ваш вопрос",
    "если вам понадобится", "если у вас есть вопросы",
    "Надеюсь, это поможет", "Надеюсь, эта информация",
    "Если вам нужна дополнительная", "Готов помочь",
    "с радостью помогу", "Отлично!", "Прекрасно!", "Замечательно!",
    "Конечно!", "Разумеется!",
    # мета-комментарии про саму структуру ответа
    "ниже приведена подробная", "ниже я приведу", "приведу подробную",
    "подробную расшифровку", "детальный разбор",
    "дам вам подробный", "дам подробный разбор",
    "рассмотрим каждый из них", "рассмотрим каждый параметр",
    "пройдёмся по", "пройдемся по",
    # цветистые заключительные обороты
    "ключ к удобному", "ключ к успешному", "залог успех",
    "изящное решение", "элегантное решение",
    "что обеспечит наилучш", "что даст нам",
    # мета-обращения в тексте пособия
    "в данном разделе мы рассмотрим", "в этом разделе мы рассмотрим",
    "мы с вами рассмотрим", "как вы могли заметить",
    "стоит отметить, что", "важно понимать, что",
    "подведя итог", "итак, мы видим",
    "шаг за шагом", "в этом нам поможет",
]


def _set_run_red(r_elem) -> None:
    """Установить красный цвет шрифта для данного <w:r>."""
    rPr = r_elem.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r_elem.insert(0, rPr)
    # удалим существующий w:color если есть
    for c in rPr.findall(qn("w:color")):
        rPr.remove(c)
    color = OxmlElement("w:color")
    color.set(qn("w:val"), "C00000")  # тёмно-красный
    rPr.append(color)
    # также поставим yellow highlight для большей заметности
    for hl in rPr.findall(qn("w:highlight")):
        rPr.remove(hl)
    hl = OxmlElement("w:highlight")
    hl.set(qn("w:val"), "yellow")
    rPr.append(hl)


def highlight_ai_text(body) -> list:
    """Вернуть список индексов параграфов, подсвеченных красным."""
    highlighted = []
    for i, p in enumerate(body.iter(qn("w:p"))):
        t = _para_text(p)
        tl = t.lower()
        for m in _AI_MARKERS:
            if m.lower() in tl:
                # красим все w:r в параграфе
                for r in p.findall(qn("w:r")):
                    _set_run_red(r)
                highlighted.append((i, t[:160]))
                break
    return highlighted


# ─── Главный пост-процессор ──────────────────────────────────────────────

TARGET_IMAGE_TABLES = {"4.5", "4.6", "4.7", "4.8", "5.2", "5.3", "5.13", "5.15", "5.16"}


def post_process(in_path: str, out_path: str) -> dict:
    doc = Document(in_path)
    body = doc.element.body

    report = {}
    report["toc_inserted"] = insert_auto_toc(body)
    report["extra_typography"] = apply_extra_typography(body)
    report["glued_bullets_split"] = split_glued_bullets(body)
    report["reclassify"] = reclassify_image_tables(body, TARGET_IMAGE_TABLES)
    report["ai_highlighted"] = highlight_ai_text(body)

    doc.save(out_path)
    return report


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("usage: post_process.py IN.docx OUT.docx")
        sys.exit(1)
    rep = post_process(sys.argv[1], sys.argv[2])
    # коротко
    print(f"TOC inserted:       {rep['toc_inserted']}")
    print(f"Typography extras:  {rep['extra_typography']}")
    print(f"Bullets split:      {rep['glued_bullets_split']}")
    print(f"Reclassify:         {rep['reclassify']}")
    print(f"AI-highlighted:     {len(rep['ai_highlighted'])}")
    for idx, snip in rep["ai_highlighted"][:20]:
        print(f"  #{idx}: {snip}")
