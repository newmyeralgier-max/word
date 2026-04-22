"""Подписи к рисункам и таблицам с учётом структуры документа.

Важное открытие: в исходнике большинство рисунков находятся внутри 2x1-таблиц,
где:
  row 0 col 0 = картинка
  row 1 col 0 = текст подписи вида «Рис.1.1. Название»

Такая таблица-«контейнер» — это рисунок, а НЕ таблица в ГОСТ-смысле.

Алгоритм:
1. Классифицируем каждую w:tbl:
   - FIGURE_CONTAINER: внутри есть w:drawing (и, возможно, подпись)
   - DATA: обычная таблица данных (без w:drawing)
2. Для FIGURE_CONTAINER: найти/распарсить существующую подпись внутри таблицы,
   нормализовать её в формат «Рисунок N.M — Название». Если номера нет —
   присвоить следующий в главе.
3. Для DATA: вставить параграф-подпись «Таблица N.M — Название» ПЕРЕД таблицей
   (если её ещё нет).
4. Для свободных w:drawing (встречаются редко — 12 шт.) — вставить подпись
   «Рисунок N.M — Название» следующим параграфом.

Все подписи сохраняем в JSON для возможной ручной доводки.
"""
from __future__ import annotations

import copy
import json
import re
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import List, Optional, Tuple

from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from common import (
    iter_body_elements,
    paragraph_has_drawing,
    paragraph_tag,
    paragraph_text,
    set_paragraph_text,
    table_tag,
)


# ─── Парсинг существующих подписей ────────────────────────────────────────

_RE_FIG_ANY = re.compile(
    r"^\s*(?:Рис(?:унок|\.)?)\s*(\d+)\s*[\.,]\s*(\d+)\s*[\.\-–—: ]*\s*(.*)$",
    re.I,
)
_RE_FIG_NO_NUM = re.compile(r"^\s*(?:Рис(?:унок|\.)?)\s+(\D.*)$", re.I)
_RE_TBL_ANY = re.compile(
    r"^\s*(?:Табл(?:ица|\.)?)\s*(\d+)\s*[\.,]\s*(\d+)\s*[\.\-–—: ]*\s*(.*)$",
    re.I,
)
_CHAPTER_RE = re.compile(r"^\s*(\d+)\.\s+[А-ЯЁA-Z][А-ЯЁA-Z ,\-\u2013]{3,}")


def detect_chapter(text: str) -> Optional[int]:
    if not text:
        return None
    m = _CHAPTER_RE.match(text.strip())
    if not m:
        return None
    body = text.strip()[m.end(1):].lstrip(". ").strip()
    if not body:
        return None
    letters = [c for c in body if c.isalpha()]
    if not letters:
        return None
    upper_ratio = sum(1 for c in letters if c.isupper()) / len(letters)
    if upper_ratio < 0.6:
        return None
    return int(m.group(1))


# ─── Нормализация подписи к рисунку ───────────────────────────────────────

_TYPO_SPLIT = [
    ("Окнообозревателябиблиотек", "Окно обозревателя библиотек "),
    ("Окнообозревателя", "Окно обозревателя "),
    ("обозревателябиблиотек", "обозревателя библиотек "),
    ("обозревателябиблиотеки", "обозревателя библиотеки "),
    ("SimulinkLibraryBrowser", "Simulink Library Browser"),
    ("SimulinkLibrary", "Simulink Library"),
    ("SimulinkModel", "Simulink Model"),
    ("MatlabSimulink", "Matlab Simulink"),
    ("SimscapeElectrical", "Simscape Electrical"),
    ("ACVoltageSource", "AC Voltage Source"),
    ("DCVoltageSource", "DC Voltage Source"),
    ("ControlledVoltageSource", "Controlled Voltage Source"),
    ("ControlledCurrentSource", "Controlled Current Source"),
    ("ProgrammableVoltageSource", "Programmable Voltage Source"),
    ("Three-PhaseSource", "Three-Phase Source"),
    ("Three-PhaseBreaker", "Three-Phase Breaker"),
    ("Three-PhasePISectionLine", "Three-Phase PI Section Line"),
    ("Three-PhaseV-IMeasurement", "Three-Phase V-I Measurement"),
    ("Three-PhaseFault", "Three-Phase Fault"),
    ("Three-PhaseProgrammableVoltageSource", "Three-Phase Programmable Voltage Source"),
    ("LinearTransformer", "Linear Transformer"),
    ("SaturableTransformer", "Saturable Transformer"),
    ("IdealSwitch", "Ideal Switch"),
    ("CurrentMeasurement", "Current Measurement"),
    ("VoltageMeasurement", "Voltage Measurement"),
    ("PiSectionLine", "Pi Section Line"),
    ("InitialStateSetting", "Initial State Setting"),
    ("SimulinkLibrary", "Simulink Library"),
    ("MATLABv2020b", "MATLAB v2020b"),
    ("БлокACVoltageSource", "Блок AC Voltage Source"),
    ("БлокControlledVoltageSource", "Блок Controlled Voltage Source"),
    ("БлокThree-PhaseSource", "Блок Three-Phase Source"),
    ("БлокThree-PhaseBreaker", "Блок Three-Phase Breaker"),
    ("БлокThree-PhasePISectionLine", "Блок Three-Phase PI Section Line"),
    ("БлокLinearTransformer", "Блок Linear Transformer"),
    ("БлокSaturableTransformer", "Блок Saturable Transformer"),
    ("БлокIdealSwitch", "Блок Ideal Switch"),
    ("БлокBreaker", "Блок Breaker"),
    ("БлокScope", "Блок Scope"),
    ("БлокDisplay", "Блок Display"),
    ("БлокMultimeter", "Блок Multimeter"),
    ("БлокThree-PhaseProgrammableVoltageSource", "Блок Three-Phase Programmable Voltage Source"),
    ("БлокFault", "Блок Fault"),
    ("БлокThree-PhaseFault", "Блок Three-Phase Fault"),
    ("Three-Phase", "Three-Phase"),
    ("библиотекSimulink", "библиотек Simulink"),
]


def fix_typo(text: str) -> str:
    # сначала длинные ключи
    keys = sorted(set(k for k, _ in _TYPO_SPLIT), key=len, reverse=True)
    out = text
    for k in keys:
        v = next(v for kk, v in _TYPO_SPLIT if kk == k)
        out = out.replace(k, v)
    # убрать двойные пробелы
    out = re.sub(r"\s+", " ", out).strip()
    return out


def normalize_fig_caption(raw: str) -> Tuple[Optional[Tuple[int, int]], str]:
    """Распарсить подпись к рисунку.
    Возвращает ((chapter, seq) или None, title)
    """
    if not raw:
        return None, ""
    m = _RE_FIG_ANY.match(raw.strip())
    if m:
        ch = int(m.group(1))
        seq = int(m.group(2))
        title = fix_typo(m.group(3).rstrip(". "))
        return (ch, seq), title
    m2 = _RE_FIG_NO_NUM.match(raw.strip())
    if m2:
        title = fix_typo(m2.group(1).rstrip(". "))
        return None, title
    return None, fix_typo(raw.strip())


def normalize_tbl_caption(raw: str) -> Tuple[Optional[Tuple[int, int]], str]:
    if not raw:
        return None, ""
    m = _RE_TBL_ANY.match(raw.strip())
    if m:
        ch = int(m.group(1))
        seq = int(m.group(2))
        title = fix_typo(m.group(3).rstrip(". "))
        return (ch, seq), title
    return None, fix_typo(raw.strip())


# ─── Классификация таблиц ─────────────────────────────────────────────────

def table_has_drawings(t_elem) -> bool:
    return bool(t_elem.findall(".//" + qn("w:drawing")))


def is_figure_container_table(t_elem) -> bool:
    """Определить, является ли таблица «контейнером для рисунка».

    Критерии:
    - 1-3 строки всего
    - хотя бы в одной ячейке есть w:drawing
    - и НЕТ признаков табличных данных: например, текстовые ячейки с
      числовыми значениями в нескольких строках подряд.

    Всё прочее (многорядные таблицы с числовыми данными, даже с случайно
    попавшим в ячейку w:drawing) — это таблица данных (DATA).
    """
    rows = [c for c in t_elem if c.tag == qn("w:tr")]
    if not rows:
        return False
    if not t_elem.findall(".//" + qn("w:drawing")):
        return False

    # Если строк много (>= 4) — почти наверняка таблица данных
    if len(rows) >= 4:
        return False

    # Считаем «текстовые» ячейки (не-пустой текст БЕЗ drawing)
    text_cells = 0
    drawing_cells = 0
    for row in rows:
        for cell in row.findall(".//" + qn("w:tc")):
            has_draw = bool(cell.findall(".//" + qn("w:drawing")))
            txt = "".join(
                (x.text or "") for x in cell.findall(".//" + qn("w:t"))
            ).strip()
            if has_draw:
                drawing_cells += 1
            elif txt:
                text_cells += 1

    # Маленькая таблица (≤3 строки) с одной ячейкой-рисунком и опциональной
    # подписью — это контейнер рисунка.
    if drawing_cells == 1 and text_cells <= 2:
        return True

    # 2 ячейки-рисунка и 0-1 текстовых — тоже контейнер (две стыкованные
    # картинки с общей подписью).
    if drawing_cells >= 2 and text_cells <= drawing_cells:
        return True

    return False


def table_first_drawing_row_caption(t_elem) -> Optional[object]:
    """Вернуть параграф (XML) из строки ПОСЛЕ строки с картинкой — обычно в ней подпись.

    Возвращает первый параграф подписи (последней текстовой строки) или None.
    """
    rows = [c for c in t_elem if c.tag == qn("w:tr")]
    # идём по рядам, ищем ряд с drawing и следующий за ним ряд
    found_drawing_row = False
    for row in rows:
        has_dr = bool(row.findall(".//" + qn("w:drawing")))
        if has_dr and not found_drawing_row:
            found_drawing_row = True
            continue
        if found_drawing_row:
            # возвращаем первый НЕпустой параграф этого ряда
            paras = row.findall(".//" + qn("w:p"))
            for p in paras:
                t = "".join(x.text or "" for x in p.findall(".//" + qn("w:t"))).strip()
                if t:
                    return p
            # если в ряду нет текста, но есть drawing — это не подпись
            if row.findall(".//" + qn("w:drawing")):
                continue
    # иначе: в той же таблице ряд drawing один, а в последующих рядов нет →
    # берём самый последний непустой параграф всей таблицы БЕЗ drawing
    last_cap = None
    for row in rows:
        if row.findall(".//" + qn("w:drawing")):
            continue
        for p in row.findall(".//" + qn("w:p")):
            t = "".join(x.text or "" for x in p.findall(".//" + qn("w:t"))).strip()
            if t:
                last_cap = p
    return last_cap


# ─── Сборка подписи ───────────────────────────────────────────────────────

def make_caption_paragraph(text: str, align: str = "center") -> object:
    p = OxmlElement("w:p")
    pPr = OxmlElement("w:pPr")
    jc = OxmlElement("w:jc")
    jc.set(qn("w:val"), {"center": "center", "left": "left"}[align])
    pPr.append(jc)
    p.append(pPr)
    r = OxmlElement("w:r")
    t = OxmlElement("w:t")
    t.text = text
    t.set(qn("xml:space"), "preserve")
    r.append(t)
    p.append(r)
    return p


# ─── Генерация черновых названий ──────────────────────────────────────────

_STOP_PHRASES = re.compile(
    r"(как\s+показано\s+на\s+рис|представлен[аоы]?\s+на\s+рис|"
    r"изображен[аоы]?\s+на\s+рис|приведен[аоы]?\s+на\s+рис|"
    r"показан[аоы]?\s+на\s+рис|см\.\s*рис|\(рис|"
    r"представлен[аоы]?\s+в\s+таблиц|приведен[аоы]?\s+в\s+таблиц|"
    r"сводн(?:ы[ех])?\s+парам|исходн(?:ы[ех])?\s+данн)",
    re.I,
)


def _clean(snippet: str, max_words: int = 10) -> str:
    s = snippet.strip().rstrip(".:,;— ")
    s = re.sub(r"\s+", " ", s)
    words = s.split()
    if len(words) > max_words:
        s = " ".join(words[:max_words]).rstrip(",;.:— ")
    return s


def draft_title_from_context(context_before: str, context_after: str = "") -> Tuple[str, str]:
    """Вернёт (title, confidence)."""
    m = _STOP_PHRASES.search(context_before)
    if m:
        head = context_before[:m.start()].strip()
        parts = re.split(r"[.!?]\s+", head)
        if parts:
            cand = _clean(parts[-1])
            if 3 <= len(cand.split()) <= 12:
                return cand, "med"
    parts = re.split(r"[.!?]\s+", context_after.strip())
    if parts and parts[0]:
        cand = _clean(parts[0])
        if 3 <= len(cand.split()) <= 12:
            return cand, "low"
    if context_before.strip():
        cand = _clean(context_before.split(".")[-1])
        if cand and len(cand.split()) >= 2:
            return cand, "low"
    return "???", "low"


# ─── Data classes for output ──────────────────────────────────────────────

@dataclass
class CaptionRec:
    kind: str  # 'figure' | 'table'
    chapter: int
    seq: int
    text: str
    confidence: str
    source: str  # 'existing_normalized' | 'generated' | 'merged'
    old_text: str = ""
    context: str = ""


# ─── Основная функция ────────────────────────────────────────────────────

def process_document(
    input_path: str, output_path: str, captions_json_path: str
) -> dict:
    doc = Document(input_path)
    body = doc.element.body
    children = [c for c in body if c.tag != qn("w:sectPr")]

    # Пред-сбор текста для контекстов
    idx_text: List[str] = []
    for el in children:
        if paragraph_tag(el):
            idx_text.append(paragraph_text(el))
        else:
            idx_text.append("")

    cur_chapter = 0
    fig_counter = {}
    tbl_counter = {}

    figure_recs: List[CaptionRec] = []
    table_recs: List[CaptionRec] = []

    inserts_after: List[Tuple[int, object]] = []
    inserts_before: List[Tuple[int, object]] = []

    def _ctx(start: int, before: int, after: int) -> Tuple[str, str]:
        b, a = [], []
        for j in range(start - 1, max(-1, start - before - 1), -1):
            t = idx_text[j]
            if t and t.strip():
                b.insert(0, t.strip())
                if len(b) >= before:
                    break
        for j in range(start + 1, min(len(idx_text), start + after + 1)):
            t = idx_text[j]
            if t and t.strip():
                a.append(t.strip())
                if len(a) >= after:
                    break
        return " ".join(b), " ".join(a)

    for i, el in enumerate(children):
        if paragraph_tag(el):
            ch = detect_chapter(idx_text[i])
            if ch is not None:
                cur_chapter = ch
            if paragraph_has_drawing(el):
                # Свободный рисунок — ищем подпись в следующих 2 параграфах
                if cur_chapter == 0:
                    cur_chapter = 1
                next_cap_idx = None
                for j in range(i + 1, min(len(children), i + 3)):
                    tj = idx_text[j]
                    if _RE_FIG_ANY.match(tj or "") or _RE_FIG_NO_NUM.match(tj or ""):
                        next_cap_idx = j
                        break
                if next_cap_idx is not None:
                    # нормализуем существующую подпись
                    parsed_num, title = normalize_fig_caption(idx_text[next_cap_idx])
                    if parsed_num:
                        ch_num, seq = parsed_num
                        fig_counter[ch_num] = max(fig_counter.get(ch_num, 0), seq)
                    else:
                        ch_num = cur_chapter
                        seq = fig_counter.get(ch_num, 0) + 1
                        fig_counter[ch_num] = seq
                    new_text = f"Рисунок {ch_num}.{seq} — {title or '???'}"
                    set_paragraph_text(children[next_cap_idx], new_text)
                    figure_recs.append(
                        CaptionRec(
                            kind="figure",
                            chapter=ch_num,
                            seq=seq,
                            text=new_text,
                            confidence="high" if title and title != "???" else "low",
                            source="existing_normalized",
                            old_text=idx_text[next_cap_idx][:200],
                        )
                    )
                else:
                    # подписи нет — генерируем
                    ch_num = cur_chapter
                    seq = fig_counter.get(ch_num, 0) + 1
                    fig_counter[ch_num] = seq
                    ctx_b, ctx_a = _ctx(i, 3, 3)
                    title, conf = draft_title_from_context(ctx_b, ctx_a)
                    new_text = f"Рисунок {ch_num}.{seq} — {title}"
                    inserts_after.append((i, make_caption_paragraph(new_text)))
                    figure_recs.append(
                        CaptionRec(
                            kind="figure",
                            chapter=ch_num,
                            seq=seq,
                            text=new_text,
                            confidence=conf,
                            source="generated",
                            context=ctx_b[:200],
                        )
                    )
            continue

        # Таблица
        if not table_tag(el):
            continue

        if cur_chapter == 0:
            cur_chapter = 1

        if is_figure_container_table(el):
            # Это «фигурный контейнер» — ищем подпись внутри
            cap_p = table_first_drawing_row_caption(el)
            if cap_p is not None:
                raw = "".join(
                    x.text or "" for x in cap_p.findall(".//" + qn("w:t"))
                ).strip()
                parsed_num, title = normalize_fig_caption(raw)
                if parsed_num:
                    ch_num, seq = parsed_num
                    fig_counter[ch_num] = max(fig_counter.get(ch_num, 0), seq)
                else:
                    ch_num = cur_chapter
                    seq = fig_counter.get(ch_num, 0) + 1
                    fig_counter[ch_num] = seq
                new_text = f"Рисунок {ch_num}.{seq} — {title or '???'}"
                set_paragraph_text(cap_p, new_text)
                figure_recs.append(
                    CaptionRec(
                        kind="figure",
                        chapter=ch_num,
                        seq=seq,
                        text=new_text,
                        confidence="high" if title and title != "???" else "low",
                        source="existing_normalized",
                        old_text=raw[:200],
                    )
                )
            else:
                # подписи в таблице нет — добавляем параграф ПОСЛЕ таблицы
                ch_num = cur_chapter
                seq = fig_counter.get(ch_num, 0) + 1
                fig_counter[ch_num] = seq
                ctx_b, ctx_a = _ctx(i, 3, 3)
                title, conf = draft_title_from_context(ctx_b, ctx_a)
                new_text = f"Рисунок {ch_num}.{seq} — {title}"
                inserts_after.append((i, make_caption_paragraph(new_text)))
                figure_recs.append(
                    CaptionRec(
                        kind="figure",
                        chapter=ch_num,
                        seq=seq,
                        text=new_text,
                        confidence=conf,
                        source="generated",
                        context=ctx_b[:200],
                    )
                )
        else:
            # Реальная таблица данных
            # Проверяем, есть ли уже подпись в предыдущих 1-2 параграфах
            existing_caption_idx = None
            for j in range(max(0, i - 3), i):
                tj = idx_text[j] or ""
                if _RE_TBL_ANY.match(tj):
                    existing_caption_idx = j
                    break
            if existing_caption_idx is not None:
                parsed_num, title = normalize_tbl_caption(idx_text[existing_caption_idx])
                if parsed_num:
                    ch_num, seq = parsed_num
                    tbl_counter[ch_num] = max(tbl_counter.get(ch_num, 0), seq)
                else:
                    ch_num = cur_chapter
                    seq = tbl_counter.get(ch_num, 0) + 1
                    tbl_counter[ch_num] = seq
                new_text = f"Таблица {ch_num}.{seq} — {title or '???'}"
                set_paragraph_text(children[existing_caption_idx], new_text)
                table_recs.append(
                    CaptionRec(
                        kind="table",
                        chapter=ch_num,
                        seq=seq,
                        text=new_text,
                        confidence="high" if title and title != "???" else "low",
                        source="existing_normalized",
                        old_text=idx_text[existing_caption_idx][:200],
                    )
                )
            else:
                ch_num = cur_chapter
                seq = tbl_counter.get(ch_num, 0) + 1
                tbl_counter[ch_num] = seq
                ctx_b, ctx_a = _ctx(i, 3, 3)
                # ★ не используем первый ряд как заголовок: это ячейки шапки
                # ("Параметр | Значение | ..."), а не название таблицы.
                title, conf = draft_title_from_context(ctx_b, ctx_a)
                new_text = f"Таблица {ch_num}.{seq} — {title}"
                inserts_before.append((i, make_caption_paragraph(new_text, align="left")))
                table_recs.append(
                    CaptionRec(
                        kind="table",
                        chapter=ch_num,
                        seq=seq,
                        text=new_text,
                        confidence=conf,
                        source="generated",
                        context=(ctx_b + " | next: " + ctx_a)[:200],
                    )
                )

    # Применяем вставки с хвоста
    all_inserts = [(i, "after", p) for i, p in inserts_after] + [
        (i, "before", p) for i, p in inserts_before
    ]
    all_inserts.sort(key=lambda x: x[0], reverse=True)
    for idx, pos, new_p in all_inserts:
        target = children[idx]
        if pos == "after":
            target.addnext(new_p)
        else:
            target.addprevious(new_p)

    # JSON
    Path(captions_json_path).parent.mkdir(parents=True, exist_ok=True)
    with open(captions_json_path, "w", encoding="utf-8") as f:
        json.dump(
            {
                "figures": [asdict(x) for x in figure_recs],
                "tables": [asdict(x) for x in table_recs],
            },
            f,
            ensure_ascii=False,
            indent=2,
        )

    doc.save(output_path)
    stats = {
        "figures_total": len(figure_recs),
        "figures_existing_normalized": sum(1 for r in figure_recs if r.source == "existing_normalized"),
        "figures_generated": sum(1 for r in figure_recs if r.source == "generated"),
        "figures_low_confidence": sum(1 for r in figure_recs if r.confidence == "low"),
        "tables_total": len(table_recs),
        "tables_existing_normalized": sum(1 for r in table_recs if r.source == "existing_normalized"),
        "tables_generated": sum(1 for r in table_recs if r.source == "generated"),
        "tables_low_confidence": sum(1 for r in table_recs if r.confidence == "low"),
    }
    return stats


if __name__ == "__main__":
    import sys
    if len(sys.argv) != 4:
        print("Usage: python captions.py <in.docx> <out.docx> <captions.json>")
        sys.exit(1)
    print(process_document(sys.argv[1], sys.argv[2], sys.argv[3]))
