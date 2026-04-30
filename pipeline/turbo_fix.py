"""
turbo_fix.py — комплексное исправление всех найденных визуальных проблем
после прохода первого pipeline.

Что чинит:
1. **Пустая страница после титульника**: убирает `pageBreakBefore` с пустого
   параграфа idx≈63 (тот, что между концом титула и СОДЕРЖАНИЕ).
2. **Дубли формул `(N.M)(N.M)` в одном параграфе**: проходит ВСЕ параграфы,
   находит подряд идущие пары и удаляет вторую (multiple passes).
3. **«ПРОПАЛА ФОРМУЛА» внутри m:oMath**: удаляет такой плейсхолдер целиком
   (вместе с обёрткой oMath, чтобы не оставлять пустую формулу).
4. **TOC stale entries**: заменяет «10.1 Моделирование рабочих режимов» на
   «10.2 …» в текстовом содержимом TOC-полей (но не пересчитывает их —
   полное обновление поля делается в Word через Update Field).
5. **w:settings → updateFields**: добавляет `<w:updateFields w:val="true"/>`,
   чтобы Word при открытии автоматически обновил TOC и PAGE-поля.
"""

import argparse
import re
import zipfile
from pathlib import Path
from copy import deepcopy
from lxml import etree

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}


def qn(tag, ns='w'):
    return f'{{{NS[ns]}}}{tag}'


# --- 1) Empty paragraph with pageBreakBefore -------------------------

def fix_empty_pagebreak_paragraphs(doc):
    """Remove the entire empty paragraph that has pageBreakBefore (causes blank page).
    Only removes if next sibling paragraph also has pageBreakBefore (so no break is lost).
    """
    body = doc.find('w:body', NS)
    paragraphs_to_remove = []
    children = list(body)
    for i, el in enumerate(children[:-1]):
        if etree.QName(el).localname != 'p':
            continue
        # find next paragraph sibling
        next_p = None
        for j in range(i + 1, len(children)):
            if etree.QName(children[j]).localname == 'p':
                next_p = children[j]
                break
        if next_p is None:
            continue
        pPr = el.find('w:pPr', NS)
        if pPr is None:
            continue
        pbb = pPr.find('w:pageBreakBefore', NS)
        if pbb is None:
            continue
        # check paragraph empty (no text in any w:t / m:t / drawing)
        text = ''.join(el.xpath('.//w:t/text()', namespaces=NS))
        mtext = ''.join(el.xpath('.//m:t/text()', namespaces=NS))
        has_drawing = bool(el.xpath('.//w:drawing', namespaces=NS))
        has_object = bool(el.xpath('.//w:object', namespaces=NS))
        if text.strip() or mtext.strip() or has_drawing or has_object:
            continue
        # only remove if next paragraph ALSO has pageBreakBefore (so no break is lost)
        next_pbb = bool(next_p.xpath('./w:pPr/w:pageBreakBefore', namespaces=NS))
        if not next_pbb:
            # just remove the pageBreakBefore from this empty paragraph
            pPr.remove(pbb)
            paragraphs_to_remove.append(('cleared_pbb', el))
        else:
            # remove entire paragraph (next has pbb so page break is preserved)
            paragraphs_to_remove.append(('removed', el))
    for action, el in paragraphs_to_remove:
        if action == 'removed':
            el.getparent().remove(el)
    return len(paragraphs_to_remove)


# --- 2) Duplicate formula numbers (N.M)(N.M) -------------------------

NUM_RE = re.compile(r'\((\d+)[\.,](\d+)\)')


def collect_text_runs(p):
    out = []
    for r in p.findall('w:r', NS):
        for t in r.findall('w:t', NS):
            out.append((r, t, t.text or ''))
    return out


def remove_duplicate_numbers_in_paragraph(p):
    """Find occurrences of (N.M)(N.M) (separated only by whitespace) and remove the second."""
    runs = collect_text_runs(p)
    if not runs:
        return 0
    full = ''.join(rt for _, _, rt in runs)
    matches = list(NUM_RE.finditer(full))
    if len(matches) < 2:
        return 0
    to_remove_ranges = []
    seen_prev = None
    for m in matches:
        if seen_prev:
            prev_key, prev_match = seen_prev
            cur_key = (m.group(1), m.group(2))
            if prev_key == cur_key:
                gap = full[prev_match.end():m.start()]
                if re.fullmatch(r'\s*', gap):
                    to_remove_ranges.append((m.start(), m.end()))
                    seen_prev = ((m.group(1), m.group(2)), m)
                    continue
        seen_prev = ((m.group(1), m.group(2)), m)

    if not to_remove_ranges:
        return 0

    # Apply removals — go in reverse to keep offsets stable
    for (rs, re_) in sorted(to_remove_ranges, reverse=True):
        runs = collect_text_runs(p)
        offsets = [0]
        for _, _, rt in runs:
            offsets.append(offsets[-1] + len(rt))
        for i, (_, t_el, _) in enumerate(runs):
            run_start = offsets[i]
            run_end = offsets[i + 1]
            if run_start < re_ and run_end > rs:
                local_s = max(0, rs - run_start)
                local_e = min(run_end - run_start, re_ - run_start)
                cur = t_el.text or ''
                t_el.text = cur[:local_s] + cur[local_e:]
    return len(to_remove_ranges)


def fix_duplicate_numbers_all(doc):
    """Run the dedup pass repeatedly until no more dups found."""
    body = doc.find('w:body', NS)
    total = 0
    for _ in range(5):  # up to 5 passes
        round_total = 0
        for p in body.iter(qn('p')):
            round_total += remove_duplicate_numbers_in_paragraph(p)
        total += round_total
        if round_total == 0:
            break
    return total


# --- 3) Remove "ПРОПАЛА ФОРМУЛА" placeholder oMath -------------------

def remove_propala_formula_placeholders(doc):
    """Find m:oMath that contains only «ПРОПАЛА ФОРМУЛА» text and remove it."""
    fixed = 0
    for omath in doc.iter(qn('oMath', 'm')):
        text = ''.join(omath.xpath('.//m:t/text()', namespaces=NS))
        if 'ПРОПАЛА' in text or 'пропала' in text.lower():
            parent = omath.getparent()
            parent.remove(omath)
            fixed += 1
    return fixed


# --- 4) TOC stale entries --------------------------------------------

TOC_REPLACEMENTS = [
    # (old_text_substring, new_text) — only specific replacements that should
    # affect SOURCE heading text. Values that appear in TOC are auto-updated
    # by Word once updateFields=true is set in settings.xml.
    ('10.1 Моделирование рабочих режимов', '10.2 Моделирование рабочих режимов'),
    ('10.1.1 Параметры питающей энергосистемы', '10.2.1 Параметры питающей энергосистемы'),
    ('10.1.2 Параметры силовых трансформаторов', '10.2.2 Параметры силовых трансформаторов'),
    ('10.1.3 Параметры кабельных линий', '10.2.3 Параметры кабельных линий'),
    ('10.1.4 Эквивалентные нагрузки', '10.2.4 Эквивалентные нагрузки'),
    ('10.1.5 Специфические настройки', '10.2.5 Специфические настройки'),
    ('10.2 Анализ результатов моделирования', '10.3 Анализ результатов моделирования'),
    ('10.3 Моделирование режимов с учетом несовпадения', '10.4 Моделирование режимов с учетом несовпадения'),
]


def fix_toc_stale_entries(doc):
    """Replace stale TOC text entries WITHOUT breaking dot-leader/page-number
    structure. We only replace text in the run-sequence within a hyperlink's
    text portion (everything before the first w:tab).
    """
    body = doc.find('w:body', NS)
    fixed = 0
    for hyp in body.iter(qn('hyperlink')):
        # Collect runs and tab positions
        run_seq = []
        for child in hyp.iter():
            if etree.QName(child).localname == 'r' and etree.QName(child).namespace == NS['w']:
                # Order matters; iter yields document order
                run_seq.append(child)
        # find the first run that contains a w:tab
        before_tab = []
        for r in run_seq:
            tabs = r.findall('w:tab', NS)
            if tabs:
                # this run begins/contains tab — stop collecting
                break
            before_tab.append(r)
        if not before_tab:
            continue
        # collect text-only of these runs
        ts_per_run = []
        for r in before_tab:
            ts = r.findall('w:t', NS)
            ts_per_run.append(ts)
        flat_ts = [t for ts in ts_per_run for t in ts]
        if not flat_ts:
            continue
        full = ''.join(t.text or '' for t in flat_ts)
        new_full = full
        for old, new in TOC_REPLACEMENTS:
            if old in new_full:
                new_full = new_full.replace(old, new)
        if new_full == full:
            continue
        # write new_full into first w:t, clear others
        flat_ts[0].text = new_full
        for t in flat_ts[1:]:
            t.text = ''
        fixed += 1
    return fixed


# --- 5) settings.xml: add updateFields ---------------------------------

def fix_settings_update_fields(zip_data):
    """Patch word/settings.xml to include <w:updateFields w:val="true"/>"""
    if 'word/settings.xml' not in zip_data:
        return False
    raw = zip_data['word/settings.xml']
    settings = etree.fromstring(raw)
    if settings.find('w:updateFields', NS) is not None:
        return False
    el = etree.SubElement(settings, qn('updateFields'))
    el.set(qn('val'), 'true')
    zip_data['word/settings.xml'] = etree.tostring(
        settings, xml_declaration=True, encoding='UTF-8', standalone=True
    )
    return True


# --- main ---

def process(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    doc = etree.fromstring(data['word/document.xml'])

    n_blank = fix_empty_pagebreak_paragraphs(doc)
    print(f'Empty paragraphs with pageBreakBefore cleared: {n_blank}')

    n_dups = fix_duplicate_numbers_all(doc)
    print(f'Duplicate (N.M) numbers removed: {n_dups}')

    n_propala = remove_propala_formula_placeholders(doc)
    print(f'«ПРОПАЛА ФОРМУЛА» placeholders removed: {n_propala}')

    n_toc = fix_toc_stale_entries(doc)
    print(f'TOC stale entries replaced: {n_toc}')

    n_settings = fix_settings_update_fields(data)
    print(f'settings.xml updateFields injected: {n_settings}')

    data['word/document.xml'] = etree.tostring(
        doc, xml_declaration=True, encoding='UTF-8', standalone=True
    )
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, content in data.items():
            zout.writestr(n, content)


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--output', required=True)
    args = ap.parse_args()
    process(Path(args.input), Path(args.output))
