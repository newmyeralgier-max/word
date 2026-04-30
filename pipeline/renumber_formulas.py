"""
renumber_formulas.py — перенумерация формул внутри глав по ГОСТ 7.32-2017.

Что делает:
  Pass 1. Идёт по document.xml, отслеживает текущую главу N (по pageBreakBefore +
          pattern «N. ТЕКСТ»). Считает порядковые номера формул внутри каждой
          главы. Каждый «параграф-номер» — параграф, содержащий oMath ИЛИ
          текст вида «(N.M)» в тексте — получает новый номер (chapter, ++idx).
  Pass 2. Создаёт mapping {(old_chapter, old_idx): (new_chapter, new_idx)}
          и применяет:
            • в подписях формулы заменяет «(old_chapter.old_idx)»
              на «(new_chapter.new_idx)»;
            • в inline-тексте параграфов («формуле (3.4)», «(2.7)», ...) —
              делает то же самое, при условии, что паттерн стоит вне формулы.

Ограничения:
  • Не делает перевод между главами на основе содержимого формулы — только
    по её текущему положению.
  • Если у параграфа-формулы НЕ задан номер в скобках, он не переномеруется
    (промежуточные выкладки оставляем).
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


CHAPTER_RE = re.compile(r'^\s*(\d{1,2})[\.\s]\s*[А-ЯЁA-Z]')
NUM_RE = re.compile(r'\((\d+)\.(\d+)\)')


def get_text(p):
    return ''.join(p.xpath('.//w:t/text()', namespaces=NS))


def has_pagebreak_before(p):
    return bool(p.xpath('./w:pPr/w:pageBreakBefore', namespaces=NS))


def is_toc_entry(p):
    for h in p.xpath('.//w:hyperlink', namespaces=NS):
        if h.get(qn('anchor'), '').startswith('_Toc'):
            return True
    return False


def is_inside_figure_table(p):
    """Skip paragraphs inside figure-wrapper tables."""
    parent = p.getparent()
    while parent is not None:
        if etree.QName(parent).localname == 'tbl':
            # check first row cell for drawing
            rows = parent.findall('w:tr', NS)
            if rows and rows[0].xpath('.//w:drawing', namespaces=NS):
                return True
        parent = parent.getparent()
    return False


def renumber_formula_paragraphs(doc):
    """Pass 1: walk paragraphs, renumber numbered formulas in-place.

    Returns (mapping, formula_paragraphs, chapter_counter) where:
      • mapping is a dict {(old_n, old_m): (new_n, new_m)} — only for
        UNIQUE old pairs (i.e., where the renumbering is unambiguous).
        Used later for inline-ref updates.
      • formula_paragraphs is a list of (p, old, new).
    """
    body = doc.find('w:body', NS)
    current_chapter = 0
    chapter_counter = {}
    formula_paragraphs = []

    # First pass collects ordering, then we rename in-place
    # detect chapter only via H1 paragraphs (pageBreakBefore + matching pattern)
    for p in body.iter(qn('p')):
        if is_toc_entry(p):
            continue
        if has_pagebreak_before(p) and not is_toc_entry(p):
            text = get_text(p).strip()
            m = CHAPTER_RE.match(text)
            if m:
                current_chapter = int(m.group(1))
                continue
        text = get_text(p)
        nums_in_p = NUM_RE.findall(text)
        has_omath = bool(p.findall('.//m:oMath', NS))
        if not (has_omath or nums_in_p):
            continue
        if not nums_in_p:
            continue
        old_n, old_m = int(nums_in_p[0][0]), int(nums_in_p[0][1])
        chap = current_chapter if current_chapter > 0 else old_n
        chapter_counter[chap] = chapter_counter.get(chap, 0) + 1
        new_n, new_m = chap, chapter_counter[chap]
        old_pair = (old_n, old_m)
        new_pair = (new_n, new_m)
        formula_paragraphs.append((p, old_pair, new_pair))
        # rename in-place: replace first occurrence of (old_n.old_m) → (new_n.new_m)
        if old_pair != new_pair:
            replace_first_in_runs(p, old_pair, new_pair)

    # build mapping from unique old pairs
    counts = {}
    for _, old, _ in formula_paragraphs:
        counts[old] = counts.get(old, 0) + 1
    mapping = {}
    for p_el, old, new in formula_paragraphs:
        if counts[old] == 1 and old != new:
            mapping[old] = new
    return mapping, formula_paragraphs, chapter_counter


def replace_first_in_runs(p, old_pair, new_pair):
    """Replace first occurrence of '(old_n.old_m)' with '(new_n.new_m)' in w:t texts.
    Operates only on text outside m:oMath.
    """
    # collect runs that are direct children (skip oMath nested runs)
    runs = []
    for r in p.findall('w:r', NS):
        for t in r.findall('w:t', NS):
            runs.append(t)
    if not runs:
        return False
    texts = [t.text or '' for t in runs]
    full = ''.join(texts)
    target = f'({old_pair[0]}.{old_pair[1]})'
    repl   = f'({new_pair[0]}.{new_pair[1]})'
    pos = full.find(target)
    if pos < 0:
        return False
    new_full = full[:pos] + repl + full[pos + len(target):]
    runs[0].text = new_full
    for t in runs[1:]:
        t.text = ''
    return True


def replace_in_runs(p, mapping, *, only_match_singleton=False):
    """Replace (old_n.old_m) → (new_n.new_m) in w:t texts of paragraph p, outside m:oMath.

    If only_match_singleton is True, only replace when the paragraph contains
    exactly ONE number (so we don't break formula captions that already passed
    through).
    """
    runs = []
    for r in p.findall('w:r', NS):
        for t in r.findall('w:t', NS):
            runs.append((r, t))
    if not runs:
        return 0
    # build full text + offsets
    texts = [t.text or '' for _, t in runs]
    offsets = [0]
    for s in texts:
        offsets.append(offsets[-1] + len(s))
    full = ''.join(texts)
    # find matches and produce replacements
    matches = list(NUM_RE.finditer(full))
    if not matches:
        return 0
    if only_match_singleton and len(matches) != 1:
        return 0
    # apply in reverse
    fixed = 0
    new_full = full
    for m in reversed(matches):
        old_pair = (int(m.group(1)), int(m.group(2)))
        if old_pair not in mapping:
            continue
        new_n, new_m = mapping[old_pair]
        repl = f'({new_n}.{new_m})'
        new_full = new_full[:m.start()] + repl + new_full[m.end():]
        fixed += 1
    if not fixed:
        return 0
    # redistribute new_full back into run texts (proportional to original lengths)
    # simplest: stuff entire new text into first w:t, clear others
    # To preserve positions/styles approximately, do per-run reassignment by
    # walking original offsets but limiting new-string length per run.
    # Use simpler strategy: just shove everything into the first run, clear rest.
    runs[0][1].text = new_full
    for _, t in runs[1:]:
        t.text = ''
    return fixed


def apply_mapping(doc, mapping):
    body = doc.find('w:body', NS)
    total_fixed = 0
    for p in body.iter(qn('p')):
        if is_toc_entry(p):
            continue
        if is_inside_figure_table(p):
            continue
        # apply in this paragraph (singleton-strict mode for formula captions,
        # multi-mode for body text)
        text = get_text(p)
        nums = NUM_RE.findall(text)
        if not nums:
            continue
        # for paragraphs with multiple numbers — do replace all (text refs)
        # use replace_in_runs default
        fixed = replace_in_runs(p, mapping)
        total_fixed += fixed
    return total_fixed


def process(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    doc = etree.fromstring(data['word/document.xml'])
    mapping, formula_paragraphs, chapter_counter = renumber_formula_paragraphs(doc)
    print(f'Found {len(formula_paragraphs)} numbered formulas, '
          f'unique-mapping size: {len(mapping)}')
    print('Per-chapter counts:')
    for c, k in sorted(chapter_counter.items()):
        print(f'  ch{c}: {k} formulas')
    print('All paragraph remappings:')
    for _, old, new in formula_paragraphs:
        if old != new:
            print(f'  ({old[0]}.{old[1]}) → ({new[0]}.{new[1]})')

    # apply mapping to entire document for inline references
    fixed = apply_mapping(doc, mapping)
    print(f'Inline reference updates: {fixed}')

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
