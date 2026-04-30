"""
turbo_renumber.py — robust per-chapter formula renumbering with cleanup.

Algorithm:
  1. Walk all paragraphs in document order, tracking current_chapter via
     pageBreakBefore + "N. ТЕКСТ" pattern (skip TOC entries).
  2. For each paragraph that has m:oMath AND has at least one (X.Y) caption
     in its w:t text → it is a "numbered formula".
     - Take the LAST (X.Y) in w:t as the caption number (most reliable —
       caption is typically at end after tab).
     - Assign new number = (current_chapter, ++counter[chapter]).
     - Replace the LAST occurrence of (X.Y) in w:t runs with new number.
       Replace ALL other (X.Y) patterns in the paragraph w:t with the SAME
       new number (cleanup of duplicates and stale references in the
       same paragraph).
  3. Build mapping {old_caption_pair → new_caption_pair} ONLY for unique
     old captions (where old appears as caption exactly once across all
     paragraphs). Use that mapping in a 2nd pass to update inline references
     in paragraphs that are NOT numbered-formula paragraphs.
"""

import argparse
import re
import zipfile
from pathlib import Path
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
    parent = p.getparent()
    while parent is not None:
        if etree.QName(parent).localname == 'tbl':
            rows = parent.findall('w:tr', NS)
            if rows and rows[0].xpath('.//w:drawing', namespaces=NS):
                return True
        parent = parent.getparent()
    return False


def collect_text_runs(p):
    """Return list of w:t elements that are direct children of w:r elements
    of paragraph p (NOT inside m:oMath or other nested structures)."""
    out = []
    for r in p.findall('w:r', NS):
        for t in r.findall('w:t', NS):
            out.append(t)
    return out


def replace_all_in_text_runs(p, old_pair, new_pair):
    """Replace ALL occurrences of (old_n.old_m) → (new_n.new_m) in paragraph
    text runs (outside m:oMath)."""
    runs = collect_text_runs(p)
    if not runs:
        return 0
    target = f'({old_pair[0]}.{old_pair[1]})'
    repl = f'({new_pair[0]}.{new_pair[1]})'
    fixed = 0
    for t in runs:
        cur = t.text or ''
        if target in cur:
            new_cur = cur.replace(target, repl)
            t.text = new_cur
            fixed += cur.count(target)
    return fixed


def replace_last_in_text_runs(p, old_pair, new_pair):
    """Replace ONLY the LAST occurrence of (old_n.old_m) → (new_n.new_m).
    Useful when paragraph has multiple references but only the caption
    (last one) should be changed."""
    runs = collect_text_runs(p)
    if not runs:
        return False
    target = f'({old_pair[0]}.{old_pair[1]})'
    repl = f'({new_pair[0]}.{new_pair[1]})'
    # walk from last run backwards
    for t in reversed(runs):
        cur = t.text or ''
        if target in cur:
            idx = cur.rfind(target)
            t.text = cur[:idx] + repl + cur[idx + len(target):]
            return True
    return False


def consolidate_caption_runs(p, new_pair):
    """Force the caption to be exactly '(new_n.new_m)' in the LAST text run
    that contains a number, removing all OTHER (X.Y) patterns in the
    paragraph (cleanup of duplicates, stale references that ended up in
    the same paragraph as the caption)."""
    runs = collect_text_runs(p)
    if not runs:
        return 0
    fixed = 0
    # Find last run containing ANY (X.Y) pattern
    last_run_idx = None
    for i, t in enumerate(runs):
        if NUM_RE.search(t.text or ''):
            last_run_idx = i
    if last_run_idx is None:
        return 0
    # Replace ALL (X.Y) in ALL runs:
    #   - in the LAST run, the LAST (X.Y) becomes new caption
    #   - all OTHER (X.Y) (in last run and in other runs) get removed
    target_repl = f'({new_pair[0]}.{new_pair[1]})'
    for i, t in enumerate(runs):
        cur = t.text or ''
        if i == last_run_idx:
            # replace LAST (X.Y) in this run with new caption, drop others
            matches = list(NUM_RE.finditer(cur))
            if not matches:
                continue
            # build new text by walking matches: drop all except the last,
            # which becomes the caption
            last_match = matches[-1]
            new_cur = ''
            cursor = 0
            for m in matches:
                new_cur += cur[cursor:m.start()]
                if m is last_match:
                    new_cur += target_repl
                # else: drop (don't append the match text)
                cursor = m.end()
            new_cur += cur[cursor:]
            if new_cur != cur:
                t.text = new_cur
                fixed += 1
        else:
            # drop all (X.Y) patterns from this run
            new_cur = NUM_RE.sub('', cur)
            if new_cur != cur:
                t.text = new_cur
                fixed += 1
    return fixed


def renumber_formulas_v2(doc):
    body = doc.find('w:body', NS)
    current_chapter = 0
    chapter_counter = {}
    formula_paragraphs = []  # list of (p, old_pair, new_pair)
    for p in body.iter(qn('p')):
        if is_toc_entry(p):
            continue
        if has_pagebreak_before(p):
            text = get_text(p).strip()
            m = CHAPTER_RE.match(text)
            if m:
                current_chapter = int(m.group(1))
                continue
        # only paragraphs inside chapters 1..N (skip TOC, intro, ВВЕДЕНИЕ)
        if current_chapter == 0:
            continue
        # detect formula paragraph: has m:oMath AND has at least one (X.Y) in w:t
        text = get_text(p)
        nums = NUM_RE.findall(text)
        has_omath = bool(p.findall('.//m:oMath', NS))
        if not has_omath:
            continue
        if not nums:
            continue
        # Take LAST number as the original caption (most reliable)
        old_n, old_m = int(nums[-1][0]), int(nums[-1][1])
        chap = current_chapter
        chapter_counter[chap] = chapter_counter.get(chap, 0) + 1
        new_n, new_m = chap, chapter_counter[chap]
        old_pair = (old_n, old_m)
        new_pair = (new_n, new_m)
        formula_paragraphs.append((p, old_pair, new_pair))
        # consolidate caption runs: drop duplicates, keep last as new_pair
        consolidate_caption_runs(p, new_pair)

    # build mapping for inline reference updates (only unique old captions)
    counts = {}
    for _, old, _ in formula_paragraphs:
        counts[old] = counts.get(old, 0) + 1
    mapping = {}
    for _, old, new in formula_paragraphs:
        if counts[old] == 1 and old != new:
            mapping[old] = new
    return mapping, formula_paragraphs, chapter_counter


def update_inline_refs(doc, mapping, formula_paragraphs):
    """Replace (old_n.old_m) → (new_n.new_m) in paragraphs that are NOT
    in the formula_paragraphs list (i.e., body text references)."""
    formula_p_set = set(id(p) for p, _, _ in formula_paragraphs)
    body = doc.find('w:body', NS)
    fixed = 0
    for p in body.iter(qn('p')):
        if id(p) in formula_p_set:
            continue
        if is_toc_entry(p):
            continue  # don't touch TOC; updateFields handles it
        if is_inside_figure_table(p):
            continue
        # find all (X.Y) and replace if in mapping
        runs = collect_text_runs(p)
        if not runs:
            continue
        for t in runs:
            cur = t.text or ''
            if not NUM_RE.search(cur):
                continue
            new_cur = cur
            for old_pair, new_pair in mapping.items():
                target = f'({old_pair[0]}.{old_pair[1]})'
                repl = f'({new_pair[0]}.{new_pair[1]})'
                if target in new_cur:
                    new_cur = new_cur.replace(target, repl)
            if new_cur != cur:
                t.text = new_cur
                fixed += 1
    return fixed


def process(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    doc = etree.fromstring(data['word/document.xml'])
    mapping, formula_paragraphs, chapter_counter = renumber_formulas_v2(doc)
    print(f'Found {len(formula_paragraphs)} numbered formulas')
    print('Per-chapter counts:')
    for c, k in sorted(chapter_counter.items()):
        print(f'  ch{c}: {k} formulas')
    print('Renames:')
    for _, old, new in formula_paragraphs:
        if old != new:
            print(f'  ({old[0]}.{old[1]}) → ({new[0]}.{new[1]})')

    fixed = update_inline_refs(doc, mapping, formula_paragraphs)
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
