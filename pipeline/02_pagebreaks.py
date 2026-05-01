"""
fix_pagebreaks.py — каждая глава с новой страницы (ГОСТ 7.32-2017 §6.2).

Стратегия:
  • Заголовки H1 определяются по тексту:
      – «N. ТЕКСТ» / «N.ТЕКСТ» / «N ТЕКСТ» где N — целое 1..99 (главы);
      – структурные элементы (СОДЕРЖАНИЕ, ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ,
        СПИСОК ЛИТЕРАТУРЫ, СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ).
  • Каждому такому параграфу добавляется <w:pageBreakBefore/> в pPr.
  • Удаляются «парные» ручные <w:br w:type="page"/> в предыдущих
    параграфах, чтобы не получалось пустой страницы.
  • Параграфы титульной зоны (до первого BR-PAGE и/или СОДЕРЖАНИЕ) не трогаем.
"""

import argparse
import re
import shutil
import zipfile
from pathlib import Path
from copy import deepcopy
from lxml import etree

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}
W = NS['w']


def qn(tag):
    return f'{{{W}}}{tag}'


STRUCTURAL = {
    'СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ',
    'ВВЕДЕНИЕ',
    'ЗАКЛЮЧЕНИЕ',
    'СПИСОК ЛИТЕРАТУРЫ',
    'СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ',
    'СПИСОК ИСПОЛЬЗОВАННОЙ ЛИТЕРАТУРЫ',
}

# H1 pattern: «1 ТЕКСТ ...», «1. ТЕКСТ ...», «1.ТЕКСТ ...»
RE_CHAPTER = re.compile(r'^\s*(\d{1,2})[\.\s]\s*([А-ЯЁ][А-ЯЁ\s\-—,A-Z0-9]{2,})')

# Pattern for TOC entries: «<title>....<page-number>» — ends with digit/page
RE_TOC_TAIL = re.compile(r'\s+\d{1,3}\s*$')


def get_paragraph_text(p):
    return ''.join(p.xpath('.//w:t/text()', namespaces=NS)).strip()


def is_toc_entry(p):
    """A TOC entry contains a hyperlink to a _Toc bookmark."""
    return bool(p.xpath('.//w:hyperlink[@w:anchor]', namespaces=NS)) and \
           any(h.get(qn('anchor'), '').startswith('_Toc')
               for h in p.xpath('.//w:hyperlink', namespaces=NS))


def is_h1(text: str) -> bool:
    t = text.strip()
    if not t:
        return False
    if t.upper() in STRUCTURAL or any(t.upper().startswith(s) for s in STRUCTURAL):
        return True
    m = RE_CHAPTER.match(t)
    if not m:
        return False
    # avoid matching list items: «1) ...», «1.1 ...», «1. Подраздел ...»
    if re.match(r'^\s*\d+\.\d', t):
        return False
    body = m.group(2)
    # require sufficient uppercase ratio
    upper = sum(1 for ch in body if ch.isupper())
    letters = sum(1 for ch in body if ch.isalpha())
    if letters == 0:
        return False
    if upper / letters < 0.7:
        return False
    return True


def ensure_pPr(p):
    pPr = p.find('w:pPr', NS)
    if pPr is None:
        pPr = etree.SubElement(p, qn('pPr'))
        p.insert(0, pPr)
    return pPr


def add_page_break_before(p):
    pPr = ensure_pPr(p)
    if pPr.find('w:pageBreakBefore', NS) is None:
        pbb = etree.SubElement(pPr, qn('pageBreakBefore'))
    # remove any contradictory <w:keepNext/> that could conflict — leave alone


def remove_manual_break_in(p):
    """Remove all <w:br w:type='page'/> inside paragraph p (but keep paragraph)."""
    removed = 0
    for br in p.xpath('.//w:br[@w:type="page"]', namespaces=NS):
        br.getparent().remove(br)
        removed += 1
    return removed


def process(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    doc = etree.fromstring(data['word/document.xml'])
    body = doc.find('w:body', NS)

    # find first BR-PAGE in body, treat preceding paragraphs as title page (skip)
    children = list(body)
    first_br_idx = -1
    for i, el in enumerate(children):
        if etree.QName(el).localname == 'p':
            if el.xpath('.//w:br[@w:type="page"]', namespaces=NS):
                first_br_idx = i
                break

    # find TOC zone: from СОДЕРЖАНИЕ paragraph up to next H1 (ВВЕДЕНИЕ)
    toc_start, toc_end = -1, -1
    for i, el in enumerate(children):
        if etree.QName(el).localname != 'p':
            continue
        text = get_paragraph_text(el)
        if toc_start == -1 and text.upper().startswith('СОДЕРЖАНИЕ') and len(text) <= 20:
            toc_start = i
            continue
        if toc_start != -1 and toc_end == -1:
            if text.upper() == 'ВВЕДЕНИЕ' or RE_CHAPTER.match(text):
                # this is a real heading after TOC — but only if it doesn't have a TOC tail
                if not RE_TOC_TAIL.search(text):
                    toc_end = i
                    break

    h1_count = 0
    removed_manual = 0
    titles_logged = []
    for i, el in enumerate(children):
        if etree.QName(el).localname != 'p':
            continue
        if i <= first_br_idx:
            # in title-page area — only allow structural keywords if found here (rare)
            text = get_paragraph_text(el)
            if text.upper() in STRUCTURAL:
                add_page_break_before(el)
                # remove br in preceding p
                if i > 0:
                    prev = children[i - 1]
                    if etree.QName(prev).localname == 'p':
                        removed_manual += remove_manual_break_in(prev)
                titles_logged.append(text)
                h1_count += 1
            continue
        # skip TOC zone
        if toc_start != -1 and toc_end != -1 and toc_start < i < toc_end:
            continue
        # skip any TOC entry (hyperlink to _Toc bookmark)
        if is_toc_entry(el):
            continue
        text = get_paragraph_text(el)
        if is_h1(text):
            add_page_break_before(el)
            # remove preceding manual page break (which would create blank page)
            if i > 0:
                prev = children[i - 1]
                if etree.QName(prev).localname == 'p':
                    prev_txt = get_paragraph_text(prev)
                    # only remove if previous paragraph is empty / very short
                    # to avoid losing breaks that separate real content
                    if len(prev_txt) < 5:
                        removed_manual += remove_manual_break_in(prev)
            titles_logged.append(text[:80])
            h1_count += 1

    data['word/document.xml'] = etree.tostring(
        doc, xml_declaration=True, encoding='UTF-8', standalone=True
    )

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, content in data.items():
            zout.writestr(n, content)

    print(f'H1 paragraphs marked: {h1_count}')
    print(f'Manual BR-PAGE removed (deduplicated): {removed_manual}')
    for t in titles_logged:
        print(f'  H1 → {t}')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--output', required=True)
    args = ap.parse_args()
    process(Path(args.input), Path(args.output))
