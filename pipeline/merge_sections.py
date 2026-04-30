"""
merge_sections.py — объединение раздела 9 в начало раздела 10.

Согласно заданию ПРАВИЛЬНО3:
  «9 раздел нужно объединить с 10 разделом, потому что 9 раздел нужно полностью
  переделать в 10 раздел».

Алгоритм:
  1. Находим заголовок «9. МОДЕЛИРОВАНИЕ АВТОМАТИЗИРОВАННОЙ ГОРОДСКОЙ СЕТИ»
     с pageBreakBefore. Запоминаем его индекс h9.
  2. Находим заголовок «10. МОДЕЛИРОВАНИЕ ГОРОДСКОЙ МАГИСТРАЛЬНОЙ
     АВТОМАТИЗИРОВАННОЙ ЛИНИИ». Запоминаем его индекс h10.
  3. Удаляем заголовок раздела 9 (h9).
  4. Содержимое (h9+1 .. h10-1) перемещаем сразу после h10 (становится
     введением раздела 10), сохраняя относительный порядок.
  5. Дополнительно — починка дубликата «10.1»: ищем два H2 с префиксом
     «10.1 …», второй переименовываем в «10.2 …». Все нижестоящие
     «10.1.K» (которые относятся ко второму) → «10.2.K».

Перенумерация рисунков/таблиц 9.x → 10.x делается ОТДЕЛЬНО скриптом
renumber_refs.py.
"""

import argparse
import re
import zipfile
from pathlib import Path
from copy import deepcopy
from lxml import etree

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}


def qn(tag):
    return f'{{{NS["w"]}}}{tag}'


CHAPTER_RE = re.compile(r'^\s*(\d{1,2})[\.\s]\s*[А-ЯЁA-Z]')


def get_text(p):
    return ''.join(p.xpath('.//w:t/text()', namespaces=NS)).strip()


def has_pagebreak_before(p):
    return bool(p.xpath('./w:pPr/w:pageBreakBefore', namespaces=NS))


def is_toc_entry(p):
    for h in p.xpath('.//w:hyperlink', namespaces=NS):
        if h.get(qn('anchor'), '').startswith('_Toc'):
            return True
    return False


def find_chapter_heading(body, chapter_num):
    """Find body-direct-child paragraph that is the H1 for the given chapter."""
    for i, el in enumerate(body):
        if etree.QName(el).localname != 'p':
            continue
        if not has_pagebreak_before(el):
            continue
        if is_toc_entry(el):
            continue
        text = get_text(el)
        m = CHAPTER_RE.match(text)
        if m and int(m.group(1)) == chapter_num:
            return i, el
    return None, None


def replace_text_in_runs(p, old_text, new_text):
    """Replace first occurrence of `old_text` with `new_text` in joined runs of p."""
    runs = []
    for r in p.findall('w:r', NS):
        for t in r.findall('w:t', NS):
            runs.append(t)
    if not runs:
        return False
    full = ''.join(t.text or '' for t in runs)
    pos = full.find(old_text)
    if pos < 0:
        return False
    new_full = full[:pos] + new_text + full[pos + len(old_text):]
    runs[0].text = new_full
    for t in runs[1:]:
        t.text = ''
    return True


def fix_duplicate_subsection_101(body, h10_idx):
    """Fix repeated '10.1' subsections.

    Walk children from h10_idx forward. Collect paragraphs starting with '10.1 '
    or '10.1\\xa0'. If 2+ — the second & beyond get re-prefixed.
    """
    h2_re = re.compile(r'^\s*10\.1(\D)')  # matches 10.1 followed by non-digit
    h3_re = re.compile(r'^\s*10\.1\.(\d)')  # 10.1.1 etc.

    # find all H2 candidates in chapter 10
    h2_indices = []
    for j in range(h10_idx + 1, len(body)):
        el = body[j]
        if etree.QName(el).localname != 'p':
            continue
        text = get_text(el)
        if h2_re.match(text):
            h2_indices.append(j)
    if len(h2_indices) < 2:
        return 0
    fixed = 0
    # for each duplicate (2nd and beyond), renumber
    for k, idx in enumerate(h2_indices[1:], start=2):
        el = body[idx]
        text = get_text(el)
        new_text = re.sub(r'^(\s*)10\.1', rf'\g<1>10.{k}', text, count=1)
        if replace_text_in_runs(el, text, new_text):
            fixed += 1
    # also fix H3-style "10.1.K" if they belong to 2nd, 3rd... H2 zone
    # (between successive H2 indices)
    for k, idx in enumerate(h2_indices[1:], start=2):
        zone_end = h2_indices[h2_indices.index(idx) + 1] if h2_indices.index(idx) + 1 < len(h2_indices) else len(body)
        for j in range(idx + 1, zone_end):
            el = body[j]
            if etree.QName(el).localname != 'p':
                continue
            text = get_text(el)
            m = h3_re.match(text)
            if m:
                new_text = re.sub(r'^(\s*)10\.1\.', rf'\g<1>10.{k}.', text, count=1)
                if replace_text_in_runs(el, text, new_text):
                    fixed += 1
    return fixed


def process(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    doc = etree.fromstring(data['word/document.xml'])
    body = doc.find('w:body', NS)

    h9_idx, h9 = find_chapter_heading(body, 9)
    h10_idx, h10 = find_chapter_heading(body, 10)

    if h9 is None or h10 is None:
        print(f'Section 9 found: {h9 is not None}, section 10 found: {h10 is not None} — skip merge')
        merge_count = 0
    elif h10_idx <= h9_idx:
        print(f'Section 10 is before 9 (h9={h9_idx}, h10={h10_idx}) — skip')
        merge_count = 0
    else:
        # children to move = body[h9_idx+1 .. h10_idx-1]
        # remove section 9 heading first, then re-resolve indices
        section9_content = []
        # take snapshot of children references
        children = list(body)
        for el in children[h9_idx + 1:h10_idx]:
            section9_content.append(el)
        merge_count = len(section9_content)
        # remove section 9 heading
        body.remove(h9)
        # remove the moved content from current position
        for el in section9_content:
            body.remove(el)
        # insert after h10
        # find new index of h10
        new_h10_idx = list(body).index(h10)
        for offset, el in enumerate(section9_content):
            body.insert(new_h10_idx + 1 + offset, el)
        # Remove pageBreakBefore from any of these moved paragraphs (since they
        # are now mid-chapter). Also remove from any leading empty paragraph.
        for el in section9_content:
            if etree.QName(el).localname == 'p':
                pPr = el.find('w:pPr', NS)
                if pPr is not None:
                    pbb = pPr.find('w:pageBreakBefore', NS)
                    if pbb is not None:
                        pPr.remove(pbb)

    # Fix duplicate 10.1 subsection numbering
    if h10 is not None:
        h10_idx_now = list(body).index(h10)
        fixes = fix_duplicate_subsection_101(body, h10_idx_now)
    else:
        fixes = 0

    print(f'Section 9 paragraphs/tables merged into section 10: {merge_count}')
    print(f'Subsection 10.x renumber fixes: {fixes}')

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
