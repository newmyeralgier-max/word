"""
fix_subscripts.py — конвертирует «P_а»-подобные обозначения в нижний индекс Word.

Что делает:
  В каждом w:t тексте ищет паттерн `<base>_<sub>` где:
    • base — короткая последовательность букв/цифр (1-4 символа), оканчивается
      буквой;
    • sub  — короткая последовательность букв/цифр (1-4 символа);
    • вокруг — границы слова (\b или конец строки/пробел/пунктуация).

  Каждое совпадение разбивается на:
    • левый текстовый run: текст до base + сам base;
    • правый текстовый run с <w:vertAlign w:val="subscript"/>: sub.
  Подчёркивание удаляется.

  Шрифт и прочие свойства rPr копируются из исходного run.
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
W = NS['w']


def qn(tag):
    return f'{{{W}}}{tag}'


# Tighter pattern: avoid matching long compound names like R_ohl_def.
# base = 1-4 alnum (letter or digit), MUST end with letter
# sub  = 1-4 alnum, MUST start with letter or digit
# Boundary: preceded by start/space/punct, followed by end/space/punct (NOT _)
SUB_RE = re.compile(
    r'(?P<lead>(?:^|[\s\(\[\{,;:.])'  # boundary on left
    r')(?P<base>[A-Za-zА-Яа-я]{1,4})'
    r'_'
    r'(?P<sub>[A-Za-zА-Яа-я0-9]{1,4})'
    r'(?P<trail>(?=[\s\)\]\},;:.\-—=]|$))'
)


def split_run_for_subscript(r_el):
    """
    If text in w:r contains 'X_y' patterns, split this run into multiple runs.
    Returns list of new <w:r> elements (or None if no change).
    """
    t_el = r_el.find('w:t', NS)
    if t_el is None or not t_el.text:
        return None
    text = t_el.text
    # collect matches
    matches = list(SUB_RE.finditer(text))
    if not matches:
        return None

    rPr_orig = r_el.find('w:rPr', NS)

    # build a list of (text, is_subscript)
    pieces = []
    cursor = 0
    for m in matches:
        # text before base (including the boundary char captured in 'lead')
        before = text[cursor:m.start('base')]
        if before:
            pieces.append((before, False))
        pieces.append((m.group('base'), False))
        pieces.append((m.group('sub'), True))
        cursor = m.end('sub')
    # tail
    if cursor < len(text):
        pieces.append((text[cursor:], False))

    # build new runs
    new_runs = []
    for txt, is_sub in pieces:
        if not txt:
            continue
        nr = etree.Element(qn('r'))
        # copy rPr
        if rPr_orig is not None:
            new_rPr = deepcopy(rPr_orig)
        else:
            new_rPr = etree.Element(qn('rPr'))
        if is_sub:
            # add <w:vertAlign w:val="subscript"/>
            va = new_rPr.find('w:vertAlign', NS)
            if va is None:
                va = etree.SubElement(new_rPr, qn('vertAlign'))
            va.set(qn('val'), 'subscript')
        nr.append(new_rPr)
        nt = etree.SubElement(nr, qn('t'))
        # preserve leading/trailing spaces
        if txt != txt.strip():
            nt.set(
                '{http://www.w3.org/XML/1998/namespace}space',
                'preserve',
            )
        nt.text = txt
        new_runs.append(nr)
    return new_runs


def process(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    doc = etree.fromstring(data['word/document.xml'])

    converted = 0
    # iterate all paragraphs (top-level + within tables)
    for p in doc.iter(qn('p')):
        # skip paragraphs inside <m:oMath> — formulas already typeset
        if p.xpath('ancestor::m:oMath', namespaces={'m':'http://schemas.openxmlformats.org/officeDocument/2006/math'}):
            continue
        runs = list(p.findall('w:r', NS))
        for r in runs:
            new_runs = split_run_for_subscript(r)
            if new_runs:
                # replace r with new_runs in-place
                parent = r.getparent()
                idx = list(parent).index(r)
                parent.remove(r)
                for offset, nr in enumerate(new_runs):
                    parent.insert(idx + offset, nr)
                converted += 1

    data['word/document.xml'] = etree.tostring(
        doc, xml_declaration=True, encoding='UTF-8', standalone=True
    )
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, content in data.items():
            zout.writestr(n, content)
    print(f'Subscript-style runs converted: {converted}')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--output', required=True)
    args = ap.parse_args()
    process(Path(args.input), Path(args.output))
