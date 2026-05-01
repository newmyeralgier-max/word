"""
clear_highlights.py — снимает все жёлтые подсветки (`<w:highlight w:val="yellow"/>`).

Исходный файл использовал жёлтые подсветки для маркировки проблемных мест.
После их фикса оставлять подсветку не нужно (финальный вид документа не должен
содержать жёлтых пятен).
"""

import argparse
import zipfile
from pathlib import Path
from lxml import etree

NS = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}


def qn(tag):
    return f'{{{NS["w"]}}}{tag}'


def process(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    doc = etree.fromstring(data['word/document.xml'])

    removed = 0
    for hl in doc.iter(qn('highlight')):
        val = hl.get(qn('val'))
        if val in ('yellow', 'cyan', 'green', 'magenta', 'red', 'lightGray'):
            parent = hl.getparent()
            parent.remove(hl)
            removed += 1

    data['word/document.xml'] = etree.tostring(
        doc, xml_declaration=True, encoding='UTF-8', standalone=True
    )
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, content in data.items():
            zout.writestr(n, content)

    print(f'Highlights removed: {removed}')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--output', required=True)
    args = ap.parse_args()
    process(Path(args.input), Path(args.output))
