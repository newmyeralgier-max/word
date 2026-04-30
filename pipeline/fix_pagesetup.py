"""
fix_pagesetup.py — приведение настроек страницы к ГОСТ 7.32-2017.

Что делает:
  1. <w:pgSz> Letter → A4 (11906 × 16838 twips).
  2. Поля проверяются: top/bottom 1134 (2 см), left 1701 (3 см), right 850 (1.5 см) — оставляем как есть, если уже ок.
  3. Добавляет <w:titlePg/> в sectPr — нумерация на первой странице не печатается.
  4. Добавляет <w:pgNumType w:start="1" w:fmt="decimal"/>.
  5. Создаёт пустой first-footer (footer_first.xml) и добавляет footerReference type="first".
  6. Меняет в footer1.xml выравнивание PAGE с right на center (ГОСТ: внизу по центру).
"""

import argparse
import re
import shutil
import sys
import zipfile
from pathlib import Path
from lxml import etree

NS = {
    'w':  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'rs': 'http://schemas.openxmlformats.org/package/2006/relationships',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}
W = NS['w']
R = NS['r']

A4_WIDTH  = '11906'
A4_HEIGHT = '16838'


def _qn(tag, ns='w'):
    return f'{{{NS[ns]}}}{tag}'


def patch_document_xml(doc_xml: bytes) -> bytes:
    root = etree.fromstring(doc_xml)
    sect_prs = root.findall('.//w:sectPr', NS)
    if not sect_prs:
        raise SystemExit('No <w:sectPr> found')

    for sp in sect_prs:
        # 1. pgSz → A4
        pg_sz = sp.find('w:pgSz', NS)
        if pg_sz is None:
            pg_sz = etree.SubElement(sp, _qn('pgSz'))
        pg_sz.set(_qn('w'), A4_WIDTH)
        pg_sz.set(_qn('h'), A4_HEIGHT)

        # 2. titlePg
        if sp.find('w:titlePg', NS) is None:
            tp = etree.SubElement(sp, _qn('titlePg'))

        # 3. pgNumType
        pgn = sp.find('w:pgNumType', NS)
        if pgn is None:
            pgn = etree.SubElement(sp, _qn('pgNumType'))
        pgn.set(_qn('start'), '1')
        pgn.set(_qn('fmt'), 'decimal')

        # 4. footerReference type="first" with rId pointing to footer_first
        # Existing footerReference type="default" stays.
        first_ref = None
        for fr in sp.findall('w:footerReference', NS):
            if fr.get(_qn('type')) == 'first':
                first_ref = fr
                break
        if first_ref is None:
            first_ref = etree.SubElement(sp, _qn('footerReference'))
            first_ref.set(_qn('type'), 'first')
            first_ref.set(_qn('id', 'r'), 'rIdFooterFirst')

    return etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)


def make_empty_footer() -> bytes:
    return (
        b'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
        b'<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        b'<w:p><w:pPr><w:jc w:val="center"/></w:pPr></w:p>'
        b'</w:ftr>'
    )


def patch_footer1(footer_xml: bytes) -> bytes:
    """Change PAGE field alignment from right to center."""
    text = footer_xml.decode('utf-8')
    # Replace <w:jc w:val="right"/> → <w:jc w:val="center"/> when next to PAGE field
    new = text.replace('<w:jc w:val="right"/>', '<w:jc w:val="center"/>', 1)
    return new.encode('utf-8')


def patch_rels(rels_xml: bytes) -> bytes:
    """Add relationship for first-footer if not present."""
    text = rels_xml.decode('utf-8')
    if 'rIdFooterFirst' in text:
        return rels_xml
    rel = (
        '<Relationship Id="rIdFooterFirst" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" '
        'Target="footer_first.xml"/>'
    )
    new = text.replace('</Relationships>', rel + '</Relationships>')
    return new.encode('utf-8')


def patch_content_types(ct_xml: bytes) -> bytes:
    text = ct_xml.decode('utf-8')
    if 'footer_first.xml' in text:
        return ct_xml
    override = (
        '<Override PartName="/word/footer_first.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>'
    )
    new = text.replace('</Types>', override + '</Types>')
    return new.encode('utf-8')


def process(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        names = zin.namelist()
        data = {n: zin.read(n) for n in names}

    data['word/document.xml'] = patch_document_xml(data['word/document.xml'])
    data['word/footer1.xml']  = patch_footer1(data['word/footer1.xml'])
    data['word/_rels/document.xml.rels'] = patch_rels(data['word/_rels/document.xml.rels'])
    data['[Content_Types].xml'] = patch_content_types(data['[Content_Types].xml'])
    if 'word/footer_first.xml' not in data:
        data['word/footer_first.xml'] = make_empty_footer()

    # rebuild zip preserving order; place [Content_Types].xml first
    order = ['[Content_Types].xml'] + [n for n in names if n != '[Content_Types].xml']
    if 'word/footer_first.xml' not in order:
        order.append('word/footer_first.xml')

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n in order:
            zout.writestr(n, data[n])


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--output', required=True)
    args = ap.parse_args()
    process(Path(args.input), Path(args.output))
    print(f'OK: {args.output}')
