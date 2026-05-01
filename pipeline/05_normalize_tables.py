"""
normalize_tables.py — приводит таблицы к единому ГОСТ-стилю.

Что делает:
  1. Идентифицирует «таблицы-данные» (исключает «таблицы-обёртки рисунков»):
     обёртки имеют первую ячейку с <w:drawing> или это новая таблица из
     wrap_figures.py с placeholder-ом.
  2. Для каждой таблицы-данных:
     • убирает <w:b/> и <w:bCs/> в текстовых run-ах заголовочной строки;
     • выставляет шрифт Times New Roman 24 (12 pt) во всех run-ах ячеек;
     • межстрочный интервал — single (240, auto);
     • выравнивание — по центру в первой строке, по умолчанию (по обстановке)
       в остальных;
     • <w:tblW> = 0 / type=auto, <w:tblLayout> = autofit.

ГОСТ 2.105-2019: внутри таблицы шрифт TNR ≤ 12 пт; одинарный интервал;
заголовки строк — без жирного начертания (по локальной методичке).
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


def is_figure_wrapper(tbl):
    """Figure-wrapper tables: first cell contains <w:drawing>; OR
    placeholder created by wrap_figures (one row with empty p, second with caption text 'Рисунок')."""
    rows = tbl.findall('w:tr', NS)
    if not rows:
        return False
    first_row = rows[0]
    if first_row.xpath('.//w:drawing', namespaces=NS):
        return True
    # placeholder check: look for <w:p> with line=4500 (placeholder)
    if first_row.xpath('.//w:spacing[@w:line="4500"]', namespaces=NS):
        return True
    # caption row check: any row with text starting with 'Рисунок'
    for r in rows[1:]:
        txt = ''.join(r.xpath('.//w:t/text()', namespaces=NS)).strip()
        if txt.startswith('Рисунок '):
            return True
    return False


def normalize_run_font(rPr):
    """Set TNR 12 pt; remove bold."""
    if rPr is None:
        return
    # Remove <w:b/> and <w:bCs/>
    for tag in ('b', 'bCs'):
        for el in rPr.findall(f'w:{tag}', NS):
            rPr.remove(el)
    # rFonts → TNR
    rfonts = rPr.find('w:rFonts', NS)
    if rfonts is None:
        rfonts = etree.SubElement(rPr, qn('rFonts'))
    rfonts.set(qn('ascii'), 'Times New Roman')
    rfonts.set(qn('hAnsi'), 'Times New Roman')
    rfonts.set(qn('cs'), 'Times New Roman')
    rfonts.set(qn('eastAsia'), 'Times New Roman')
    # size
    sz = rPr.find('w:sz', NS)
    if sz is None:
        sz = etree.SubElement(rPr, qn('sz'))
    sz.set(qn('val'), '24')  # 12 pt half-points
    szcs = rPr.find('w:szCs', NS)
    if szcs is None:
        szcs = etree.SubElement(rPr, qn('szCs'))
    szcs.set(qn('val'), '24')


def normalize_paragraph_in_cell(p, *, header=False):
    pPr = p.find('w:pPr', NS)
    if pPr is None:
        pPr = etree.SubElement(p, qn('pPr'))
        p.insert(0, pPr)
    # spacing single
    sp = pPr.find('w:spacing', NS)
    if sp is None:
        sp = etree.SubElement(pPr, qn('spacing'))
    sp.set(qn('before'), '20')
    sp.set(qn('after'),  '20')
    sp.set(qn('line'),   '240')
    sp.set(qn('lineRule'), 'auto')
    # firstLine 0
    ind = pPr.find('w:ind', NS)
    if ind is None:
        ind = etree.SubElement(pPr, qn('ind'))
    ind.set(qn('firstLine'), '0')
    # for headers — center
    if header:
        jc = pPr.find('w:jc', NS)
        if jc is None:
            jc = etree.SubElement(pPr, qn('jc'))
        jc.set(qn('val'), 'center')
    # process runs
    for r in p.findall('w:r', NS):
        rPr = r.find('w:rPr', NS)
        if rPr is None:
            rPr = etree.SubElement(r, qn('rPr'))
            r.insert(0, rPr)
        normalize_run_font(rPr)
    # Also normalize runs inside <w:hyperlink>
    for r in p.xpath('.//w:hyperlink/w:r', namespaces=NS):
        rPr = r.find('w:rPr', NS)
        if rPr is None:
            rPr = etree.SubElement(r, qn('rPr'))
            r.insert(0, rPr)
        normalize_run_font(rPr)


def normalize_table(tbl):
    rows = tbl.findall('w:tr', NS)
    if not rows:
        return
    # tblPr: tblW auto, tblLayout autofit
    tblPr = tbl.find('w:tblPr', NS)
    if tblPr is None:
        tblPr = etree.Element(qn('tblPr'))
        tbl.insert(0, tblPr)
    tw = tblPr.find('w:tblW', NS)
    if tw is None:
        tw = etree.SubElement(tblPr, qn('tblW'))
    tw.set(qn('w'), '0')
    tw.set(qn('type'), 'auto')
    layout = tblPr.find('w:tblLayout', NS)
    if layout is None:
        layout = etree.SubElement(tblPr, qn('tblLayout'))
    layout.set(qn('type'), 'autofit')

    # process rows
    for ridx, tr in enumerate(rows):
        is_header = ridx == 0
        for tc in tr.findall('w:tc', NS):
            for p in tc.findall('w:p', NS):
                normalize_paragraph_in_cell(p, header=is_header)


def process(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    doc = etree.fromstring(data['word/document.xml'])
    body = doc.find('w:body', NS)

    total = 0
    normalized = 0
    figure_wrappers = 0
    for tbl in body.iter(qn('tbl')):
        total += 1
        if is_figure_wrapper(tbl):
            figure_wrappers += 1
            continue
        normalize_table(tbl)
        normalized += 1

    data['word/document.xml'] = etree.tostring(
        doc, xml_declaration=True, encoding='UTF-8', standalone=True
    )
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, content in data.items():
            zout.writestr(n, content)

    print(f'Total tables: {total}, normalized: {normalized}, figure wrappers skipped: {figure_wrappers}')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--output', required=True)
    args = ap.parse_args()
    process(Path(args.input), Path(args.output))
