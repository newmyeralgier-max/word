"""
wrap_figures.py — оборачивает рисунки в 2-ячеечные таблицы (как Рисунок 1.1).

Эталон (из существующего файла):
  <w:tbl>
    <w:tblPr>tblStyle aff0; tblW auto; jc center;
              tblBorders single 4 0 000000 (top/left/bottom/right/insideH/insideV);
              tblLook 04A0; firstRow=1 noVBand=1</w:tblPr>
    <w:tblGrid><w:gridCol w:w="9679"/></w:tblGrid>
    <w:tr><w:trPr><w:tblHeader/><w:jc center/></w:trPr>
      <w:tc><w:tcPr><w:tcW 9905 dxa/></w:tcPr><w:p>...drawing...</w:p></w:tc>
    </w:tr>
    <w:tr><w:trPr><w:jc center/></w:trPr>
      <w:tc><w:tcPr><w:tcW 9905 dxa/></w:tcPr><w:p>Рисунок X.Y — название</w:p></w:tc>
    </w:tr>
  </w:tbl>

Алгоритм:
  • Идём по верхнему уровню body.
  • Находим параграфы-подписи: «Рисунок N.M — ...» (parent = body, не tc).
  • Идём назад максимум 5 шагов в поисках параграфа с <w:drawing> (на верхнем уровне body).
  • Если нашли:
      – оборачиваем эти два параграфа в 2-ячеечную таблицу;
      – удаляем оригиналы.
  • Если не нашли (1.13 — без картинки):
      – создаём таблицу с пустой верхней ячейкой (placeholder) + caption.

Для рисунков 3.8 и 8.5: если рядом есть таблица с одной ячейкой, содержащая
drawing, и нет caption-row, — добавляем caption-row внутрь и удаляем
параграф-подпись снаружи.
"""

import argparse
import re
import zipfile
from pathlib import Path
from copy import deepcopy
from lxml import etree

NS = {
    'w':  'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r':  'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
W = NS['w']


def qn(tag):
    return f'{{{W}}}{tag}'


CAPTION_RE = re.compile(r'^\s*Рисунок\s+(\d+\.?\d*)\b', re.I)


def get_text(el):
    return ''.join(el.xpath('.//w:t/text()', namespaces=NS)).strip()


def has_drawing(el):
    return bool(el.xpath('.//w:drawing', namespaces=NS))


def is_caption_paragraph(p):
    if etree.QName(p).localname != 'p':
        return False
    return bool(CAPTION_RE.match(get_text(p)))


def make_figure_table(drawing_p, caption_p, *, empty_top=False):
    """Build a 2-row 1-column table that contains drawing in top, caption in bottom.

    drawing_p — параграф с <w:drawing> (или None при empty_top=True)
    caption_p — параграф с подписью «Рисунок X.Y — …»
    """
    tbl_xml = (
        '<w:tbl xmlns:w="{w}">'
        '<w:tblPr>'
        '<w:tblStyle w:val="aff0"/>'
        '<w:tblW w:w="0" w:type="auto"/>'
        '<w:jc w:val="center"/>'
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
        '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" '
        'w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9679"/></w:tblGrid>'
        '<w:tr><w:trPr><w:tblHeader/><w:jc w:val="center"/></w:trPr>'
        '<w:tc><w:tcPr><w:tcW w:w="9905" w:type="dxa"/></w:tcPr>'
        '__TOP_P__'
        '</w:tc></w:tr>'
        '<w:tr><w:trPr><w:jc w:val="center"/></w:trPr>'
        '<w:tc><w:tcPr><w:tcW w:w="9905" w:type="dxa"/></w:tcPr>'
        '__CAP_P__'
        '</w:tc></w:tr>'
        '</w:tbl>'
    ).format(w=W)
    tbl = etree.fromstring(tbl_xml)

    # locate placeholder cells
    rows = tbl.findall('w:tr', NS)
    top_tc = rows[0].find('w:tc', NS)
    cap_tc = rows[1].find('w:tc', NS)

    # remove placeholder children "__TOP_P__" / "__CAP_P__" (xml parser turned these
    # into text nodes in tc → but lxml kept text strings; we just clear and add)
    # For safety, remove all <w:p> currently in tc (none if placeholder didn't parse)
    for tc in (top_tc, cap_tc):
        for p in tc.findall('w:p', NS):
            tc.remove(p)
        # purge tail / text artifacts
        if tc.text:
            tc.text = None

    if empty_top or drawing_p is None:
        # empty placeholder paragraph in top cell with fixed height
        ph = etree.fromstring(
            '<w:p xmlns:w="{w}"><w:pPr>'
            '<w:ind w:firstLine="0"/>'
            '<w:jc w:val="center"/>'
            '<w:spacing w:before="0" w:after="0" w:line="4500" w:lineRule="exact"/>'
            '</w:pPr></w:p>'.format(w=W)
        )
        top_tc.append(ph)
    else:
        top_p_copy = deepcopy(drawing_p)
        # ensure pPr has firstLine 0
        pPr = top_p_copy.find('w:pPr', NS)
        if pPr is None:
            pPr = etree.SubElement(top_p_copy, qn('pPr'))
            top_p_copy.insert(0, pPr)
        ind = pPr.find('w:ind', NS)
        if ind is None:
            ind = etree.SubElement(pPr, qn('ind'))
        ind.set(qn('firstLine'), '0')
        # center the image in cell
        jc = pPr.find('w:jc', NS)
        if jc is None:
            jc = etree.SubElement(pPr, qn('jc'))
        jc.set(qn('val'), 'center')
        top_tc.append(top_p_copy)

    cap_p_copy = deepcopy(caption_p)
    pPr = cap_p_copy.find('w:pPr', NS)
    if pPr is None:
        pPr = etree.SubElement(cap_p_copy, qn('pPr'))
        cap_p_copy.insert(0, pPr)
    # caption pPr: firstLine=0, jc=center, single-line spacing
    ind = pPr.find('w:ind', NS)
    if ind is None:
        ind = etree.SubElement(pPr, qn('ind'))
    ind.set(qn('firstLine'), '0')
    jc = pPr.find('w:jc', NS)
    if jc is None:
        jc = etree.SubElement(pPr, qn('jc'))
    jc.set(qn('val'), 'center')
    sp = pPr.find('w:spacing', NS)
    if sp is None:
        sp = etree.SubElement(pPr, qn('spacing'))
    sp.set(qn('before'), '40')
    sp.set(qn('after'), '40')
    sp.set(qn('line'), '240')
    sp.set(qn('lineRule'), 'auto')
    cap_tc.append(cap_p_copy)

    return tbl


def find_nearby_drawing_paragraph(body_children, idx, max_back=5):
    """Look back up to max_back steps for a top-level paragraph with <w:drawing>."""
    for j in range(idx - 1, max(-1, idx - 1 - max_back), -1):
        el = body_children[j]
        if etree.QName(el).localname != 'p':
            continue
        if has_drawing(el):
            return j, el
        # if non-empty text, abort (caption→drawing must be adjacent-ish)
        txt = get_text(el)
        if txt:
            return None, None
    return None, None


def process(input_path: Path, output_path: Path, log_path: Path = None):
    with zipfile.ZipFile(input_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    doc = etree.fromstring(data['word/document.xml'])
    body = doc.find('w:body', NS)
    children = list(body)

    wrapped = []
    skipped_already_in_table = []
    placeholder_made = []

    # iterate from end to start so indices stay valid
    for i in range(len(children) - 1, -1, -1):
        el = children[i]
        if etree.QName(el).localname != 'p':
            continue
        # only top-level body captions
        if el.getparent() is not body:
            continue
        if not is_caption_paragraph(el):
            continue
        cap_text = get_text(el)
        m = CAPTION_RE.match(cap_text)
        fig_num = m.group(1) if m else '?'

        # search backward for drawing
        idx_d, draw_p = find_nearby_drawing_paragraph(children, i, max_back=5)

        if draw_p is None:
            # SPECIAL: figure 1.13 (no image) — create placeholder
            # Or the caption is orphaned because image is in adjacent table
            # Check if previous element is a tbl with single-cell containing drawing
            prev = children[i - 1] if i > 0 else None
            if prev is not None and etree.QName(prev).localname == 'tbl':
                # Append caption as new row to the existing tbl IF it doesn't already
                # have a caption row.
                rows = prev.findall('w:tr', NS)
                # collect existing caption text in last row
                last_row_text = ''.join(rows[-1].xpath('.//w:t/text()', namespaces=NS)).strip() if rows else ''
                if last_row_text and last_row_text.startswith('Рисунок'):
                    skipped_already_in_table.append(fig_num)
                    continue
                # append caption row
                # use grid width from existing tblGrid
                grid = prev.find('w:tblGrid', NS)
                grid_w = '9905'
                if grid is not None:
                    cols = grid.findall('w:gridCol', NS)
                    if cols:
                        try:
                            grid_w = cols[0].get(qn('w'))
                        except Exception:
                            pass
                # number of cells = number of <w:gridCol>
                ncols = max(1, len(grid.findall('w:gridCol', NS))) if grid is not None else 1
                # build caption row with merged cell across all columns (gridSpan)
                row_xml = (
                    '<w:tr xmlns:w="{w}"><w:trPr><w:jc w:val="center"/></w:trPr>'
                    '<w:tc><w:tcPr><w:tcW w:w="{tw}" w:type="dxa"/>'
                    + (f'<w:gridSpan w:val="{ncols}"/>' if ncols > 1 else '')
                    + '</w:tcPr>'
                    '<w:p><w:pPr><w:ind w:firstLine="0"/><w:jc w:val="center"/>'
                    '<w:spacing w:before="40" w:after="40" w:line="240" w:lineRule="auto"/>'
                    '</w:pPr></w:p>'
                    '</w:tc></w:tr>'
                ).format(w=W, tw=grid_w)
                row = etree.fromstring(row_xml)
                # populate the inner <w:p> by copying from caption
                tc = row.find('w:tc', NS)
                # remove auto-created empty p in tc
                for p in tc.findall('w:p', NS):
                    tc.remove(p)
                cap_copy = deepcopy(el)
                pPr = cap_copy.find('w:pPr', NS)
                if pPr is None:
                    pPr = etree.SubElement(cap_copy, qn('pPr'))
                    cap_copy.insert(0, pPr)
                # set caption pPr
                ind = pPr.find('w:ind', NS)
                if ind is None:
                    ind = etree.SubElement(pPr, qn('ind'))
                ind.set(qn('firstLine'), '0')
                jc = pPr.find('w:jc', NS)
                if jc is None:
                    jc = etree.SubElement(pPr, qn('jc'))
                jc.set(qn('val'), 'center')
                sp = pPr.find('w:spacing', NS)
                if sp is None:
                    sp = etree.SubElement(pPr, qn('spacing'))
                sp.set(qn('before'), '40')
                sp.set(qn('after'), '40')
                sp.set(qn('line'), '240')
                sp.set(qn('lineRule'), 'auto')
                tc.append(cap_copy)
                prev.append(row)
                # remove the original caption paragraph
                body.remove(el)
                wrapped.append(f'{fig_num} (приписана к существующей таблице)')
                continue

            # No nearby drawing & no preceding table → make placeholder (e.g. 1.13)
            tbl = make_figure_table(None, el, empty_top=True)
            body.replace(el, tbl)
            placeholder_made.append(fig_num)
            continue

        # found drawing → wrap into new table
        tbl = make_figure_table(draw_p, el)
        # position: replace drawing_p with tbl, then remove caption_p
        body.replace(draw_p, tbl)
        body.remove(el)
        wrapped.append(fig_num)

    # serialize
    data['word/document.xml'] = etree.tostring(
        doc, xml_declaration=True, encoding='UTF-8', standalone=True
    )
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, content in data.items():
            zout.writestr(n, content)

    print(f'Wrapped figures: {len(wrapped)} → {wrapped}')
    print(f'Placeholder made (no image): {len(placeholder_made)} → {placeholder_made}')
    print(f'Skipped (already in table): {len(skipped_already_in_table)} → {skipped_already_in_table}')

    if log_path:
        with open(log_path, 'w', encoding='utf-8') as f:
            f.write('wrapped: ' + ', '.join(wrapped) + '\n')
            f.write('placeholder: ' + ', '.join(placeholder_made) + '\n')
            f.write('skipped: ' + ', '.join(skipped_already_in_table) + '\n')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--output', required=True)
    ap.add_argument('--log', default=None)
    args = ap.parse_args()
    process(Path(args.input), Path(args.output), Path(args.log) if args.log else None)
