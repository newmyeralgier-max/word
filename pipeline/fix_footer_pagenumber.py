"""
fix_footer_pagenumber.py — rewrite footer1.xml with a clean, Word-canonical
PAGE field, in case the existing one isn't being rendered by Word for some
reason.

Issue: user reports that page numbers are not appearing in the body of the
document (only in TOC after F9). Footer1.xml has a PAGE field with cached
value "91", but the structure may be subtly malformed:
  - <w:instrText> without xml:space="preserve"
  - No \\* MERGEFORMAT switch
  - No explicit font/size on the runs

This script writes a known-good structure that mirrors what Word emits when
you Insert > Page Number > Bottom > Plain Number 2.
"""

import argparse
import zipfile
from pathlib import Path


CANONICAL_FOOTER1 = '''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:ftr xmlns:wpc="http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas" xmlns:cx="http://schemas.microsoft.com/office/drawing/2014/chartex" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" xmlns:v="urn:schemas-microsoft-com:vml" xmlns:wp14="http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:w10="urn:schemas-microsoft-com:office:word" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml" xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml" xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" xmlns:wpi="http://schemas.microsoft.com/office/word/2010/wordprocessingInk" xmlns:wne="http://schemas.microsoft.com/office/word/2006/wordml" xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" mc:Ignorable="w14 w15 wp14"><w:p w14:paraId="11111111" w14:textId="77777777" w:rsidR="004D394F" w:rsidRDefault="004D394F"><w:pPr><w:pStyle w:val="a7"/><w:jc w:val="center"/><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr></w:pPr><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:instrText xml:space="preserve"> PAGE   \\* MERGEFORMAT </w:instrText></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:noProof/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:t>2</w:t></w:r><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:cs="Times New Roman"/><w:sz w:val="28"/><w:szCs w:val="28"/></w:rPr><w:fldChar w:fldCharType="end"/></w:r></w:p></w:ftr>'''


def process(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    # Rewrite footer1.xml with canonical structure
    if 'word/footer1.xml' in data:
        old_size = len(data['word/footer1.xml'])
        data['word/footer1.xml'] = CANONICAL_FOOTER1.encode('utf-8')
        new_size = len(data['word/footer1.xml'])
        print(f'word/footer1.xml: {old_size} → {new_size} bytes')
    else:
        print('WARNING: word/footer1.xml not found')

    # Make sure settings.xml has updateFields=true (Word may have stripped it)
    if 'word/settings.xml' in data:
        s = data['word/settings.xml'].decode('utf-8')
        if '<w:updateFields' not in s:
            # inject before </w:settings>
            inject = '<w:updateFields w:val="true"/>'
            s = s.replace('</w:settings>', f'{inject}</w:settings>')
            data['word/settings.xml'] = s.encode('utf-8')
            print('word/settings.xml: injected <w:updateFields w:val="true"/>')
        else:
            print('word/settings.xml: already has updateFields')

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, content in data.items():
            zout.writestr(n, content)


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--output', required=True)
    args = ap.parse_args()
    process(Path(args.input), Path(args.output))
