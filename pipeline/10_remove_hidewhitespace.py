"""
remove_hidewhitespace.py — remove <w:doNotDisplayPageBoundaries/> from
word/settings.xml.

This setting in OOXML corresponds to Word's "Hide White Space Between
Pages" / "Скрыть пустое пространство между страницами" option (toggled
via View tab or by double-clicking the page boundary in Print Layout).

When enabled, Word collapses inter-page white space — and as a side
effect HIDES headers and footers visually, even though they exist
correctly in the file. The user reported page numbers don't show in
the body of the document despite being correctly defined in
footer1.xml; this flag was set in settings.xml because Word saved it
that way after the user briefly toggled the view.

Removing the flag forces Word to display the full page boundaries
(including footers with the PAGE field).

The setting can also be re-enabled by the user at any time via
Word's UI; this script just makes sure the file ships in the
"normal" view state.
"""

import argparse
import re
import zipfile
from pathlib import Path


def process(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    if 'word/settings.xml' not in data:
        print('WARNING: word/settings.xml missing — nothing to do')
    else:
        s = data['word/settings.xml'].decode('utf-8')
        before = '<w:doNotDisplayPageBoundaries' in s
        s = re.sub(r'<w:doNotDisplayPageBoundaries\s*/>', '', s)
        s = re.sub(
            r'<w:doNotDisplayPageBoundaries[^/]*></w:doNotDisplayPageBoundaries>',
            '', s,
        )
        after = '<w:doNotDisplayPageBoundaries' in s
        if before and not after:
            print('Removed <w:doNotDisplayPageBoundaries/>')
        elif not before:
            print('No <w:doNotDisplayPageBoundaries/> found (already clean)')
        else:
            print('WARNING: regex did not match expected pattern')
        data['word/settings.xml'] = s.encode('utf-8')

    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, content in data.items():
            zout.writestr(n, content)


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--output', required=True)
    args = ap.parse_args()
    process(Path(args.input), Path(args.output))
