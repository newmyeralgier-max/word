import os

files = [
    r'd:\1. Project\Word\WORD\execution\build_docx.py',
    r'd:\1. Project\Word\WORD\execution\format_docx.py',
    r'd:\1. Project\Word\WORD\execution\word_utils.py',
]

# Russian chars to keep
RU = set('АаБбВвГгДдЕеЁёЖжЗзИиЙйКкЛлМмНнОоПпРрСсТтУуФфХхЦцЧчШшЩщЪъЫыЬьЭэЮюЯя')

REPLACE = {
    '\u2011': '-',
    '\u2014': '--',
    '\u2013': '-',
    '\u2019': "'",
    '\u201c': '"',
    '\u201d': '"',
    '\u2026': '...',
    '\u00ab': '"',
    '\u00bb': '"',
    '\u2212': '-',
    '\u00d7': 'x',
}

for fpath in files:
    with open(fpath, 'r', encoding='utf-8') as f:
        text = f.read()

    out = []
    for ch in text:
        if ord(ch) <= 127 or ch in RU:
            out.append(ch)
        elif ch in REPLACE:
            out.append(REPLACE[ch])
        else:
            out.append('?')

    with open(fpath, 'w', encoding='utf-8') as f:
        f.write(''.join(out))
    print(f'Cleaned: {os.path.basename(fpath)}')
