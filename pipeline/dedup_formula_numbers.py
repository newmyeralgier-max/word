"""
dedup_formula_numbers.py — убирает дублирующиеся номера формул вида «(N.M)(N.M)».

Что делает (per-paragraph пасс через lxml):
  • для каждого параграфа склеивает текст всех runs (вне oMath);
  • если в склейке встречается «(N.M)\s*(N.M)» (с одним и тем же N.M),
    удаляет последнее вхождение, найдя соответствующие run/runs и
    выкидывая их text;
  • если в run-ах встречается «(N,M)» — заменяет на «(N.M)»;
  • если параграф содержит ТОЛЬКО номер формулы (например «(2.1)»), и
    предыдущий параграф содержит oMath без номера — оставляем как есть
    (это нормальная ситуация для align_formulas).
"""

import argparse
import re
import zipfile
from pathlib import Path
from copy import deepcopy
from lxml import etree

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
}
W = NS['w']


def qn(tag, ns='w'):
    return f'{{{NS[ns]}}}{tag}'


NUM_RE = re.compile(r'\((\d+)[\.,](\d+)\)')


def collect_text_runs(p):
    """Return list of (run_element, w_t_element, text) for all w:t directly in w:r of p
    (excluding text inside m:oMath)."""
    out = []
    for r in p.findall('w:r', NS):
        # skip runs inside m:oMath
        for t in r.findall('w:t', NS):
            out.append((r, t, t.text or ''))
    return out


def text_of_paragraph_outside_math(p):
    """Joined text outside m:oMath."""
    out = []
    for t in p.iter(qn('t')):
        anc = t
        in_math = False
        while anc is not None:
            ql = etree.QName(anc).localname
            qns = etree.QName(anc).namespace
            if qns == NS['m'] and ql in ('oMath', 'oMathPara'):
                in_math = True
                break
            anc = anc.getparent()
        if not in_math and t.text:
            out.append(t.text)
    return ''.join(out)


def replace_comma_in_text(text):
    """Replace (N,M) with (N.M) in formula-like patterns."""
    return re.sub(r'\((\d+),(\d+)\)', r'(\1.\2)', text)


def dedup_runs(p):
    """Find duplicate (N.M) and remove the LAST run containing it; remove leading whitespace before it."""
    runs = collect_text_runs(p)
    if not runs:
        return 0, 0
    full = ''.join(rt for _, _, rt in runs)
    dedup = 0
    comma = 0

    # First, fix commas in numbers
    for r, t_el, txt in runs:
        new_txt = replace_comma_in_text(txt)
        if new_txt != txt:
            t_el.text = new_txt
            comma += 1
    # refresh
    runs = collect_text_runs(p)
    full = ''.join(rt for _, _, rt in runs)

    # find duplicate (N.M) sequences in full text
    matches = list(NUM_RE.finditer(full))
    if len(matches) < 2:
        return dedup, comma
    # group by (N,M) — if same (N,M) appears twice or more in immediate sequence
    seen = []
    to_remove_ranges = []  # global character ranges to delete
    for i, m in enumerate(matches):
        key = (m.group(1), m.group(2))
        prev = seen[-1] if seen else None
        if prev and prev[0] == key:
            # check that between prev end and current start only whitespace/separator chars
            gap = full[prev[1].end():m.start()]
            if re.fullmatch(r'\s*', gap):
                to_remove_ranges.append((m.start(), m.end()))
                continue
        seen.append((key, m))

    if not to_remove_ranges:
        return dedup, comma

    # apply removals: walk runs, accumulate offsets
    runs_text = [rt for _, _, rt in runs]
    # build cumulative offsets
    offsets = [0]
    for rt in runs_text:
        offsets.append(offsets[-1] + len(rt))
    # for each removal range, identify which runs and indices
    for (rs, re_) in sorted(to_remove_ranges, reverse=True):
        # determine starting run
        for i in range(len(runs)):
            run_start = offsets[i]
            run_end = offsets[i + 1]
            if run_start < re_ and run_end > rs:
                # this run intersects removal
                # local start/end
                local_s = max(0, rs - run_start)
                local_e = min(run_end - run_start, re_ - run_start)
                _, t_el, _ = runs[i]
                cur = t_el.text or ''
                t_el.text = cur[:local_s] + cur[local_e:]
        dedup += 1
        # rebuild runs_text/offsets
        runs = collect_text_runs(p)
        runs_text = [rt for _, _, rt in runs]
        offsets = [0]
        for rt in runs_text:
            offsets.append(offsets[-1] + len(rt))

    return dedup, comma


def process(input_path: Path, output_path: Path):
    with zipfile.ZipFile(input_path, 'r') as zin:
        data = {n: zin.read(n) for n in zin.namelist()}

    doc = etree.fromstring(data['word/document.xml'])

    total_dedup = 0
    total_comma = 0
    for p in doc.iter(qn('p')):
        d, c = dedup_runs(p)
        total_dedup += d
        total_comma += c

    data['word/document.xml'] = etree.tostring(
        doc, xml_declaration=True, encoding='UTF-8', standalone=True
    )
    with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zout:
        for n, content in data.items():
            zout.writestr(n, content)

    print(f'duplicate formula numbers removed: {total_dedup}')
    print(f'(N,M) → (N.M) fixes: {total_comma}')


if __name__ == '__main__':
    ap = argparse.ArgumentParser()
    ap.add_argument('--input', required=True)
    ap.add_argument('--output', required=True)
    args = ap.parse_args()
    process(Path(args.input), Path(args.output))
