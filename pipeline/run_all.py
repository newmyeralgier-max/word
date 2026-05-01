"""
run_all.py — единый оркестратор всех 12 шагов пайплайна.

Применяет шаги по порядку 01..12 к входному .docx, сохраняет результат
в выходной .docx. Промежуточные файлы — в временной папке.

Каждый шаг — это отдельный скрипт с интерфейсом
    python pipeline/NN_<name>.py --input X.docx --output Y.docx

Зачем такая архитектура (вместо одного большого скрипта):
- Каждая проблема изолирована в своём файле.
- Любой шаг можно прогнать отдельно или пропустить (--skip / --only).
- Идемпотентность гарантируется на уровне отдельных скриптов.

Использование:
    # Полный прогон
    python pipeline/run_all.py \\
        --input  data/файл.docx \\
        --output data/файл_GOST.docx

    # Прогнать только шаги 09,10 (восстановление номеров страниц)
    python pipeline/run_all.py \\
        --input  data/файл.docx \\
        --output data/файл_fixed.docx \\
        --only 09,10

    # Прогнать всё, кроме шага 08 (объединение разделов)
    python pipeline/run_all.py \\
        --input  data/файл.docx \\
        --output data/файл_GOST.docx \\
        --skip 08

Бэкап оригинала автоматически создаётся в `data/backups/<имя>_pre_run.docx`.
"""

import argparse
import shutil
import subprocess
import sys
from pathlib import Path

PIPELINE_DIR = Path(__file__).resolve().parent
REPO_ROOT = PIPELINE_DIR.parent

# Список шагов в порядке применения. Префикс файла = ID шага.
STEPS = [
    ('01', 'pagesetup',           '01_pagesetup.py',           'A4, поля, titlePg, footerReference'),
    ('02', 'pagebreaks',          '02_pagebreaks.py',          'pageBreakBefore на H1 (главы с новой страницы)'),
    ('03', 'subscripts',          '03_subscripts.py',          'P_a → Pₐ (vertAlign=subscript)'),
    ('04', 'wrap_figures',        '04_wrap_figures.py',        'Рисунки в 2-ячеечные таблицы (рисунок + подпись)'),
    ('05', 'normalize_tables',    '05_normalize_tables.py',    'Таблицы данных: TNR 12pt, без жирных заголовков'),
    ('06', 'dedup_formulas',      '06_dedup_formulas.py',      'Удалить дубли (N.M)(N.M) в подписях формул'),
    ('07', 'renumber_formulas',   '07_renumber_formulas.py',   'Перенумерация формул сквозная по главе: (N.M)'),
    ('08', 'merge_sections',      '08_merge_sections.py',      'Объединить раздел 9 в 10 (в данном проекте)'),
    ('09', 'footer_pagenumber',   '09_footer_pagenumber.py',   'Канонический PAGE-field в footer1.xml'),
    ('10', 'remove_hidewhitespace', '10_remove_hidewhitespace.py', 'Снять <w:doNotDisplayPageBoundaries/>'),
    ('11', 'turbo_postpass',      '11_turbo_postpass.py',      'Постпроход: пустые pbb-параграфы, TOC stale, ПРОПАЛА, updateFields'),
    ('12', 'clear_highlights',    '12_clear_highlights.py',    'Снять все жёлтые подсветки (последним)'),
]


def _resolve_backup_dir() -> Path:
    """Find the right backups directory.

    The user's preferred layout is `WORD/data/backups/`. If that exists, use it.
    Fallback: legacy `data/backups/` for older clones.
    """
    candidates = [
        REPO_ROOT / 'WORD' / 'data' / 'backups',
        REPO_ROOT / 'data' / 'backups',
    ]
    for c in candidates:
        if c.exists():
            return c
    # Default to user-preferred path even if missing — will be created.
    return candidates[0]


def parse_filter(s: str) -> set[str]:
    if not s:
        return set()
    out = set()
    for part in s.split(','):
        part = part.strip()
        if not part:
            continue
        # allow "07" or "renumber_formulas"
        if part.isdigit():
            out.add(part.zfill(2))
        else:
            for sid, name, _, _ in STEPS:
                if name == part:
                    out.add(sid)
                    break
    return out


def main():
    ap = argparse.ArgumentParser(
        description='Apply all pipeline steps in order 01..12 to a .docx file.',
    )
    ap.add_argument('--input', help='Input .docx path')
    ap.add_argument('--output', help='Output .docx path')
    ap.add_argument('--only', default='',
                    help='Comma-separated list of step IDs/names to run '
                         '(e.g. "07,08" or "renumber_formulas")')
    ap.add_argument('--skip', default='',
                    help='Comma-separated list of step IDs/names to skip')
    ap.add_argument('--no-backup', action='store_true',
                    help='Do not create a backup in data/backups/')
    ap.add_argument('--list', action='store_true', help='List steps and exit')
    args = ap.parse_args()

    if args.list:
        print('Pipeline steps:')
        for sid, name, fname, desc in STEPS:
            print(f'  {sid}  {name:24s}  {desc}')
        return

    if not args.input or not args.output:
        ap.error('--input and --output are required (unless --list is given)')

    only = parse_filter(args.only)
    skip = parse_filter(args.skip)

    inp = Path(args.input).resolve()
    out = Path(args.output).resolve()
    if not inp.exists():
        print(f'ERROR: input not found: {inp}', file=sys.stderr)
        sys.exit(2)

    # Backup
    backup_dir = _resolve_backup_dir()
    if not args.no_backup:
        backup_dir.mkdir(parents=True, exist_ok=True)
        backup_path = backup_dir / f'{inp.stem}_pre_run.docx'
        shutil.copy2(inp, backup_path)
        try:
            rel = backup_path.relative_to(REPO_ROOT)
        except ValueError:
            rel = backup_path
        print(f'Backup → {rel}')

    # Tmp dir
    tmp_dir = backup_dir / '.run_tmp'
    if tmp_dir.exists():
        shutil.rmtree(tmp_dir)
    tmp_dir.mkdir(parents=True, exist_ok=True)

    cur = inp
    applied = 0
    for sid, name, fname, desc in STEPS:
        if only and sid not in only:
            continue
        if sid in skip:
            print(f'[skip {sid}] {name}: {desc}')
            continue
        next_path = tmp_dir / f'{sid}_after_{name}.docx'
        script = PIPELINE_DIR / fname
        print(f'\n=== Step {sid}: {name} ({desc}) ===')
        proc = subprocess.run(
            [sys.executable, str(script),
             '--input', str(cur),
             '--output', str(next_path)],
            capture_output=False,
        )
        if proc.returncode != 0:
            print(f'\nERROR: step {sid} ({name}) failed with code {proc.returncode}',
                  file=sys.stderr)
            sys.exit(proc.returncode)
        cur = next_path
        applied += 1

    # Copy final to output
    out.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(cur, out)
    print(f'\nDone. Applied {applied} step(s). Output → {out}')


if __name__ == '__main__':
    main()
