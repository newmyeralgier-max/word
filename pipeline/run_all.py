"""Полный прогон ГОСТ-пайплайна v4.1.

Шаги:
1. Извлекаем титульник (01_extract_title.py) → title.docx + body_00.docx
2. Типография + авто-сплит слипшихся слов (normalize.apply_typography)
3. Десятичные разделители (normalize.fix_decimals)
4. Замена «•» → «—» + снятие bold с body (normalize.replace_bullets_and_unbold)
5. Срез точек в конце заголовков (normalize.strip_trailing_periods_in_headings)
6. Подписи рисунков/таблиц (captions.process_document)
7. ГОСТ-форматирование тела (WORD/execution/format_docx.process_document)
8. Склейка: title.docx + body → final.docx (stitch_title.stitch)
9. Валидация (WORD/execution/validate_formatter.py) — опционально

Перенумерация глав в этой версии ВЫКЛЮЧЕНА (SECTION_MAP пуст) — в прошлый
раз она сломала иерархию разделов (жалобы пользователя #11, #13, #14).
Оригинальные номера из исходника НЕ трогаются.

Использование:
    cd pipeline && python run_all.py <input.docx> <output.docx>
"""
from __future__ import annotations

import importlib.util
import os
import shutil
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).parent.resolve()
REPO = HERE.parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(REPO / "WORD" / "execution"))

from normalize import (
    apply_typography,
    fix_decimals,
    renumber_sections,
    replace_bullets_and_unbold,
    strip_trailing_periods_in_headings,
)
from restructure import restructure
from captions import process_document as apply_captions
from stitch_title import stitch


def _load(path: Path, name: str):
    spec = importlib.util.spec_from_file_location(name, str(path))
    mod = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(mod)
    return mod


extract_mod = _load(HERE / "01_extract_title.py", "_extract_title")
format_mod = _load(REPO / "WORD" / "execution" / "format_docx.py", "_format_docx")


def run(input_path: str, output_path: str, keep_intermediates: bool = True) -> dict:
    input_path = str(Path(input_path).resolve())
    output_path = str(Path(output_path).resolve())
    work_dir = Path(output_path).parent / "_pipeline_work"
    work_dir.mkdir(parents=True, exist_ok=True)

    title_docx = str(work_dir / "title.docx")
    body_00 = str(work_dir / "body_00_raw.docx")
    body_01 = str(work_dir / "body_01_typo.docx")
    body_02 = str(work_dir / "body_02_decimals.docx")
    body_03 = str(work_dir / "body_03_renumber.docx")
    body_04 = str(work_dir / "body_04_bullets_unbold.docx")
    body_05 = str(work_dir / "body_05_strip_periods.docx")
    body_06 = str(work_dir / "body_06_captions.docx")
    captions_json = str(work_dir / "captions.json")
    body_07 = str(work_dir / "body_07_gost.docx")

    report: dict = {}

    print(f"[1/8] extract title : {input_path}")
    extract_mod.extract(input_path, title_docx, body_00)

    print(f"[2/8] typography    : {body_00}")
    report["typography"] = apply_typography(body_00, body_01)
    print("       ", report["typography"])

    print(f"[3/8] decimals      : {body_01}")
    report["decimals"] = fix_decimals(body_01, body_02)
    print("       ", report["decimals"])

    print(f"[4/8] restructure   : {body_02}")
    report["restructure"] = restructure(body_02, body_03)
    print("       ",
          {k: v for k, v in report["restructure"].items() if k != "log_sample"})

    print(f"[4b]   bullets/unbold : {body_03}")
    report["bullets_unbold"] = replace_bullets_and_unbold(body_03, body_04)
    print("       ", report["bullets_unbold"])

    print(f"[5/8] strip periods : {body_04}")
    report["strip_periods"] = strip_trailing_periods_in_headings(body_04, body_05)
    print("       ", report["strip_periods"])

    print(f"[6/8] captions      : {body_05}")
    report["captions"] = apply_captions(body_05, body_06, captions_json)
    print("       ",
          {k: v for k, v in report["captions"].items() if not isinstance(v, list)})

    print(f"[7/8] GOST format   : {body_06}")
    # format_docx.process_document pattern: принимает input path и пишет рядом
    # с суффиксом _GOST.docx. Нам нужен точно body_07 на выходе.
    # Используем copy + process + rename.
    tmp_src = str(work_dir / "body_06_captions_for_gost.docx")
    shutil.copy(body_06, tmp_src)
    format_mod.process_document(tmp_src, fast=True)
    produced = tmp_src.replace(".docx", "_GOST.docx")
    shutil.move(produced, body_07)
    try:
        os.remove(tmp_src)
    except OSError:
        pass

    print(f"[8/8] stitch title  : {title_docx} + {body_07} → {output_path}")
    stitch(body_07, title_docx, output_path)

    print(f"[DONE] {output_path}")
    return report


if __name__ == "__main__":
    if len(sys.argv) != 3:
        print("Usage: python run_all.py <input.docx> <output.docx>")
        sys.exit(1)
    run(sys.argv[1], sys.argv[2])
