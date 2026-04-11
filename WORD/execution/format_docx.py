"""
format_docx.py -- Неразрушающий ГОСТ-форматер для готовых Word-документов.

Принимает .docx ? анализирует абзацы ? применяет ГОСТ-стили.
Защищает: титульный лист, авторскую верстку, таблицы, списки, формулы.
Все токены -- из word_config.py, утилиты -- из word_utils.py.
"""

import sys, os, re

# ?? Путь к модулям рядом ??????????????????????????????????????????
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import word_config as cfg
import word_utils  as wu

from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


# ??????????????????????????????????????????????????????????????????
#  Умная вставка TOC
# ??????????????????????????????????????????????????????????????????

def _insert_toc_before_intro(doc):
    """Вставить автооглавление непосредственно перед ВВЕДЕНИЕМ."""
    if not doc.paragraphs:
        return

    target = doc.paragraphs[0]
    for p in doc.paragraphs:
        if p.text.strip().upper() == 'ВВЕДЕНИЕ':
            target = p
            break

    # Заголовок СОДЕРЖАНИЕ
    ph = target.insert_paragraph_before()
    ph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    ph.paragraph_format.space_before      = Pt(24)
    ph.paragraph_format.space_after       = Pt(12)
    ph.paragraph_format.first_line_indent = Cm(0)
    rh = ph.add_run('СОДЕРЖАНИЕ')
    rh.bold      = False
    rh.font.name = cfg.FONT_NAME
    rh.font.size = cfg.FONT_SIZE_TOC_TITLE

    # TOC field
    pt = target.insert_paragraph_before()
    pt.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pt.paragraph_format.first_line_indent = Cm(0)
    rt = pt.add_run()
    for tag, attr, text in [
        ('w:fldChar', 'begin', None),
        ('w:instrText', None, 'TOC \\o "1-3" \\h \\z \\u'),
        ('w:fldChar', 'separate', None),
    ]:
        el = OxmlElement(tag)
        if tag == 'w:instrText':
            el.set(qn('xml:space'), 'preserve')
            el.text = text
        else:
            el.set(qn('w:fldCharType'), attr)
        rt._r.append(el)
    # placeholder
    pr = OxmlElement('w:r')
    pt_el = OxmlElement('w:t')
    pt_el.text = '[Оглавление собрано автоматически]'
    pr.append(pt_el)
    rt._r.append(pr)
    end = OxmlElement('w:fldChar')
    end.set(qn('w:fldCharType'), 'end')
    rt._r.append(end)

    # Разрыв страницы
    pb = target.insert_paragraph_before()
    pb.add_run().add_break()


# ??????????????????????????????????????????????????????????????????
#  Главная функция
# ??????????????????????????????????????????????????????????????????

def format_document(input_path, no_com=False):
    print(f"+ Читаем: {input_path}")
    doc = Document(input_path)

    # ?? Стили ?????????????????????????????????????????????????????
    # Настраиваем Normal (только шрифт, без size -- чтобы не раздуть таблицы)
    doc.styles['Normal'].font.name = cfg.FONT_NAME

    wu.setup_gost_styles(doc)       # Heading 1-3 + TOC 1-3

    # ?? Анализ абзацев ????????????????????????????????????????????
    print("[*] Анализ абзацев...")
    formatting_on = False

    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue

        upper = text.upper()

        # Ожидание начала основного текста
        if not formatting_on:
            if re.match(r'^\s*(ВВЕДЕНИЕ|ОГЛАВЛЕНИЕ|СОДЕРЖАНИЕ|РЕФЕРАТ)\s*$', upper):
                formatting_on = True
        if not formatting_on:
            continue

        # Пропускаем TOC-записи
        if p.style.name.startswith('TOC'):
            continue

        # ?? Определяем заголовок ??????????????????????????????????
        is_heading = False
        level = 1
        too_long = len(text) > 120
        is_toc_title = upper in ('СОДЕРЖАНИЕ', 'ОГЛАВЛЕНИЕ')

        if is_toc_title:
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before      = Pt(24)
            p.paragraph_format.space_after       = Pt(12)
            p.paragraph_format.line_spacing      = cfg.LINE_SPACING
            p.paragraph_format.keep_with_next    = True
            for run in p.runs:
                run.bold      = False
                run.font.size = cfg.FONT_SIZE_TOC_TITLE
                run.font.name = cfg.FONT_NAME
                run.font.color.rgb = cfg.COLOR_BLACK
            continue

        if not too_long:
            if upper.strip() in cfg.STRICT_H1 or upper.strip().startswith('ПРИЛОЖЕНИЕ'):
                is_heading = True
                level = 1
            elif p.style.name.startswith('Heading') or p.style.name.startswith('Заголовок'):
                is_heading = True
                try:
                    level = int(p.style.name.split()[-1])
                except (ValueError, IndexError):
                    level = 1
            else:
                pPr = p._element.pPr
                if pPr is not None:
                    ol = pPr.find(qn('w:outlineLvl'))
                    if ol is not None:
                        val = int(ol.get(qn('w:val')))
                        if val < 9:
                            is_heading = True
                            level = val + 1

        # ?? Применяем стиль заголовка ?????????????????????????????
        if is_heading:
            is_strict = (upper.strip() in cfg.STRICT_H1
                         or upper.strip().startswith('ПРИЛОЖЕНИЕ'))

            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_strict else WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(0) if is_strict else cfg.FIRST_LINE_INDENT
            p.paragraph_format.line_spacing      = cfg.LINE_SPACING
            p.paragraph_format.keep_with_next    = True

            # Outline Level
            pPr = p._element.get_or_add_pPr()
            ol  = pPr.find(qn('w:outlineLvl'))
            if ol is None:
                ol = OxmlElement('w:outlineLvl')
                pPr.append(ol)
            ol.set(qn('w:val'), str(level - 1))

            p.paragraph_format.space_before = Pt(24) if level == 1 else Pt(18)
            p.paragraph_format.space_after  = Pt(12)

            bold_map = {1: cfg.BOLD_H1, 2: cfg.BOLD_H2, 3: cfg.BOLD_H3}
            size_map = {1: cfg.FONT_SIZE_H1, 2: cfg.FONT_SIZE_H2, 3: cfg.FONT_SIZE_H3}

            for run in p.runs:
                run.font.name = cfg.FONT_NAME
                run.font.size = size_map.get(level, cfg.FONT_SIZE_MAIN)
                run.bold      = bold_map.get(level, False)
            continue

        # ?? Обычный текст ?????????????????????????????????????????

        # Защита: центр/правое/индент- авторской верстки
        if p.alignment in (WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT):
            continue
        if p.paragraph_format.left_indent and p.paragraph_format.left_indent.cm > 0.5:
            continue
        # Защита: жёсткие табуляции (формулы)
        if '\t' in p.text:
            continue

        # Подписи к рисункам/таблицам
        if re.match(r'^(Рисунок|Таблица|Продолжение\s*таблицы)\s*\d+', text, re.I):
            p.paragraph_format.first_line_indent = Cm(0)
            if text.lower().startswith('рисунок'):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            else:
                p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after  = Pt(0)
        else:
            is_list = p.style.name.startswith('List') or 'numPr' in p._element.xml
            if not is_list:
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = cfg.FIRST_LINE_INDENT
            p.paragraph_format.line_spacing  = cfg.LINE_SPACING
            p.paragraph_format.space_before  = Pt(0)
            p.paragraph_format.space_after   = Pt(0)
            p.paragraph_format.widow_control = True

        for run in p.runs:
            if run.font.name and run.font.name.lower() in ('consolas', 'courier new'):
                continue
            run.font.name = cfg.FONT_NAME
            if run.font.size is None:
                run.font.size = cfg.FONT_SIZE_MAIN
            elif run.font.size != cfg.FONT_SIZE_MAIN and not (run.font.subscript or run.font.superscript):
                run.font.size = cfg.FONT_SIZE_MAIN

    # ?? Нумерация страниц ?????????????????????????????????????????
    wu.add_page_numbering(doc, smart_skip=True)

    # ?? TOC ???????????????????????????????????????????????????????
    doc_xml = doc._element.xml.upper()
    has_toc = bool(re.search(r'INSTRTEXT[^>]*>.*?TOC\s+', doc_xml)) \
              or ('СОДЕРЖАНИЕ' in doc_xml and 'SDT' in doc_xml)
    if not has_toc:
        for p in doc.paragraphs[:100]:
            if 'СОДЕРЖАНИЕ' in p.text.upper() or 'ОГЛАВЛЕНИЕ' in p.text.upper():
                has_toc = True
                break
    if not has_toc:
        print("[+] Вставляем автооглавление...")
        _insert_toc_before_intro(doc)

    # ?? Сохранение ????????????????????????????????????????????????
    out_dir = os.path.join(os.path.dirname(os.path.abspath(input_path)), 'data')
    if not os.path.exists(out_dir):
        out_dir = os.path.dirname(os.path.abspath(input_path))

    base = re.sub(r'_v\d+$', '', re.sub(r'_GOST$', '', os.path.splitext(os.path.basename(input_path))[0]))
    v = 1
    while True:
        output = os.path.join(out_dir, f"{base}_GOST_v{v}.docx")
        if not os.path.exists(output):
            break
        v += 1

    wu.save_document_safe(doc, output)

    # ?? COM-обновление ????????????????????????????????????????????
    if not no_com:
        wu.update_document_via_com(output)


import argparse

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Форматирование Word документов по ГОСТ")
    parser.add_argument('-i', '--input', required=True, help="Путь к файлу или папке для обработки")
    parser.add_argument('-o', '--output', default=None, help="Путь для сохранения результата")
    parser.add_argument('--no-com', action='store_true', help="Пропустить обновление через MS Word")
    
    args = parser.parse_args()
    
    target = args.input
    if not os.path.exists(target):
        print(f"[!] Путь не найден: {target}")
        sys.exit(1)
        
    if os.path.isdir(target):
        files = glob.glob(os.path.join(target, "*.docx"))
        from utils.docx_utils import find_working_files
        working_files = [f for f in files if "GOST" not in f and not os.path.basename(f).startswith("~")]
        for f in working_files:
            format_document(f, args.no_com)
    else:
        format_document(target, args.no_com)
