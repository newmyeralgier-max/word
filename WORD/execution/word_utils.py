"""
word_utils.py -- Общие утилиты для ГОСТ-форматирования Word-документов.

Используется и build_docx.py, и format_docx.py.
Содержит: стилизацию, нумерацию страниц, MML2OMML, безопасное сохранение, COM-обновление.
"""

import os
import sys
import re
import glob
import shutil

from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree

# Загружаем конфиг (файл лежит рядом)
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import word_config as cfg


# ??????????????????????????????????????????????????????????????????
#  ГОСТ-стили
# ??????????????????????????????????????????????????????????????????

def setup_gost_styles(doc):
    """Настроить встроенные стили Word под ГОСТ 7.32 (Normal, Heading 1-3, TOC 1-3)."""

    # ?? Normal ??????????????????????????????????????????????????
    sn = doc.styles['Normal']
    sn.font.name      = cfg.FONT_NAME
    sn.font.size      = cfg.FONT_SIZE_MAIN
    sn.font.color.rgb  = cfg.COLOR_BLACK
    pf = sn.paragraph_format
    pf.space_before       = Pt(0)
    pf.space_after        = Pt(0)
    pf.line_spacing       = cfg.LINE_SPACING
    pf.first_line_indent  = cfg.FIRST_LINE_INDENT
    pf.alignment          = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.widow_control      = True

    # ?? Heading 1-3 (английское + русское имя) ??????????????????
    bold_map = {1: cfg.BOLD_H1, 2: cfg.BOLD_H2, 3: cfg.BOLD_H3}
    size_map = {1: cfg.FONT_SIZE_H1, 2: cfg.FONT_SIZE_H2, 3: cfg.FONT_SIZE_H3}

    for lvl in range(1, 4):
        for name in (f'Heading {lvl}', f'Заголовок {lvl}'):
            try:
                hs = doc.styles[name]
            except KeyError:
                continue
            hs.font.name       = cfg.FONT_NAME
            hs.font.color.rgb  = cfg.COLOR_BLACK
            hs.font.bold       = bold_map[lvl]
            hs.font.size       = size_map[lvl]
            hs.paragraph_format.space_before      = Pt(24) if lvl == 1 else Pt(18)
            hs.paragraph_format.space_after        = Pt(12)
            hs.paragraph_format.line_spacing       = cfg.LINE_SPACING
            hs.paragraph_format.keep_with_next     = True
            hs.paragraph_format.alignment          = WD_ALIGN_PARAGRAPH.CENTER
            hs.paragraph_format.first_line_indent  = Cm(0)

    # ?? TOC 1-3 ????????????????????????????????????????????????
    for i in range(1, 4):
        try:
            ts = doc.styles[f'TOC {i}']
        except KeyError:
            continue
        ts.font.name  = cfg.FONT_NAME
        ts.font.size  = cfg.FONT_SIZE_MAIN
        ts.font.bold  = False
        ts.paragraph_format.space_before  = Pt(0)
        ts.paragraph_format.space_after   = Pt(0)
        ts.paragraph_format.line_spacing  = cfg.LINE_SPACING


# ??????????????????????????????????????????????????????????????????
#  MML2OMML -- расположение XSL-файла и кэш
# ??????????????????????????????????????????????????????????????????

_xsl_transform = None          # кэш XSLT-объекта

def _local_xsl_path():
    return os.path.join(os.path.dirname(__file__), 'MML2OMML.XSL')


def get_omml_xslt():
    """Вернуть lxml.etree.XSLT или False, если XSLT не найден."""
    global _xsl_transform
    if _xsl_transform is not None:
        return _xsl_transform

    local = _local_xsl_path()

    # 1) Нет локальной копии -- ищем в Office и копируем
    if not os.path.exists(local):
        patterns = [
            r"C:\Program Files\Microsoft Office\root\Office*\MML2OMML.XSL",
            r"C:\Program Files (x86)\Microsoft Office\root\Office*\MML2OMML.XSL",
            r"C:\Program Files\Microsoft Office\Office*\MML2OMML.XSL",
            r"C:\Program Files (x86)\Microsoft Office\Office*\MML2OMML.XSL",
        ]
        source = None
        for pat in patterns:
            matches = glob.glob(pat)
            if matches:
                source = matches[0]
                break
        if source:
            print(f"  [+] Копируем MML2OMML.XSL ? {local}")
            shutil.copy(source, local)
        else:
            print("  [!] MML2OMML.XSL не найден -- формулы будут текстовыми")
            _xsl_transform = False
            return False

    # 2) Парсим
    try:
        _xsl_transform = etree.XSLT(etree.parse(local))
    except Exception as e:
        print(f"  [!] Ошибка загрузки XSLT: {e}")
        _xsl_transform = False

    return _xsl_transform


def latex_to_omml(latex_str):
    """Конвертировать LaTeX ? OMML XML-элемент (или None)."""
    import latex2mathml.converter

    xslt = get_omml_xslt()
    if not xslt:
        return None
    try:
        mathml_str  = latex2mathml.converter.convert(latex_str)
        mathml_tree = etree.fromstring(mathml_str.encode('utf-8'))
        return xslt(mathml_tree).getroot()
    except Exception as e:
        print(f"  [!] Формула '{latex_str[:40]}...': {e}")
        return None


# ??????????????????????????????????????????????????????????????????
#  Нумерация страниц
# ??????????????????????????????????????????????????????????????????

def add_page_field(run):
    """Вставить поле PAGE в данный run."""
    for tag, attr in [
        ('w:fldChar', 'begin'),
        ('w:instrText', None),
        ('w:fldChar', 'separate'),
        ('w:fldChar', 'end'),
    ]:
        el = OxmlElement(tag)
        if tag == 'w:instrText':
            el.set(qn('xml:space'), 'preserve')
            el.text = 'PAGE'
        else:
            el.set(qn('w:fldCharType'), attr)
        run._r.append(el)


def add_page_numbering(doc, *, smart_skip=True, align=WD_ALIGN_PARAGRAPH.RIGHT):
    """
    Добавить номера страниц в футер.
    smart_skip=True: пропускает секции, где нумерация уже есть.
    """
    for section in doc.sections:
        footer = section.footer

        if smart_skip:
            f_xml = footer._element.xml.upper() if footer else ''
            h_xml = section.header._element.xml.upper() if section.header else ''
            if (re.search(r'INSTRTEXT.*?(PAGE|NUMPAGES)', f_xml)
                    or re.search(r'INSTRTEXT.*?(PAGE|NUMPAGES)', h_xml)):
                continue
            # цифры в footer
            if footer and any(re.search(r'\d', p.text) for p in footer.paragraphs):
                continue

        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = align
        p.clear()
        run = p.add_run()
        run.font.name = cfg.FONT_NAME
        run.font.size = cfg.FONT_SIZE_MAIN
        add_page_field(run)


# ??????????????????????????????????????????????????????????????????
#  Безопасное сохранение
# ??????????????????????????????????????????????????????????????????

def save_document_safe(doc, output_path):
    """Сохранить с перехватом PermissionError (файл открыт в Word)."""
    while True:
        try:
            doc.save(output_path)
            print(f"\n[OK] Документ сохранён: {os.path.abspath(output_path)}")
            return True
        except PermissionError:
            print(f"\n[!] Файл '{os.path.basename(output_path)}' открыт в другой программе!")
            input("   Закройте его и нажмите Enter для повторной попытки...")


# ??????????????????????????????????????????????????????????????????
#  Word COM: обновление TOC + страницы
# ??????????????????????????????????????????????????????????????????

def update_document_via_com(output_path, *, visible=True):
    """Открывает документ через COM, обновляет поля, TOC, убирает жирность из TOC, сохраняет."""
    import win32com.client

    print("\n[...] Запускаем MS Word для обновления оглавления...")
    word = None
    doc_com = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = visible
        word.DisplayAlerts = 0          # wdAlertsNone

        abs_path = os.path.abspath(output_path)
        doc_com = word.Documents.Open(abs_path)

        doc_com.Fields.Update()
        for toc in doc_com.TablesOfContents:
            toc.Update()
            toc.Range.Font.Bold = False     # содержание без жирного

        doc_com.Save()
        print("[OK] Оглавление и нумерация обновлены MS Word!")

    except Exception as e:
        print(f"[!]? Ошибка Word COM: {e}")
        if doc_com:
            try:
                doc_com.Close(False)
            except Exception:
                pass
    finally:
        if word:
            try:
                word.DisplayAlerts = -1     # wdAlertsAll
            except Exception:
                pass
