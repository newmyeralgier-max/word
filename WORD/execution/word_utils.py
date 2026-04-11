import os
import sys
import subprocess
import glob
import shutil
import re

# === Автоустановка зависимостей ===
def ensure_package(pkg, pip_name=None):
    try:
        __import__(pkg)
    except ImportError:
        d_null = open(os.devnull, 'w')
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', pip_name or pkg, '-q'], stdout=d_null, stderr=d_null)

# Установка базовых библиотек
ensure_package('docx', 'python-docx')
ensure_package('win32com', 'pywin32')
ensure_package('lxml')
ensure_package('mistune') # для build_docx.py
ensure_package('latex2mathml')

import win32com.client
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree
import word_config as cfg

# === Стилизация ===
def setup_gost_styles(doc):
    """Инициализация встроенных стилей Word под ГОСТ для избежания прямого форматирования."""
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = cfg.FONT_NAME
    font.size = cfg.FONT_SIZE_MAIN
    font.color.rgb = RGBColor(0, 0, 0)
    pf = style_normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.5
    pf.first_line_indent = Cm(1.25)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.widow_control = True

    for i in range(1, 4):
        h_name = f'Heading {i}'
        try:
            h_style = doc.styles[h_name]
            h_style.font.name = cfg.FONT_NAME
            h_style.font.color.rgb = RGBColor(0, 0, 0)
            h_style.font.bold = True if i == 2 else False
            h_style.font.size = cfg.FONT_SIZE_H1 if i == 1 else cfg.FONT_SIZE_MAIN
            h_style.paragraph_format.space_before = Pt(24) if i == 1 else Pt(18)
            h_style.paragraph_format.space_after = Pt(12)
            h_style.paragraph_format.line_spacing = 1.5
            h_style.paragraph_format.keep_with_next = True
            h_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            h_style.paragraph_format.first_line_indent = Cm(0)
        except KeyError:
            pass

        ru_name = f'Заголовок {i}'
        try:
            ru_style = doc.styles[ru_name]
            ru_style.font.name = cfg.FONT_NAME
            ru_style.font.color.rgb = RGBColor(0, 0, 0)
            ru_style.font.bold = True if i == 2 else False
            ru_style.font.size = cfg.FONT_SIZE_H1 if i == 1 else cfg.FONT_SIZE_MAIN
            ru_style.paragraph_format.space_before = Pt(24) if i == 1 else Pt(18)
            ru_style.paragraph_format.space_after = Pt(12)
            ru_style.paragraph_format.line_spacing = 1.5
            ru_style.paragraph_format.keep_with_next = True
            ru_style.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            ru_style.paragraph_format.first_line_indent = Cm(0)
        except KeyError:
            pass

    for i in range(1, 4):
        try:
            toc_style = doc.styles[f'TOC {i}']
            toc_style.font.name = cfg.FONT_NAME
            toc_style.font.size = cfg.FONT_SIZE_MAIN
            toc_style.font.bold = False
            toc_style.paragraph_format.space_before = Pt(0)
            toc_style.paragraph_format.space_after = Pt(0)
            toc_style.paragraph_format.line_spacing = 1.5
        except KeyError:
            pass

# === MML2OMML ===
def get_local_xsl_path():
    return os.path.join(os.path.dirname(__file__), "MML2OMML.XSL")

_xsl_transform = None

def get_omml_xslt():
    global _xsl_transform
    if _xsl_transform is not None:
        return _xsl_transform
        
    local_path = get_local_xsl_path()
    
    if not os.path.exists(local_path):
        patterns = [
            r"C:\Program Files\Microsoft Office\root\Office*\MML2OMML.XSL",
            r"C:\Program Files (x86)\Microsoft Office\root\Office*\MML2OMML.XSL",
            r"C:\Program Files\Microsoft Office\Office*\MML2OMML.XSL",
            r"C:\Program Files (x86)\Microsoft Office\Office*\MML2OMML.XSL",
        ]
        found = None
        for pat in patterns:
            matches = glob.glob(pat)
            if matches:
                found = matches[-1]
                break
                
        if found:
            print(f"  Копируется XSLT из {found} в локальную директорию...")
            shutil.copy(found, local_path)
        else:
            print("  ⚠ MML2OMML.XSL не найден на ПК — формулы в Word будут текстовыми")
            _xsl_transform = False
            return False

    try:
        xsl_tree = etree.parse(local_path)
        _xsl_transform = etree.XSLT(xsl_tree)
    except Exception as e:
        print(f"  ⚠ Ошибка парсинга {local_path}: {e}")
        _xsl_transform = False
        
    return _xsl_transform

# === Базовые утилиты Word ===
def add_page_number(run):
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = "PAGE"
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def add_page_numbering(doc):
    for section in doc.sections:
        footer = section.footer
        header = section.header
        
        f_xml = footer._element.xml.upper() if footer else ""
        h_xml = header._element.xml.upper() if header else ""
        
        if re.search(r'INSTRTEXT.*?(PAGE|NUMPAGES)', f_xml) or re.search(r'INSTRTEXT.*?(PAGE|NUMPAGES)', h_xml) or 'PAGE<' in f_xml:
            continue
            
        has_text_digit = False
        if footer:
            for p in footer.paragraphs:
                if re.search(r'\d', p.text):
                    has_text_digit = True
                    break
        if has_text_digit:
            continue
            
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT 
        p.clear()
        run = p.add_run()
        run.font.name = cfg.FONT_NAME
        run.font.size = cfg.FONT_SIZE_MAIN
        add_page_number(run)

def save_document_safe(doc, output_path):
    saved = False
    while not saved:
        try:
            doc.save(output_path)
            saved = True
        except PermissionError:
            print(f"\n[ОШИБКА] Файл {output_path} открыт в другой программе (скорее всего в MS Word).")
            input("Пожалуйста, закройте файл в Word и нажмите Enter, чтобы продолжить...")
            
    print(f"\n✅ Документ сохранён: {os.path.abspath(output_path)}")
    return True

def update_toc_and_page_numbers_via_com(output_path):
    print("\n⏳ Запускаем MS Word для актуализации оглавления...")
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        word.DisplayAlerts = 0 
        
        try:
            abs_output = os.path.abspath(output_path)
            doc_com = word.Documents.Open(abs_output)
            
            doc_com.Fields.Update()
            
            for toc in doc_com.TablesOfContents:
                toc.Update()
                toc.Range.Font.Bold = False 
                
            doc_com.Save()
            print("✅ Документ успешно обновлен MS Word!")
        finally:
            word.DisplayAlerts = -1
    except Exception as e:
        print(f"⚠️ Ошибка при открытии файла в Word: {e}")
