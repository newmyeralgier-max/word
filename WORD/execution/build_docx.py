"""
build_docx.py -- Сборка ГОСТ-документа из Markdown-файла.

* Парсинг Markdown через mistune (AST-рендерер, плагины table)
* Pygments для графического выделения кода
* SEQ автонумерация таблиц, рисунков, формул
* Мультифайловая сборка
"""

import sys, os, re, glob, argparse

# ?? Путь к модулям рядом ??????????????????????????????????????????
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import word_config as cfg
import word_utils  as wu

import mistune
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement

from pygments import lex
from pygments.lexers import get_lexer_by_name, guess_lexer
from pygments.token import Token

# ??????????????????????????????????????????????????????????????????
#  Вспомогательные: Code Highlighting
# ??????????????????????????????????????????????????????????????????

def _add_source_code(doc, code, language=""):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)

    try:
        lexer = get_lexer_by_name(language) if language else guess_lexer(code)
    except Exception:
        lexer = get_lexer_by_name("text")

    for token, text in lex(code, lexer):
        run = p.add_run(text)
        run.font.name = cfg.FONT_NAME_CODE
        run.font.size = cfg.FONT_SIZE_CODE
        if token in Token.Keyword:
            run.font.color.rgb = RGBColor(0, 0, 255)
            run.bold = True
        elif token in Token.String:
            run.font.color.rgb = RGBColor(163, 21, 21)
        elif token in Token.Comment:
            run.font.color.rgb = RGBColor(0, 128, 0)
            run.italic = True
        elif token in Token.Number:
            run.font.color.rgb = RGBColor(9, 134, 88)
        elif token in Token.Name.Builtin:
            run.font.color.rgb = RGBColor(121, 94, 38)
        else:
            run.font.color.rgb = cfg.COLOR_BLACK

# ??????????????????????????????????????????????????????????????????
#  Вспомогательные: формулы
# ??????????????????????????????????????????????????????????????????

def _add_equation(doc, latex_str):
    """Display-формула (по центру). Авто-нумерация."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.line_spacing = cfg.LINE_SPACING

    omml = wu.latex_to_omml(latex_str)
    if omml is not None:
        p._element.append(omml)
    else:
        run = p.add_run(latex_str)
        run.font.name = cfg.FONT_NAME_MATH
        run.font.size = cfg.FONT_SIZE_MAIN
        run.italic = True

    # Авто-нумерация через SEQ
    run = p.add_run("\t\t(")
    run.font.name = cfg.FONT_NAME
    run.font.size = cfg.FONT_SIZE_MAIN
    wu.add_seq_field(run, "Формула")
    
    run_close = p.add_run(")")
    run_close.font.name = cfg.FONT_NAME
    run_close.font.size = cfg.FONT_SIZE_MAIN
    return p


def _add_inline_math(paragraph, latex_str):
    """Инлайн-формула внутри абзаца."""
    omml = wu.latex_to_omml(latex_str)
    if omml is not None:
        paragraph._element.append(omml)
    else:
        run = paragraph.add_run(latex_str)
        run.font.name = cfg.FONT_NAME_MATH
        run.font.size = cfg.FONT_SIZE_MAIN
        run.italic = True


# ??????????????????????????????????????????????????????????????????
#  Вспомогательные: таблицы
# ??????????????????????????????????????????????????????????????????

def _set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcPr.append(parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'</w:tcBorders>'
    ))


def _add_table(doc, header_cells, data_rows, caption=""):
    if caption:
        # Очищаем хардкод 'Таблица 1.1 -- '
        clean_caption = re.sub(r'^(?:\*?Таблица\s*[\d\.]*|\*?Таблица)\s*(?:--|-|—)?\s*', '', caption.strip('* '))
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after  = Pt(6)
        p.paragraph_format.line_spacing = cfg.LINE_SPACING
        p.paragraph_format.keep_with_next = True
        
        run = p.add_run("Таблица ")
        run.font.name = cfg.FONT_NAME
        run.font.size = cfg.FONT_SIZE_MAIN
        wu.add_seq_field(run, "Таблица")
        
        if clean_caption:
            run2 = p.add_run(" -- " + clean_caption)
            run2.font.name = cfg.FONT_NAME
            run2.font.size = cfg.FONT_SIZE_MAIN

    ncols = len(header_cells)
    table = doc.add_table(rows=len(data_rows) + 1, cols=ncols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    tl = OxmlElement('w:tblLayout')
    tl.set(qn('w:type'), 'autofit')
    table._tbl.tblPr.append(tl)

    for j, txt in enumerate(header_cells):
        cell = table.cell(0, j)
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after  = Pt(2)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(txt.strip())
        run.bold = True
        run.font.name = cfg.FONT_NAME
        run.font.size = cfg.FONT_SIZE_TABLE
        _set_cell_border(cell)

    for i, row in enumerate(data_rows):
        for j in range(ncols):
            cell = table.cell(i + 1, j)
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j > 0 else WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after  = Pt(2)
            p.paragraph_format.line_spacing = 1.0
            txt = row[j].strip() if j < len(row) else ''
            run = p.add_run(txt)
            run.font.name = cfg.FONT_NAME
            run.font.size = cfg.FONT_SIZE_TABLE
            _set_cell_border(cell)

    spacer = doc.add_paragraph()
    spacer.paragraph_format.space_before = Pt(6)
    spacer.paragraph_format.space_after  = Pt(0)
    spacer.paragraph_format.first_line_indent = Cm(0)


# ??????????????????????????????????????????????????????????????????
#  Вспомогательные: изображения и подписи
# ??????????????????????????????????????????????????????????????????

def _add_figure_caption(doc, text):
    clean_caption = re.sub(r'^(?:\*?Рисунок\s*[\d\.]*|\*?Рисунок)\s*(?:--|-|—)?\s*', '', text.strip('* '))
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(12)
    p.paragraph_format.line_spacing = cfg.LINE_SPACING
    
    run = p.add_run("Рисунок ")
    run.font.name = cfg.FONT_NAME
    run.font.size = cfg.FONT_SIZE_MAIN
    wu.add_seq_field(run, "Рисунок")
    
    if clean_caption:
        run2 = p.add_run(" -- " + clean_caption)
        run2.font.name = cfg.FONT_NAME
        run2.font.size = cfg.FONT_SIZE_MAIN


def _add_image_centered(doc, image_path):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after  = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run()
    run.add_picture(image_path, width=Cm(16))


# ??????????????????????????????????????????????????????????????????
#  TOC
# ??????????????????????????????????????????????????????????????????

def _add_toc(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    run = p.add_run('СОДЕРЖАНИЕ')
    run.bold      = False
    run.font.name = cfg.FONT_NAME
    run.font.size = cfg.FONT_SIZE_TOC_TITLE

    pt = doc.add_paragraph()
    pt.alignment = WD_ALIGN_PARAGRAPH.LEFT
    pt.paragraph_format.first_line_indent = Cm(0)
    rt = pt.add_run()
    for tag, attr, text in [
        ('w:fldChar', 'begin', None),
        ('w:instrText', None, 'TOC \\o "1-3" \\h \\z \\u'),
        ('w:fldChar', 'separate', None),
        ('w:fldChar', 'end', None),
    ]:
        el = OxmlElement(tag)
        if tag == 'w:instrText':
            el.set(qn('xml:space'), 'preserve')
            el.text = text
        else:
            el.set(qn('w:fldCharType'), attr)
        rt._r.append(el)

    doc.add_page_break()


# ??????????????????????????????????????????????????????????????????
#  Heading
# ??????????????????????????????????????????????????????????????????

def _add_heading(doc, text, level=1):
    bold_map = {1: cfg.BOLD_H1, 2: cfg.BOLD_H2, 3: cfg.BOLD_H3}
    size_map = {1: cfg.FONT_SIZE_H1, 2: cfg.FONT_SIZE_H2, 3: cfg.FONT_SIZE_H3}

    p = doc.add_paragraph()
    pPr = p._element.get_or_add_pPr()
    ol  = OxmlElement('w:outlineLvl')
    ol.set(qn('w:val'), str(level - 1))
    pPr.append(ol)

    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before      = Pt(24) if level == 1 else Pt(18)
    p.paragraph_format.space_after       = Pt(12)
    p.paragraph_format.line_spacing      = cfg.LINE_SPACING
    p.paragraph_format.keep_with_next    = True

    run = p.add_run(text.rstrip('.'))
    run.bold       = bold_map.get(level, False)
    run.font.name  = cfg.FONT_NAME
    run.font.size  = size_map.get(level, cfg.FONT_SIZE_MAIN)
    run.font.color.rgb = cfg.COLOR_BLACK
    return p


# ??????????????????????????????????????????????????????????????????
#  Rich-text paragraph / Inline rendering
# ??????????????????????????????????????????????????????????????????

def _add_rich_paragraph(doc, text, indent=True, align='justify'):
    p = doc.add_paragraph()
    ALIGN = {
        'justify': WD_ALIGN_PARAGRAPH.JUSTIFY,
        'center':  WD_ALIGN_PARAGRAPH.CENTER,
        'left':    WD_ALIGN_PARAGRAPH.LEFT,
    }
    p.alignment = ALIGN.get(align, WD_ALIGN_PARAGRAPH.JUSTIFY)
    p.paragraph_format.first_line_indent = cfg.FIRST_LINE_INDENT if indent else Cm(0)
    p.paragraph_format.line_spacing      = cfg.LINE_SPACING
    p.paragraph_format.space_before      = Pt(0)
    p.paragraph_format.space_after       = Pt(0)

    token_re = re.compile(r'(\*\*.*?\*\*|\*[^*]+?\*|`[^`]+?`|\$[^$]+?\$)')
    for part in token_re.split(text):
        if not part: continue
        if part.startswith('**') and part.endswith('**'):
            run = p.add_run(part[2:-2])
            run.bold = True
            run.font.name = cfg.FONT_NAME
            run.font.size = cfg.FONT_SIZE_MAIN
        elif part.startswith('*') and part.endswith('*'):
            run = p.add_run(part[1:-1])
            run.italic    = True
            run.font.name = cfg.FONT_NAME
            run.font.size = cfg.FONT_SIZE_MAIN
        elif part.startswith('`') and part.endswith('`'):
            run = p.add_run(part[1:-1])
            run.font.name = cfg.FONT_NAME_CODE
            run.font.size = cfg.FONT_SIZE_CODE
        elif part.startswith('$') and part.endswith('$'):
            _add_inline_math(p, part[1:-1])
        else:
            run = p.add_run(part)
            run.font.name = cfg.FONT_NAME
            run.font.size = cfg.FONT_SIZE_MAIN
    return p

_SUB_RE = re.compile(
    r'(?<![A-Za-zА-Яа-яЁё0-9])([A-Za-zА-Яа-яЁё\u03c9\u03c0]+_[A-Za-zА-Яа-яЁё0-9]+)(?![A-Za-zА-Яа-яЁё0-9])'
)

def _auto_wrap_subscripts(text):
    parts = re.split(r'(`[^`]+?`|\$[^$]+?\$)', text)
    for i in range(0, len(parts), 2):
        if parts[i]:
            parts[i] = _SUB_RE.sub(r'$\1$', parts[i])
    return ''.join(parts)


def _render_inline_children(paragraph, children):
    for child in children:
        tp = child['type']
        raw = child.get('raw', child.get('text', ''))
        
        if tp == 'text':
            raw = _auto_wrap_subscripts(raw)
            parts = re.split(r'(\$[^$]+?\$)', raw)
            for part in parts:
                if not part: continue
                if part.startswith('$') and part.endswith('$'):
                    _add_inline_math(paragraph, part[1:-1])
                else:
                    run = paragraph.add_run(part)
                    run.font.name = cfg.FONT_NAME
                    run.font.size = cfg.FONT_SIZE_MAIN
        elif tp == 'strong':
            run = paragraph.add_run(raw or _children_text(child))
            run.bold      = True
            run.font.name = cfg.FONT_NAME
            run.font.size = cfg.FONT_SIZE_MAIN
        elif tp == 'emphasis':
            run = paragraph.add_run(raw or _children_text(child))
            run.italic    = True
            run.font.name = cfg.FONT_NAME
            run.font.size = cfg.FONT_SIZE_MAIN
        elif tp == 'codespan':
            run = paragraph.add_run(raw)
            run.font.name = cfg.FONT_NAME_CODE
            run.font.size = cfg.FONT_SIZE_CODE
        elif tp in ('softbreak', 'linebreak'):
            pass
        else:
            if raw:
                run = paragraph.add_run(raw)
                run.font.name = cfg.FONT_NAME
                run.font.size = cfg.FONT_SIZE_MAIN

def _children_text(node):
    if 'children' not in node:
        return node.get('raw', node.get('text', ''))
    return ''.join(_children_text(c) for c in node['children'])

def _flat_text(node):
    if isinstance(node, str): return node
    text = node.get('raw', '') or node.get('text', '')
    if 'children' in node and node['children']:
        text += ''.join(_flat_text(c) for c in node['children'])
    return text


def _read_inputs(input_path):
    """Слияние Markdown файлов из директории или чтение одного файла."""
    if os.path.isdir(input_path):
        files = sorted(glob.glob(os.path.join(input_path, "*.md")))
        if not files:
            print(f"[!] MD файлы не найдены в папке {input_path}")
            sys.exit(1)
        texts = []
        for f in files:
            with open(f, 'r', encoding='utf-8') as fid:
                texts.append(fid.read())
        print(f"+ Склеено файлов: {len(files)}")
        return "\n\n---\n\n".join(texts), files[0]
    else:
        with open(input_path, 'r', encoding='utf-8') as f:
            return f.read(), input_path

# ??????????????????????????????????????????????????????????????????
#  Главная сборка
# ??????????????????????????????????????????????????????????????????

def build_document(input_path, output=None, no_com=False):
    if not os.path.exists(input_path):
        print(f"[!] Путь не найден: {input_path}")
        return

    raw_md, sample_path = _read_inputs(input_path)

    base_dir = os.path.dirname(os.path.abspath(sample_path))
    images_dir = os.path.join(base_dir, '..', 'data', 'images')
    if not os.path.exists(images_dir):
        images_dir = base_dir

    md = mistune.create_markdown(renderer='ast', plugins=['table'])
    ast_tokens = md(raw_md)
    print(f"  + AST-нод: {len(ast_tokens)}")

    doc = Document()
    for section in doc.sections:
        section.top_margin    = cfg.MARGIN_TOP
        section.bottom_margin = cfg.MARGIN_BOTTOM
        section.left_margin   = cfg.MARGIN_LEFT
        section.right_margin  = cfg.MARGIN_RIGHT

    wu.setup_gost_styles(doc)
    wu.add_page_numbering(doc, smart_skip=False)

    has_h1 = False
    i = 0
    while i < len(ast_tokens):
        node = ast_tokens[i]
        tp = node['type']

        if tp == 'heading':
            level = node['attrs']['level']
            text  = _flat_text(node).strip()
            if level == 1:
                if has_h1:
                    doc.add_page_break()
                has_h1 = True
            _add_heading(doc, text, level=min(level, 3))
            i += 1
            continue

        if tp == 'paragraph':
            raw = _flat_text(node).strip()

            if raw.upper() in ('[TOC]', '[[TOC]]'):
                _add_toc(doc)
                i += 1
                continue

            # Подпись Рисунка
            if raw.startswith('*Рисунок ') and raw.endswith('*'):
                caption_text = raw[1:-1]
                m = re.search(r'Рисунок\s+(\d+\.\d+)', caption_text)
                if m:
                    fig_num = m.group(1)
                    for ext in ('png', 'jpg'):
                        img = os.path.join(images_dir, f'fig{fig_num}.{ext}')
                        if os.path.exists(img):
                            _add_image_centered(doc, img)
                            break
                _add_figure_caption(doc, caption_text)
                i += 1
                continue

            # Формулы
            dm = re.match(r'^\$\$(.+?)\$\$\s*(\(\d+\.\d+\))?\s*$', raw)
            if dm:
                _add_equation(doc, dm.group(1).strip())
                i += 1
                continue

            if '$$' in raw:
                parts = re.split(r'\$\$(.+?)\$\$', raw)
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = cfg.FIRST_LINE_INDENT
                p.paragraph_format.line_spacing      = cfg.LINE_SPACING
                for idx, part in enumerate(parts):
                    if idx % 2 == 0:
                        if part.strip():
                            run = p.add_run(part)
                            run.font.name = cfg.FONT_NAME
                            run.font.size = cfg.FONT_SIZE_MAIN
                    else:
                        _add_inline_math(p, part.strip())
                i += 1
                continue

            # Таблицы (caption)
            if re.match(r'^Таблица\s+\d', raw) or re.match(r'^\*Таблица\s+\d', raw):
                caption_text = raw
                if i + 1 < len(ast_tokens) and ast_tokens[i + 1]['type'] == 'table':
                    i += 1
                    node = ast_tokens[i]
                    tp = 'table'
                    node['_caption'] = caption_text
                else:
                    _add_figure_caption(doc, caption_text)
                    i += 1
                    continue

            if tp == 'paragraph' and 'children' in node and node['children']:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = cfg.FIRST_LINE_INDENT
                p.paragraph_format.line_spacing      = cfg.LINE_SPACING
                _render_inline_children(p, node['children'])
                i += 1
                continue

            _add_rich_paragraph(doc, _auto_wrap_subscripts(raw))
            i += 1
            continue

        if tp == 'table':
            caption = node.get('_caption', '')
            header_cells = []
            data_rows    = []
            if 'children' in node:
                for child in node['children']:
                    if child['type'] == 'table_head':
                        header_cells = [_flat_text(c) for c in child.get('children', [])]
                    elif child['type'] == 'table_body':
                        for row in child.get('children', []):
                            data_rows.append([_flat_text(c) for c in row.get('children', [])])
            if header_cells:
                _add_table(doc, header_cells, data_rows, caption)
            i += 1
            continue

        if tp == 'list':
            ordered = node.get('attrs', {}).get('ordered', False)
            for li_idx, item in enumerate(node.get('children', [])):
                text = _flat_text(item).strip()
                prefix = f"{li_idx + 1}. " if ordered else "-- "
                _add_rich_paragraph(doc, _auto_wrap_subscripts(prefix + text))
            i += 1
            continue

        if tp == 'code':
            raw = node.get('raw', node.get('text', ''))
            info = node.get('attrs', {}).get('info', '')
            _add_source_code(doc, raw, info)
            i += 1
            continue

        if tp == 'thematic_break':
            doc.add_page_break()
            i += 1
            continue

        if tp in ('blank_line', 'newline'):
            i += 1
            continue

        raw = node.get('raw', node.get('text', ''))
        if raw and raw.strip():
            _add_rich_paragraph(doc, _auto_wrap_subscripts(raw.strip()))
        i += 1

    if output is None:
        out_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', '..', 'data')
        os.makedirs(out_dir, exist_ok=True)
        base = os.path.splitext(os.path.basename(sample_path))[0]
        v = 1
        while True:
            output = os.path.join(out_dir, f"{base}_builder_v{v}.docx")
            if not os.path.exists(output): break
            v += 1

    wu.save_document_safe(doc, output)
    print(f"   Параграфов: {len(doc.paragraphs)}")
    print(f"   Таблиц: {len(doc.tables)}")

    if not no_com:
        wu.update_document_via_com(output)

if __name__ == '__main__':
    parser = argparse.ArgumentParser(description="Сборка Markdown в Word GOST")
    parser.add_argument('-i', '--input', help="Путь к файлу или папке", default=None)
    parser.add_argument('-o', '--output', help="Путь для сохранения", default=None)
    parser.add_argument('--no-com', action='store_true', help="Пропустить обновление через MS Word")
    
    args = parser.parse_args()
    
    input_val = args.input
    if not input_val:
        input_val = os.path.join(os.path.dirname(__file__), '..', '..', '.tmp', 'rewritten_guide_section1.md')
        
    build_document(input_val, args.output, args.no_com)
