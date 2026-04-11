"""
build_docx.py -- Сборка ГОСТ-документа из Markdown-файла.

? Парсинг Markdown через mistune (AST-рендерер)
? LaTeX-формулы ? OMML (нативные Word-уравнения)
? Автооглавление (TOC), нумерация страниц
? Все стили и токены из word_config.py / word_utils.py
"""

import sys, os, re

# ?? Путь к модулям рядом ??????????????????????????????????????????
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import word_config as cfg
import word_utils  as wu

# Зависимости (ставятся автоматически через word_utils ? ensure не нужен)
import mistune
from docx import Document
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement


# ??????????????????????????????????????????????????????????????????
#  Вспомогательные: формулы
# ??????????????????????????????????????????????????????????????????

def _add_equation(doc, latex_str, formula_number=""):
    """Display-формула (по центру)."""
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

    if formula_number:
        run = p.add_run(f"\t\t{formula_number}")
        run.font.name = cfg.FONT_NAME
        run.font.size = cfg.FONT_SIZE_MAIN
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
        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(12)
        p.paragraph_format.space_after  = Pt(6)
        p.paragraph_format.line_spacing = cfg.LINE_SPACING
        p.paragraph_format.keep_with_next = True
        run = p.add_run(caption)
        run.font.name = cfg.FONT_NAME
        run.font.size = cfg.FONT_SIZE_MAIN

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
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after  = Pt(12)
    p.paragraph_format.line_spacing = cfg.LINE_SPACING
    run = p.add_run(text)
    run.italic    = True
    run.font.name = cfg.FONT_NAME
    run.font.size = cfg.FONT_SIZE_MAIN


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
    # TOC field
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
    """Добавить заголовок с Outline Level (видимый для TOC)."""
    bold_map = {1: cfg.BOLD_H1, 2: cfg.BOLD_H2, 3: cfg.BOLD_H3}
    size_map = {1: cfg.FONT_SIZE_H1, 2: cfg.FONT_SIZE_H2, 3: cfg.FONT_SIZE_H3}

    p = doc.add_paragraph()

    # Outline Level
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
#  Rich-text paragraph
# ??????????????????????????????????????????????????????????????????

def _add_rich_paragraph(doc, text, indent=True, align='justify'):
    """
    Добавить абзац с инлайн-разметкой:
        **bold**, *italic*, `code`, $math$
    """
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
        if not part:
            continue
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


# ??????????????????????????????????????????????????????????????????
#  Авто-обёртка подстрочных индексов  T_s ? $T_s$
# ??????????????????????????????????????????????????????????????????

_SUB_RE = re.compile(
    r'(?<![A-Za-zА-Яа-яЁё0-9])'
    r'([A-Za-zА-Яа-яЁё\u03c9\u03c0]+_[A-Za-zА-Яа-яЁё0-9]+)'
    r'(?![A-Za-zА-Яа-яЁё0-9])'
)

def _auto_wrap_subscripts(text):
    parts = re.split(r'(`[^`]+?`|\$[^$]+?\$)', text)
    for i in range(0, len(parts), 2):
        if parts[i]:
            parts[i] = _SUB_RE.sub(r'$\1$', parts[i])
    return ''.join(parts)


# ??????????????????????????????????????????????????????????????????
#  Главный парсер -- PRE-PROCESS + MISTUNE AST
# ??????????????????????????????????????????????????????????????????

def _preprocess_md(raw: str) -> str:
    """
    Выделяем конструкции, которые mistune не знает:
    - [TOC]  ? оставляем как есть (отловим в AST по тексту)
    - $$...$$ display формулы внутри строки ? оборачиваем в отдельные строки
    """
    return raw


def _render_inline_children(paragraph, children):
    """Рекурсивно добавить инлайн-ноды AST в один абзац."""
    for child in children:
        tp = child['type']

        if tp == 'text':
            raw = child.get('raw', child.get('text', ''))
            # проверяем на подстрочные индексы
            raw = _auto_wrap_subscripts(raw)
            # Если внутри есть $...$  -- разделяем
            parts = re.split(r'(\$[^$]+?\$)', raw)
            for part in parts:
                if not part:
                    continue
                if part.startswith('$') and part.endswith('$'):
                    _add_inline_math(paragraph, part[1:-1])
                else:
                    run = paragraph.add_run(part)
                    run.font.name = cfg.FONT_NAME
                    run.font.size = cfg.FONT_SIZE_MAIN
        elif tp == 'strong':
            run = paragraph.add_run(child.get('raw', '') or _children_text(child))
            run.bold      = True
            run.font.name = cfg.FONT_NAME
            run.font.size = cfg.FONT_SIZE_MAIN
        elif tp == 'emphasis':
            run = paragraph.add_run(child.get('raw', '') or _children_text(child))
            run.italic    = True
            run.font.name = cfg.FONT_NAME
            run.font.size = cfg.FONT_SIZE_MAIN
        elif tp == 'codespan':
            run = paragraph.add_run(child.get('raw', child.get('text', '')))
            run.font.name = cfg.FONT_NAME_CODE
            run.font.size = cfg.FONT_SIZE_CODE
        elif tp == 'softbreak' or tp == 'linebreak':
            pass  # пропускаем мягкие переносы
        else:
            # fallback -- просто текст
            txt = child.get('raw', child.get('text', ''))
            if txt:
                run = paragraph.add_run(txt)
                run.font.name = cfg.FONT_NAME
                run.font.size = cfg.FONT_SIZE_MAIN


def _children_text(node):
    """Извлечь плоский текст из children."""
    if 'children' not in node:
        return node.get('raw', node.get('text', ''))
    return ''.join(_children_text(c) for c in node['children'])


def _flat_text(node):
    """Извлечь полный плоский текст из AST-ноды рекурсивно."""
    if isinstance(node, str):
        return node
    text = node.get('raw', '') or node.get('text', '')
    if 'children' in node and node['children']:
        text += ''.join(_flat_text(c) for c in node['children'])
    return text


def build_document():
    """Точка входа: чтение .md ? сборка .docx ? COM-обновление."""

    if len(sys.argv) > 1:
        src = sys.argv[1]
    else:
        tmp_dir = os.path.join(os.path.dirname(__file__), '..', '..', '.tmp')
        src = os.path.join(tmp_dir, 'rewritten_guide_section1.md')

    if not os.path.exists(src):
        print(f"[!] Файл не найден: {src}")
        return

    print(f"+ Исходный файл: {src}")

    with open(src, 'r', encoding='utf-8') as f:
        raw_md = f.read()

    raw_md = _preprocess_md(raw_md)

    # ?? Парсинг Markdown ? AST ????????????????????????????????????
    md = mistune.create_markdown(renderer='ast')
    ast_tokens = md(raw_md)
    print(f"  + AST-нод: {len(ast_tokens)}")

    # ?? Создание документа ????????????????????????????????????????
    doc = Document()
    for section in doc.sections:
        section.top_margin    = cfg.MARGIN_TOP
        section.bottom_margin = cfg.MARGIN_BOTTOM
        section.left_margin   = cfg.MARGIN_LEFT
        section.right_margin  = cfg.MARGIN_RIGHT

    wu.setup_gost_styles(doc)
    wu.add_page_numbering(doc, smart_skip=False)

    has_h1 = False

    # ?? Обход AST ?????????????????????????????????????????????????
    i = 0
    while i < len(ast_tokens):
        node = ast_tokens[i]
        tp = node['type']

        # ?? Заголовки ?????????????????????????????????????????????
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

        # ?? Параграф ??????????????????????????????????????????????
        if tp == 'paragraph':
            raw = _flat_text(node).strip()

            # [TOC]
            if raw.upper() in ('[TOC]', '[[TOC]]'):
                _add_toc(doc)
                i += 1
                continue

            # Подпись к рисунку: *Рисунок X.X -- ...*
            if raw.startswith('*Рисунок ') and raw.endswith('*'):
                caption_text = raw[1:-1]
                m = re.search(r'Рисунок\s+(\d+\.\d+)', caption_text)
                if m:
                    fig_num = m.group(1)
                    base = os.path.join(os.path.dirname(__file__), '..', '..', 'data', 'images')
                    for ext in ('png', 'jpg'):
                        img = os.path.join(base, f'fig{fig_num}.{ext}')
                        if os.path.exists(img):
                            _add_image_centered(doc, img)
                            break
                _add_figure_caption(doc, caption_text)
                i += 1
                continue

            # Display-формула $$...$$
            dm = re.match(r'^\$\$(.+?)\$\$\s*(\(\d+\.\d+\))?\s*$', raw)
            if dm:
                _add_equation(doc, dm.group(1).strip(), dm.group(2) or "")
                i += 1
                continue

            # Строка с $$...$$ внутри текста
            if '$$' in raw:
                parts = re.split(r'\$\$(.+?)\$\$', raw)
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = cfg.FIRST_LINE_INDENT
                p.paragraph_format.line_spacing      = cfg.LINE_SPACING
                p.paragraph_format.space_before      = Pt(0)
                p.paragraph_format.space_after       = Pt(0)
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

            # Подпись к таблице
            if re.match(r'^Таблица\s+\d', raw):
                # Запоминаем caption, следующий блок -- таблица
                caption_text = raw
                # Пробуем найти следующий токен-таблицу
                if i + 1 < len(ast_tokens) and ast_tokens[i + 1]['type'] == 'table':
                    i += 1  # перейдём к таблице, caption передадим ниже
                    node = ast_tokens[i]
                    tp = 'table'
                    node['_caption'] = caption_text
                else:
                    p = doc.add_paragraph()
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    p.paragraph_format.first_line_indent = Cm(0)
                    p.paragraph_format.keep_with_next    = True
                    run = p.add_run(caption_text)
                    run.font.name = cfg.FONT_NAME
                    run.font.size = cfg.FONT_SIZE_MAIN
                    i += 1
                    continue

            # Подпись рисунка (старый формат): *Таблица ...*
            if raw.startswith('*Таблица ') and raw.endswith('*'):
                _add_figure_caption(doc, raw[1:-1])
                i += 1
                continue

            # Обычный абзац с rich-инлайн
            if tp == 'paragraph' and 'children' in node and node['children']:
                p = doc.add_paragraph()
                p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                p.paragraph_format.first_line_indent = cfg.FIRST_LINE_INDENT
                p.paragraph_format.line_spacing      = cfg.LINE_SPACING
                p.paragraph_format.space_before      = Pt(0)
                p.paragraph_format.space_after       = Pt(0)
                _render_inline_children(p, node['children'])
                i += 1
                continue

            # Простой абзац-фолбэк
            _add_rich_paragraph(doc, _auto_wrap_subscripts(raw))
            i += 1
            continue

        # ?? Таблица ???????????????????????????????????????????????
        if tp == 'table':
            caption = node.get('_caption', '')
            header_cells = []
            data_rows    = []

            if 'children' in node:
                for child in node['children']:
                    if child['type'] == 'table_head':
                        for row in child.get('children', []):
                            header_cells = [_flat_text(c) for c in row.get('children', [])]
                    elif child['type'] == 'table_body':
                        for row in child.get('children', []):
                            data_rows.append([_flat_text(c) for c in row.get('children', [])])

            if header_cells:
                _add_table(doc, header_cells, data_rows, caption)
            i += 1
            continue

        # ?? Список ????????????????????????????????????????????????
        if tp == 'list':
            ordered = node.get('attrs', {}).get('ordered', False)
            for li_idx, item in enumerate(node.get('children', [])):
                text = _flat_text(item).strip()
                if ordered:
                    prefix = f"{li_idx + 1}. "
                else:
                    prefix = "- "
                _add_rich_paragraph(doc, _auto_wrap_subscripts(prefix + text))
            i += 1
            continue

        # ?? Блок кода ?????????????????????????????????????????????
        if tp == 'code':
            raw = node.get('raw', node.get('text', ''))
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.line_spacing      = 1.0
            p.paragraph_format.space_before      = Pt(6)
            p.paragraph_format.space_after       = Pt(6)
            run = p.add_run(raw)
            run.font.name = cfg.FONT_NAME_CODE
            run.font.size = cfg.FONT_SIZE_CODE
            i += 1
            continue

        # ?? Тематический разрыв (---) ????????????????????????????
        if tp == 'thematic_break':
            doc.add_page_break()
            i += 1
            continue

        # ?? Blank / неизвестный ????????????????????????????????????
        if tp == 'blank_line' or tp == 'newline':
            i += 1
            continue

        # fallback -- raw text
        raw = node.get('raw', node.get('text', ''))
        if raw and raw.strip():
            _add_rich_paragraph(doc, _auto_wrap_subscripts(raw.strip()))
        i += 1

    # ?? Сохранение ????????????????????????????????????????????????
    out_dir = os.path.join(os.path.dirname(__file__), '..', '..', 'data')
    os.makedirs(out_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(src))[0]
    version = 1
    while True:
        out_name = f"{base_name}_v{version}.docx"
        output   = os.path.join(out_dir, out_name)
        if not os.path.exists(output):
            break
        version += 1

    wu.save_document_safe(doc, output)
    print(f"   Параграфов: {len(doc.paragraphs)}")
    print(f"   Таблиц: {len(doc.tables)}")

    # ?? COM-обновление ????????????????????????????????????????????
    wu.update_document_via_com(output)


if __name__ == '__main__':
    build_document()
