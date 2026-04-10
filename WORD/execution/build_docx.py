"""
Скрипт для сборки переписанного гайда по MATLAB/Simulink в формат .docx
Форматирование по ГОСТ 7.32-2017 (для методических указаний):
- Шрифт: Times New Roman, 14 пт (основной текст), 16 пт (заголовки разделов)
- Интервал: 1.5 строки
- Отступ первой строки: 1.25 см
- Поля: левое 30 мм, правое 15 мм, верхнее/нижнее 20 мм
- Выравнивание основного текста: по ширине
- Заголовки разделов: по центру, полужирный
- Формулы: нативные Word-формулы (OMML) через LaTeX → MathML → OMML
"""

import subprocess, sys, os, re, glob

# === Автоустановка зависимостей ===
def ensure_package(pkg, pip_name=None):
    try:
        __import__(pkg)
    except ImportError:
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', pip_name or pkg, '-q'])

ensure_package('docx', 'python-docx')
ensure_package('latex2mathml')
ensure_package('win32com', 'pywin32')

import win32com.client
from docx import Document
from docx.shared import Pt, Cm, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml, OxmlElement
from lxml import etree
import latex2mathml.converter


# ======================================================================
#  OMML (Office Math Markup) — конвертер LaTeX → Word-формулы
# ======================================================================

def find_mml2omml_xsl():
    """Найти XSLT-файл MML2OMML.XSL из установки Microsoft Office."""
    patterns = [
        r"C:\Program Files\Microsoft Office\root\Office*\MML2OMML.XSL",
        r"C:\Program Files (x86)\Microsoft Office\root\Office*\MML2OMML.XSL",
        r"C:\Program Files\Microsoft Office\Office*\MML2OMML.XSL",
        r"C:\Program Files (x86)\Microsoft Office\Office*\MML2OMML.XSL",
    ]
    for pat in patterns:
        found = glob.glob(pat)
        if found:
            return found[0]
    return None

# Кэш XSLT для повторного использования
_xsl_transform = None

def _get_xslt():
    global _xsl_transform
    if _xsl_transform is not None:
        return _xsl_transform
    
    xsl_path = find_mml2omml_xsl()
    if xsl_path:
        xsl_tree = etree.parse(xsl_path)
        _xsl_transform = etree.XSLT(xsl_tree)
        print(f"  OMML XSLT найден: {xsl_path}")
    else:
        _xsl_transform = False  # Маркер «не найдено»
        print("  ⚠ MML2OMML.XSL не найден — формулы будут в текстовом виде")
    return _xsl_transform


def latex_to_omml(latex_str):
    """Конвертировать LaTeX → OMML XML-элемент для вставки в Word."""
    xslt = _get_xslt()
    if not xslt:
        return None
    
    try:
        # LaTeX → MathML
        mathml_str = latex2mathml.converter.convert(latex_str)
        mathml_tree = etree.fromstring(mathml_str.encode('utf-8'))
        
        # MathML → OMML
        omml_tree = xslt(mathml_tree)
        return omml_tree.getroot()
    except Exception as e:
        print(f"  ⚠ Ошибка конвертации формулы: {e}")
        return None


def add_equation(doc, latex_str, formula_number=""):
    """Добавить формулу в документ как нативное Word-уравнение (OMML)."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    
    omml = latex_to_omml(latex_str)
    if omml is not None:
        # Вставляем OMML-элемент в параграф
        p._element.append(omml)
    else:
        # Фоллбэк: текстовая формула
        run = p.add_run(latex_str)
        run.font.name = 'Cambria Math'
        run.font.size = Pt(14)
        run.italic = True
    
    # Номер формулы справа
    if formula_number:
        run = p.add_run(f"\t\t{formula_number}")
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
    
    return p


def add_inline_math(paragraph, latex_str):
    """Добавить инлайн-формулу в существующий параграф как OMML."""
    omml = latex_to_omml(latex_str)
    if omml is not None:
        paragraph._element.append(omml)
    else:
        # Фоллбэк: курсив Cambria Math
        run = paragraph.add_run(latex_str)
        run.font.name = 'Cambria Math'
        run.font.size = Pt(14)
        run.italic = True


# ======================================================================
#  ГОСТ-форматирование документа
# ======================================================================

def create_gost_document():
    doc = Document()
    for section in doc.sections:
        section.top_margin = Mm(20)
        section.bottom_margin = Mm(20)
        section.left_margin = Mm(30)
        section.right_margin = Mm(15)
    
    style_normal = doc.styles['Normal']
    font = style_normal.font
    font.name = 'Times New Roman'
    font.size = Pt(14)
    font.color.rgb = RGBColor(0, 0, 0)
    pf = style_normal.paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 1.5
    pf.first_line_indent = Cm(1.25)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.widow_control = True  # Контроль висячих строк
    
    # Настраиваем стили оглавления, чтобы они не были разреженными
    for i in range(1, 4):
        try:
            toc_style = doc.styles[f'TOC {i}']
            toc_style.font.name = 'Times New Roman'
            toc_style.font.size = Pt(14)
            toc_style.font.bold = False
            toc_style.paragraph_format.space_before = Pt(0)
            toc_style.paragraph_format.space_after = Pt(0)
            toc_style.paragraph_format.line_spacing = 1.5
        except KeyError:
            pass
            
    add_page_numbering(doc)
    return doc


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
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.RIGHT  # Изменено по просьбе: справа
        p.clear()
        run = p.add_run()
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        add_page_number(run)


def add_heading(doc, text, level=1):
    p = doc.add_paragraph()
    
    # Внедряем Outline Level для автооглавления (TOC)
    pPr = p._element.get_or_add_pPr()
    outlineLvl = OxmlElement('w:outlineLvl')
    outlineLvl.set(qn('w:val'), str(level - 1))
    pPr.append(outlineLvl)

    # ВСЕ заголовки по центру и без отступа первой строки
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.keep_with_next = True

    # ГОСТ Форматирование
    if level == 1:
        # H1: По центру, 16 пт, НЕ жирный
        p.paragraph_format.space_before = Pt(24)
        font_size = Pt(16)
        is_bold = False
    elif level == 2:
        # H2: По центру, 14 пт, жирный
        p.paragraph_format.space_before = Pt(18)
        font_size = Pt(14)
        is_bold = True
    else:
        # H3: По центру, 14 пт, НЕ жирный
        p.paragraph_format.space_before = Pt(18)
        font_size = Pt(14)
        is_bold = False
    
    run = p.add_run(text.rstrip('.'))
    run.bold = is_bold
    run.font.name = 'Times New Roman'
    run.font.size = font_size
    run.font.color.rgb = RGBColor(0, 0, 0)


def add_paragraph_rich(doc, text, indent=True, align='justify'):
    """
    Добавить абзац с поддержкой:
    - **bold**  →  полужирный
    - *italic*  →  курсив
    - `code`    →  Courier New 12 пт
    - $latex$   →  инлайн OMML-формула
    """
    p = doc.add_paragraph()
    if align == 'justify':
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif align == 'center':
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    
    p.paragraph_format.first_line_indent = Cm(1.25) if indent else Cm(0)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    
    # Разбиваем текст на части: **bold**, *italic*, `code`, $math$
    # Порядок важен: сначала ** (чтобы не перехватил *), затем *, затем `, затем $
    token_re = re.compile(r'(\*\*.*?\*\*|\*[^*]+?\*|`[^`]+?`|\$[^$]+?\$)')
    parts = token_re.split(text)
    
    for part in parts:
        if not part:
            continue
        if part.startswith('**') and part.endswith('**'):
            run = p.add_run(part[2:-2])
            run.bold = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
        elif part.startswith('*') and part.endswith('*'):
            run = p.add_run(part[1:-1])
            run.italic = True
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
        elif part.startswith('`') and part.endswith('`'):
            run = p.add_run(part[1:-1])
            run.font.name = 'Courier New'
            run.font.size = Pt(12)
        elif part.startswith('$') and part.endswith('$'):
            latex = part[1:-1]
            add_inline_math(p, latex)
        else:
            run = p.add_run(part)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(14)
    
    return p


def add_figure_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(12)
    p.paragraph_format.line_spacing = 1.5
    run = p.add_run(text)
    run.italic = True
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)
def add_image_centered(doc, image_path):
    """Вставляет картинку по центру документа."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.keep_with_next = True
    run = p.add_run()
    # Ограничиваем максимальную ширину картинки
    run.add_picture(image_path, width=Cm(16))


def add_toc(doc):
    """Добавить автособираемое оглавление."""
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    # ГОСТ: СОДЕРЖАНИЕ прописными буквами
    run = p.add_run('СОДЕРЖАНИЕ')
    run.bold = False
    run.font.name = 'Times New Roman'
    run.font.size = Pt(16)
    
    p_toc = doc.add_paragraph()
    p_toc.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_toc.paragraph_format.first_line_indent = Cm(0)
    run_toc = p_toc.add_run()
    
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    run_toc._r.append(fldChar1)
    run_toc._r.append(instrText)
    run_toc._r.append(fldChar2)
    run_toc._r.append(fldChar3)
    
    # Добавляем разрыв страницы после TOC
    doc.add_page_break()





# ======================================================================
#  Таблицы
# ======================================================================

def set_cell_border(cell):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = parse_xml(
        f'<w:tcBorders {nsdecls("w")}>'
        f'  <w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'  <w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'  <w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'  <w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        f'</w:tcBorders>'
    )
    tcPr.append(tcBorders)


def add_table_caption(doc, text):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.first_line_indent = Cm(0)
    p.paragraph_format.space_before = Pt(12)
    p.paragraph_format.space_after = Pt(6)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.keep_with_next = True
    run = p.add_run(text)
    run.font.name = 'Times New Roman'
    run.font.size = Pt(14)


def add_markdown_table(doc, header_cells, data_rows):
    num_cols = len(header_cells)
    table = doc.add_table(rows=len(data_rows) + 1, cols=num_cols)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = True
    
    # Позволяем Word самому распределять ширину колонок
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'autofit')
    table._tbl.tblPr.append(tblLayout)
    
    for j, cell_text in enumerate(header_cells):
        cell = table.cell(0, j)
        cell.text = ''
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.space_before = Pt(2)
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.0
        run = p.add_run(cell_text.strip())
        run.bold = True
        run.font.name = 'Times New Roman'
        run.font.size = Pt(12)
        set_cell_border(cell)
    
    for i, row_cells in enumerate(data_rows):
        for j in range(num_cols):
            cell = table.cell(i + 1, j)
            cell.text = ''
            p = cell.paragraphs[0]
            p.alignment = WD_ALIGN_PARAGRAPH.LEFT if j > 0 else WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(2)
            p.paragraph_format.space_after = Pt(2)
            p.paragraph_format.line_spacing = 1.0
            text = row_cells[j].strip() if j < len(row_cells) else ''
            run = p.add_run(text)
            run.font.name = 'Times New Roman'
            run.font.size = Pt(12)
            set_cell_border(cell)
    
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(0)
    p.paragraph_format.first_line_indent = Cm(0)


# ======================================================================
#  Главный парсер
# ======================================================================

def auto_wrap_subscripts(text):
    """
    Автоматически находит переменные с нижним индексом (например: T_s, U_н, ω_б)
    и оборачивает их в $...$ для рендеринга как OMML-формулы, 
    если они ещё не обернуты.
    """
    # Разделяем текст на уже защищенные блоки (код и формулы) и обычный текст
    parts = re.split(r'(`[^`]+?`|\$[^$]+?\$)', text)
    for i in range(0, len(parts), 2):
        if not parts[i]: continue
        # Ищем: (не буква/цифра) + (буквы/греческие) + _ + (буквы/цифры) + (не буква/цифра)
        parts[i] = re.sub(r'(?<![A-Za-zА-Яа-яЁё0-9])([A-Za-zА-Яа-яЁё\u03c9\u03c0]+_[A-Za-zА-Яа-яЁё0-9]+)(?![A-Za-zА-Яа-яЁё0-9])', r'$\1$', parts[i])
    return ''.join(parts)


def build_document():
    import sys
    
    if len(sys.argv) > 1:
        src = sys.argv[1]
    else:
        tmp_dir = os.path.join(os.path.dirname(__file__), '..', '..', '.tmp')
        src = os.path.join(tmp_dir, 'rewritten_guide_section1.md')
        
    if not os.path.exists(src):
        print(f"Файл не найден: {src}")
        return

    print(f"Исходный файл: {src}")
    
    # Считаем общее число строк
    with open(src, 'r', encoding='utf-8') as f:
        lines = f.readlines()
        
    print(f"  Прочитано {len(lines)} строк.")
    
    doc = create_gost_document()
    
    has_h1 = False

    i = 0
    while i < len(lines):
        line = lines[i].rstrip('\n').rstrip('\r')
        
        if not line.strip():
            i += 1
            continue
            
        # --- Оглавление (TOC) ---
        if line.strip().upper() == '[TOC]' or line.strip() == '[[TOC]]':
            add_toc(doc)
            i += 1
            continue
        
        # --- Заголовок H1 ---
        if line.startswith('# ') and not line.startswith('## '):
            if has_h1:
                doc.add_page_break()
            has_h1 = True
            add_heading(doc, line[2:].strip(), level=1)
            i += 1
            continue
        
        # --- Заголовок H2 ---
        if line.startswith('## ') and not line.startswith('### '):
            add_heading(doc, line[3:].strip(), level=2)
            i += 1
            continue
        
        # --- Заголовок H3 ---
        if line.startswith('### '):
            add_heading(doc, line[4:].strip(), level=3)
            i += 1
            continue
        
        # --- Подпись к рисунку ---
        if line.strip().startswith('*Рисунок ') and line.strip().endswith('*'):
            caption_text = line.strip()[1:-1]
            # Пытаемся найти номер рисунка для авто-изображения
            m = re.search(r'Рисунок\s+(\d+\.\d+)', caption_text)
            if m:
                fig_num = m.group(1)
                img_path_png = os.path.join(os.path.dirname(__file__), '..', '..', 'data', 'images', f'fig{fig_num}.png')
                img_path_jpg = os.path.join(os.path.dirname(__file__), '..', '..', 'data', 'images', f'fig{fig_num}.jpg')
                if os.path.exists(img_path_png):
                    add_image_centered(doc, img_path_png)
                elif os.path.exists(img_path_jpg):
                    add_image_centered(doc, img_path_jpg)
                    
            add_figure_caption(doc, caption_text)
            i += 1
            continue
        
        # --- Таблица (подпись + markdown table) ---
        if re.match(r'^Таблица \d', line.strip()):
            caption_text = line.strip()
            i += 1
            while i < len(lines) and not lines[i].strip():
                i += 1
            table_lines = []
            while i < len(lines) and lines[i].strip().startswith('|'):
                table_lines.append(lines[i].strip())
                i += 1
            if table_lines:
                header = [c.strip() for c in table_lines[0].split('|')[1:-1]]
                data = [[c.strip() for c in tl.split('|')[1:-1]] for tl in table_lines[2:]]
                add_table_caption(doc, caption_text)
                add_markdown_table(doc, header, data)
            else:
                add_table_caption(doc, caption_text)
            continue
        
        # --- Подпись к таблице (старый формат *Таблица ...*) ---
        if line.strip().startswith('*Таблица ') and line.strip().endswith('*'):
            add_figure_caption(doc, line.strip()[1:-1])
            i += 1
            continue
        
        # --- Display-формула: $$latex$$ (номер) ---
        display_match = re.match(r'^\$\$(.+?)\$\$\s*(\(\d+\.\d+\))?\s*$', line.strip())
        if display_match:
            latex = display_match.group(1).strip()
            number = display_match.group(2) or ""
            add_equation(doc, latex, number)
            i += 1
            continue
        
        # --- Строка с $$формулой$$ внутри текста ---
        if '$$' in line.strip():
            # Разделяем текст и формулу
            parts = re.split(r'\$\$(.+?)\$\$', line.strip())
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(1.25)
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.space_before = Pt(0)
            p.paragraph_format.space_after = Pt(0)
            for idx, part in enumerate(parts):
                if idx % 2 == 0:
                    # Обычный текст
                    if part.strip():
                        run = p.add_run(part)
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(14)
                else:
                    # Формула
                    add_inline_math(p, part.strip())
            i += 1
            continue
        
        # --- Элемент списка через тире ---
        if line.strip().startswith('– ') or line.strip().startswith('- '):
            txt = line.strip()[2:] if line.strip().startswith('– ') else line.strip()[2:]
            add_paragraph_rich(doc, auto_wrap_subscripts('– ' + txt))
            i += 1
            continue
        
        # --- Элемент списка буквами типа "а)", "б)", "в)" ---
        letter_match = re.match(r'^([а-яА-ЯёЁ]\))\s+(.+)$', line.strip())
        if letter_match:
            add_paragraph_rich(doc, auto_wrap_subscripts(f"{letter_match.group(1)} {letter_match.group(2)}"))
            i += 1
            continue

        # --- Нумерованный список (не 1.1, 1.2, ...) ---
        num_match = re.match(r'^(\d+)\.\s+(.+)$', line.strip())
        if num_match and not re.match(r'^1\.\d', line.strip()):
            add_paragraph_rich(doc, auto_wrap_subscripts(f"{num_match.group(1)}. {num_match.group(2)}"))
            i += 1
            continue
        
        # --- Элемент описания рисунка: "1 — текст" ---
        if re.match(r'^\d+\s*—\s*.+$', line.strip()):
            add_paragraph_rich(doc, auto_wrap_subscripts(line.strip()))
            i += 1
            continue
        
        # --- Обычный абзац (с поддержкой $inline$ математики) ---
        add_paragraph_rich(doc, auto_wrap_subscripts(line.strip()))
        i += 1
    
    out_dir = os.path.join(os.path.dirname(__file__), '..', '..', 'data')
    os.makedirs(out_dir, exist_ok=True)
    
    base_name = os.path.splitext(os.path.basename(src))[0]
    
    version = 1
    while True:
        out_name = f"{base_name}_v{version}.docx"
        output = os.path.join(out_dir, out_name)
        if not os.path.exists(output):
            break
        version += 1
        
    doc.save(output)
    print(f"\n✅ Документ сохранён: {os.path.abspath(output)}")
    print(f"   Параграфов: {len(doc.paragraphs)}")
    print(f"   Таблиц: {len(doc.tables)}")
    
    # === Открываем Word и обновляем оглавление программно ===
    print("\n⏳ Запускаем MS Word для генерации номеров страниц и открытия файла...")
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True
        
        abs_output = os.path.abspath(output)
        doc_com = word.Documents.Open(abs_output)
        
        # Обновляем все поля
        doc_com.Fields.Update()
        
        # Прямое обновление оглавления
        for toc in doc_com.TablesOfContents:
            toc.Update()
            toc.Range.Font.Bold = False
            
        doc_com.Save()
        print("✅ Оглавление и страницы успешно обновлены Word!")
        
    except Exception as e:
        print(f"⚠️ Ошибка при открытии файла в Word (возможно не установлен MS Office): {e}")


if __name__ == '__main__':
    build_document()
