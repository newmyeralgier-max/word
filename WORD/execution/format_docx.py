import sys
import os
import re
import subprocess
import win32com.client
from docx import Document
from docx.shared import Pt, Cm, Mm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

from utils.docx_utils import add_page_number_to_run, generate_toc_run

# Removed add_page_number in favor of utils/docx_utils.py

def add_page_numbering(doc):
    """Добавляет нумерацию страниц по ГОСТу, ТОЛЬКО ЕСЛИ её нет в футере или хедере"""
    for section in doc.sections:
        footer = section.footer
        header = section.header
        
        # Проверяем наличие существующих полей PAGE во всём сыром XML хедера и футера!
        f_xml = footer._element.xml.upper() if footer else ""
        h_xml = header._element.xml.upper() if header else ""
        
        # Надежный поиск полей `PAGE`, `NUMPAGES` внутри `w:instrText`
        if re.search(r'INSTRTEXT.*?(PAGE|NUMPAGES)', f_xml) or re.search(r'INSTRTEXT.*?(PAGE|NUMPAGES)', h_xml) or 'PAGE<' in f_xml:
            # Номера УЖЕ есть (глубоко в XML), пропускаем секцию!
            continue
            
        # А также если есть обычные цифры
        has_text_digit = False
        if footer:
            for p in footer.paragraphs:
                if re.search(r'\d', p.text):
                    has_text_digit = True
                    break
        if has_text_digit:
            continue
            
        p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.clear()
        run = p.add_run()
        run.font.name = 'Times New Roman'
        run.font.size = Pt(14)
        add_page_number_to_run(run)

def insert_toc_smart(doc):
    """Вставляет автооглавление (TOC) непосредственно перед ВВЕДЕНИЕМ"""
    if len(doc.paragraphs) == 0:
        return
        
    toc_target = doc.paragraphs[0]
    for p in doc.paragraphs:
        if p.text.strip().upper() == "ВВЕДЕНИЕ":
            toc_target = p
            break
            
    # Заголовок СОДЕРЖАНИЕ
    p_header = toc_target.insert_paragraph_before()
    p_header.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_header.paragraph_format.space_before = Pt(24)
    p_header.paragraph_format.space_after = Pt(12)
    p_header.paragraph_format.first_line_indent = Cm(0)
    
    run_h = p_header.add_run('СОДЕРЖАНИЕ')
    run_h.bold = False
    run_h.font.name = 'Times New Roman'
    run_h.font.size = Pt(16)
    
    # Блок самого поля оглавления
    p_toc = toc_target.insert_paragraph_before()
    p_toc.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_toc.paragraph_format.first_line_indent = Cm(0)
    run_toc = p_toc.add_run()
    
    generate_toc_run(run_toc, placeholder_text="[Оглавление собрано автоматически]")
    
    # Разрыв страницы перед ВВЕДЕНИЕМ
    p_break = toc_target.insert_paragraph_before()
    run_break = p_break.add_run()
    run_break.add_break()


def format_document(input_path):
    print(f"📖 Читаем оригинальный файл: {input_path}")
    doc = Document(input_path)
    
    # === 1. Автоматическая настройка параметров ОТКЛЮЧЕНА ===
    # Принудительная перезапись section.top_margin = Mm(20) сдвигала всю верстку
    # авторских таблиц и титульных листов. Поля оставляем оригинальными!

    # === 2. Настройка базового стиля "Обычный" на весь документ (ГОСТ) ===
    style_normal = doc.styles['Normal']
    style_normal.font.name = 'Times New Roman'
    # Чтобы не раздувать таблицы до гигантских масштабов, НЕ ставим size = Pt(14) глобально.
    # Размер шрифта устанавливается точечно только к нужному тексту ниже.
    
    # Убираем лишние интервалы оглавления
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
            continue

    # === 2. Анализ и форматирование всех абзацев ===
    print("⚙️ Анализ абзацев и применение ГОСТ-стилей...")
    formatting_enabled = False
    
    for p in doc.paragraphs:
        text = p.text.strip()
        if not text:
            continue
            
        # Защита Титульного листа и Задания:
        # Пропускаем всё до первого оглавления, реферата или введения (точное совпадение строки!)
        upper_text = text.upper()
        if not formatting_enabled:
            # Ищем точное вхождение заголовка
            if re.match(r'^\s*(ВВЕДЕНИЕ|ОГЛАВЛЕНИЕ|СОДЕРЖАНИЕ|РЕФЕРАТ)\s*$', upper_text):
                formatting_enabled = True
            
        if not formatting_enabled:
            continue
            
        # Игнорируем элементы самого автооглавления, чтобы не сломать его структуру
        if p.style.name.startswith('TOC'):
            continue
            
        # Пытаемся понять, заголовок ли это
        is_heading = False
        level = 1
        
        # Заголовки по правилам не могут заканчиваться точкой, вопросом или двоеточием.
        # Более того, списки обычно длиннее обычного короткого заголовка.
        ends_with_punctuation = text.endswith(('.', ';', ':', '?', '!'))
        is_too_long_for_heading = len(text) > 120
        is_toc_title = text.upper() in ["СОДЕРЖАНИЕ", "ОГЛАВЛЕНИЕ"]
        
        if is_toc_title:
            # Делаем его красивым по центру, но БЕЗ outlineLvl, чтобы он сам не залез в авто-оглавление
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.first_line_indent = Cm(0)
            p.paragraph_format.space_before = Pt(24)
            p.paragraph_format.space_after = Pt(12)
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.keep_with_next = True
            for run in p.runs:
                run.bold = False
                run.font.size = Pt(16)
                run.font.name = 'Times New Roman'
                run.font.color.rgb = RGBColor(0, 0, 0)
            continue
            
        if not is_too_long_for_heading:
            # СТРОГАЯ защита: только конкретные ГОСТ-заголовки считаются H1 без нумерации
            strict_h1 = ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ЛИТЕРАТУРЫ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "РЕФЕРАТ", "ЗАДАНИЕ"]
            if text.strip().upper() in strict_h1 or text.strip().upper().startswith("ПРИЛОЖЕНИЕ"):
                is_heading = True
                level = 1
                
            # Проверяем, является ли это УЖЕ заголовком в самом Word (по стилю или Outline Level)
            elif p.style.name.startswith('Heading') or p.style.name.startswith('Заголовок'):
                is_heading = True
                try:
                    level = int(p.style.name.split()[-1])
                except:
                    level = 1
            else:
                pPr = p._element.pPr
                if pPr is not None:
                    outlineLvl = pPr.find(qn('w:outlineLvl'))
                    if outlineLvl is not None:
                        val = int(outlineLvl.get(qn('w:val')))
                        if val < 9:
                            is_heading = True
                            level = val + 1

        if is_heading:
            # Превращаем в ГОСТ заголовок (ГОСТ 7.32: полужирный, с абзацевого отступа)
            # ВВЕДЕНИЕ, ЗАКЛЮЧЕНИЕ и т.д. обычно центрируют
            is_strict_h1 = text.strip().upper() in ["ВВЕДЕНИЕ", "ЗАКЛЮЧЕНИЕ", "СПИСОК ЛИТЕРАТУРЫ", "СПИСОК ИСПОЛЬЗОВАННЫХ ИСТОЧНИКОВ", "РЕФЕРАТ", "ЗАДАНИЕ"] or text.strip().upper().startswith("ПРИЛОЖЕНИЕ")
            
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER if is_strict_h1 else WD_ALIGN_PARAGRAPH.JUSTIFY
            p.paragraph_format.first_line_indent = Cm(0) if is_strict_h1 else Cm(1.25)
            p.paragraph_format.line_spacing = 1.5
            p.paragraph_format.keep_with_next = True
            
            # Внедряем Outline Level для оглавления
            pPr = p._element.get_or_add_pPr()
            outlineLvl = pPr.find(qn('w:outlineLvl'))
            if outlineLvl is None:
                outlineLvl = OxmlElement('w:outlineLvl')
                pPr.append(outlineLvl)
            outlineLvl.set(qn('w:val'), str(level - 1))
            
            p.paragraph_format.space_before = Pt(18) if level > 1 else Pt(24)
            p.paragraph_format.space_after = Pt(12)
            tgt_size = Pt(16) if level == 1 else Pt(14)
            tgt_bold = True if level == 2 else False
                
            for run in p.runs:
                run.font.name = 'Times New Roman'
                run.font.size = tgt_size
                run.bold = tgt_bold
                # Не трогаем цвет, чтобы не сломать гиперссылки
                
        else:
            # Защита 1: Возврат авторской задумки (Центрированные, правые, или смещенные линейкой абзацы)
            if p.alignment in [WD_ALIGN_PARAGRAPH.CENTER, WD_ALIGN_PARAGRAPH.RIGHT]:
                continue
                
            if p.paragraph_format.left_indent is not None and p.paragraph_format.left_indent.cm > 0.5:
                # Если студент сдвинул текст линейкой (например, для шапки "Утверждаю"), пропускаем!
                continue
                
            # Защита 2: Формулы или списки с жесткой табуляцией
            if '\t' in p.text:
                continue
                
            # Защита 3: Подписи к рисункам, таблицам и "Продолжение таблицы"
            if re.match(r'^(Рисунок|Таблица|Продолжение\s*таблицы)\s*\d+', text, re.I):
                p.paragraph_format.first_line_indent = Cm(0)
                if text.lower().startswith('рисунок'):
                    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                else:
                    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
            else:
                # Обычный текст
                
                # Защита от разрушения списков (Абзац списка) с висячими отступами
                is_list = p.style.name.startswith('List') or 'numPr' in p._element.xml
                
                if not is_list:
                    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
                    p.paragraph_format.first_line_indent = Cm(1.25)
                    
                p.paragraph_format.line_spacing = 1.5
                p.paragraph_format.space_before = Pt(0)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.widow_control = True
            
            for run in p.runs:
                # Защита шрифтов исходного кода
                if run.font.name and run.font.name.lower() in ['consolas', 'courier new']:
                    continue
                    
                run.font.name = 'Times New Roman'
                if run.font.size is None:
                    run.font.size = Pt(14)
                # Если шрифт был другим, принудительно делаем 14 пт, кроме индексов
                elif run.font.size != Pt(14) and not (run.font.subscript or run.font.superscript):
                   run.font.size = Pt(14) 
                # Не закрашиваем принудительно в черный, чтобы не убить синие гиперссылки
                
    # Интеллектуальное добавление нумерации страниц (если нет)
    add_page_numbering(doc)
    
    # Интеллектуальное внедрение оглавления (если его нет)
    # python-docx игнорирует содержимое авто-океглавления (sdt), поэтому ищем в сыром XML!
    doc_xml = doc._element.xml.upper()
    
    # Надежный поиск автоматического оглавления (TOC поля внутри instrText)
    has_toc = bool(re.search(r'INSTRTEXT[^>]*>.*?TOC\s+', doc_xml)) or ('СОДЕРЖАНИЕ' in doc_xml and 'SDT' in doc_xml)
    
    # Резервная проверка на текстовое СОДЕРЖАНИЕ, если оно было вставлено вручную
    if not has_toc:
        for p in doc.paragraphs[:100]:
            if "СОДЕРЖАНИЕ" in p.text.upper() or "ОГЛАВЛЕНИЕ" in p.text.upper():
                has_toc = True
                break
                
    if not has_toc:
        print("📑 Вставляем недостающее автооглавление...")
        insert_toc_smart(doc)

    # === Версионирование сохранения ===
    out_dir = os.path.join(os.path.dirname(os.path.abspath(input_path)), 'data')
    if not os.path.exists(out_dir):
        out_dir = os.path.dirname(os.path.abspath(input_path))
        
    base_name = os.path.splitext(os.path.basename(input_path))[0]
    # Убираем суффиксы старых версий если есть
    base_name = re.sub(r'_v\d+$', '', base_name)
    base_name = re.sub(r'_GOST$', '', base_name)
    
    version = 1
    while True:
        out_name = f"{base_name}_GOST_v{version}.docx"
        output = os.path.join(out_dir, out_name)
        if not os.path.exists(output):
            break
        version += 1
        
    doc.save(output)
    print(f"\n✅ Документ отформатирован и сохранён: {os.path.abspath(output)}")
    
    # === Открытие в Word (Ультимативный расчет) ===
    print("\n⏳ Запускаем MS Word для генерации оглавления и открытия файла...")
    word = None
    doc_com = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = True  
        word.DisplayAlerts = 0  # wdAlertsNone: Защита зависания (подавление окон типа "Обновить целиком?")
        
        abs_output = os.path.abspath(output)
        doc_com = word.Documents.Open(abs_output)
        
        doc_com.Fields.Update()
        for toc in doc_com.TablesOfContents:
            toc.Update()
            
        doc_com.Save()
        print("✅ Оглавление и страницы успешно обновлены Word!")
        
    except Exception as e:
        print(f"⚠️ Ошибка при открытии файла в Word: {e}")
        if doc_com is not None:
            try:
                doc_com.Close(False)
            except:
                pass
    finally:
        if word is not None:
            try:
                word.DisplayAlerts = -1  # wdAlertsAll: Вернуть как было
                # Мы не вызываем word.Quit(), так как хотим оставить документ открытым для пользователя,
                # но защита от зависаний doc.Close() сработала бы при ошибке.
            except:
                pass

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Использование: python format_docx.py <документ.docx>")
        sys.exit(1)
        
    target_file = sys.argv[1]
    if not os.path.exists(target_file):
        print(f"Файл {target_file} не найден!")
        sys.exit(1)
        
    format_document(target_file)
