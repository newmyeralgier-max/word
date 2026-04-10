"""
docx_utils.py — Общие утилиты для работы с Word-документами.

Библиотека вспомогательных функций, используемая другими скриптами.
НЕ запускается как CLI.

Зависимости: python-docx
"""

import os
import shutil
from datetime import datetime
from pathlib import Path
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement


def find_reference_file(directory: str) -> str | None:
    """
    Находит эталонный файл в директории.
    Эталон определяется по наличию слова 'эталон' в имени файла.
    
    Args:
        directory: Путь к директории поиска
        
    Returns:
        Путь к эталонному файлу или None если не найден
    """
    for f in Path(directory).glob("*.docx"):
        if "эталон" in f.name.lower():
            return str(f)
    return None


def find_working_files(directory: str) -> list[str]:
    """
    Находит все рабочие .docx файлы (не эталонные, не временные).
    
    Args:
        directory: Путь к директории поиска
        
    Returns:
        Список путей к рабочим файлам
    """
    result = []
    for f in Path(directory).glob("*.docx"):
        name_lower = f.name.lower()
        if "эталон" not in name_lower and not f.name.startswith("~") and not f.name.startswith("backup_"):
            result.append(str(f))
    return result


def create_backup(filepath: str) -> str:
    """
    Создаёт резервную копию файла с временной меткой.
    
    Args:
        filepath: Путь к файлу
        
    Returns:
        Путь к созданной резервной копии
    """
    path = Path(filepath)
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_name = f"backup_{timestamp}_{path.name}"
    backup_path = path.parent / backup_name
    shutil.copy2(filepath, backup_path)
    return str(backup_path)


def get_paragraph_by_text(doc: Document, search_text: str, exact: bool = False) -> list:
    """
    Ищет параграфы по тексту.
    
    Args:
        doc: Объект Document
        search_text: Текст для поиска
        exact: Если True — точное совпадение, иначе — содержит подстроку
        
    Returns:
        Список кортежей (индекс, параграф)
    """
    results = []
    for i, para in enumerate(doc.paragraphs):
        if exact and para.text.strip() == search_text.strip():
            results.append((i, para))
        elif not exact and search_text.lower() in para.text.lower():
            results.append((i, para))
    return results


def get_paragraphs_by_style(doc: Document, style_name: str) -> list:
    """
    Ищет параграфы по названию стиля.
    
    Args:
        doc: Объект Document
        style_name: Название стиля (Heading 1, Normal и т.д.)
        
    Returns:
        Список кортежей (индекс, параграф)
    """
    results = []
    for i, para in enumerate(doc.paragraphs):
        if para.style and para.style.name == style_name:
            results.append((i, para))
    return results


def extract_run_formatting(run) -> dict:
    """
    Извлекает форматирование из run-объекта.
    
    Args:
        run: Объект Run из python-docx
        
    Returns:
        Словарь с параметрами форматирования
    """
    fmt = {
        "text": run.text,
        "bold": run.bold,
        "italic": run.italic,
        "underline": run.underline,
        "font_name": None,
        "font_size": None,
        "font_color": None,
    }
    if run.font:
        fmt["font_name"] = run.font.name
        fmt["font_size"] = str(run.font.size) if run.font.size else None
        if run.font.color and run.font.color.rgb:
            fmt["font_color"] = str(run.font.color.rgb)
    return fmt


def fix_decimal_separators(text: str) -> str:
    """
    Заменяет точки на запятые в десятичных числах.
    Пример: '32.09' → '32,09', '0.75' → '0,75'
    Не трогает точки в конце предложений и между словами.
    
    Args:
        text: Исходный текст
        
    Returns:
        Текст с замененными разделителями
    """
    import re
    # Заменяем только точки между цифрами: 32.09 → 32,09
    return re.sub(r'(\d)\.(\d)', r'\1,\2', text)


def log_operation(operation: str, details: str, log_dir: str = ".tmp"):
    """
    Логирует операцию в файл.
    
    Args:
        operation: Название операции
        details: Детали
        log_dir: Директория для логов
    """
    os.makedirs(log_dir, exist_ok=True)
    log_path = os.path.join(log_dir, "operations.log")
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    with open(log_path, "a", encoding="utf-8") as f:
        f.write(f"[{timestamp}] {operation}: {details}\n")


def add_page_number_to_run(run):
    """
    Добавляет поле номера страницы (PAGE) в указанный run (OxmlElement).
    """
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


def generate_toc_run(run, placeholder_text="[Оглавление собрано автоматически]"):
    """
    Добавляет поле автоматического оглавления (TOC \\o "1-3" \\h \\z \\u) в указанный run.
    """
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    
    instrText = OxmlElement('w:instrText')
    instrText.set(qn('xml:space'), 'preserve')
    instrText.text = 'TOC \\o "1-3" \\h \\z \\u'
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'separate')
    
    placeholder_run = OxmlElement('w:r')
    placeholder_t = OxmlElement('w:t')
    placeholder_t.text = placeholder_text
    placeholder_run.append(placeholder_t)
    
    fldChar3 = OxmlElement('w:fldChar')
    fldChar3.set(qn('w:fldCharType'), 'end')
    
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(placeholder_run)
    run._r.append(fldChar3)
