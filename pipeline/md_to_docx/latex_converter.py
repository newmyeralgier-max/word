"""
latex_converter.py — Конвертация формул в LaTeX-нотацию.

Принимает текстовую формулу и выводит в двойном формате:
1. Обычная формула (с подставленными числами)
2. LaTeX-блок кода

Также может сканировать .docx файл и извлекать все формулы.

Использование:
    python latex_converter.py --formula "r = h * tg(α)"
    python latex_converter.py --formula "r = 15 * tg(45°) = 15 м" --fix-decimals
    python latex_converter.py --scan "файл.docx"

Зависимости: python-docx (для --scan)
"""

import argparse
import re
import sys


# Таблица замен для конвертации текста → LaTeX
LATEX_REPLACEMENTS = [
    # Греческие буквы
    ("α", r"\alpha"),
    ("β", r"\beta"),
    ("γ", r"\gamma"),
    ("δ", r"\delta"),
    ("ε", r"\varepsilon"),
    ("η", r"\eta"),
    ("θ", r"\theta"),
    ("λ", r"\lambda"),
    ("μ", r"\mu"),
    ("π", r"\pi"),
    ("ρ", r"\rho"),
    ("σ", r"\sigma"),
    ("τ", r"\tau"),
    ("φ", r"\varphi"),
    ("ω", r"\omega"),
    ("Ω", r"\Omega"),
    ("Δ", r"\Delta"),
    ("Σ", r"\Sigma"),

    # Математические операции
    ("·", r"\cdot"),
    ("×", r"\times"),
    ("÷", r"\div"),
    ("±", r"\pm"),
    ("≈", r"\approx"),
    ("≠", r"\neq"),
    ("≤", r"\leq"),
    ("≥", r"\geq"),
    ("∞", r"\infty"),
    ("√", r"\sqrt"),
    ("°", r"^\circ"),

    # Тригонометрические функции
    ("tg", r"\tan"),
    ("ctg", r"\cot"),
    ("sin", r"\sin"),
    ("cos", r"\cos"),
    ("arcsin", r"\arcsin"),
    ("arccos", r"\arccos"),
    ("arctg", r"\arctan"),
    ("lg", r"\lg"),
    ("ln", r"\ln"),
    ("log", r"\log"),
]

# Русские единицы измерения
UNITS = [
    "кВт", "кВА", "кВАр", "МВА", "МВт", "Вт",
    "кВ", "В", "мВ",
    "кА", "А", "мА",
    "кОм", "Ом", "мОм",
    "Гц", "кГц", "МГц",
    "м", "см", "мм", "км",
    "с", "мс", "мин", "ч",
    "кг", "т", "Н",
    "Дж", "кДж",
]


def text_to_latex(formula: str) -> str:
    """
    Конвертирует текстовую формулу в LaTeX.
    
    Args:
        formula: Текстовая формула, например "r = h · tg(α)"
        
    Returns:
        LaTeX-строка, например "r = h \\cdot \\tan(\\alpha)"
    """
    result = formula

    # Заменяем единицы измерения на \text{...}
    for unit in sorted(UNITS, key=len, reverse=True):
        result = re.sub(
            rf'(?<=[0-9\s,.]){re.escape(unit)}(?=[\s.,;)\]}}]|$)',
            rf'\\,\\text{{{unit}}}',
            result
        )

    # Заменяем математические символы и функции
    for text_val, latex_val in LATEX_REPLACEMENTS:
        result = result.replace(text_val, latex_val)

    # Обработка дробей: a/b → \frac{a}{b}
    # Поддержка дробей с десятичными числами (точка или запятая)
    result = re.sub(
        r'([\w\.,]+)\s*/\s*([\w\.,]+)',
        r'\\frac{\1}{\2}',
        result
    )

    # Обработка индексов
    result = re.sub(r'_([\w]+)', r'_{\1}', result)

    # Обработка степеней
    result = re.sub(r'\^([\w]+)', r'^{\1}', result)

    return result


def fix_decimal_dots(formula: str) -> str:
    """Заменяет точки на запятые в десятичных числах."""
    return re.sub(r'(\d)\.(\d)', r'\1,\2', formula)


def format_dual_output(original: str, latex: str) -> str:
    """
    Форматирует формулу в двойном формате:
    1. Обычная формула
    2. LaTeX-блок кода
    """
    output = f"{original}\n\n```latex\n{latex}\n```"
    return output


def scan_docx_for_formulas(filepath: str) -> list[dict]:
    """
    Сканирует .docx файл и ищет строки, похожие на формулы.
    Критерии: содержит знак =, математические символы, греческие буквы.
    
    Returns:
        Список словарей с полями: index, text, latex
    """
    from docx import Document

    doc = Document(filepath)
    formulas = []

    formula_indicators = r'[=+\-*/·×÷√∑∫αβγδεηθλμπρστφωΩΔΣ]'
    
    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if not text:
            continue
        
        # Проверяем: содержит ли параграф формулу?
        has_equals = "=" in text
        has_math = bool(re.search(formula_indicators, text))
        is_short = len(text) < 200  # Формулы обычно короче обычного текста
        has_digits = bool(re.search(r'\d', text))
        
        if is_short and (has_equals or has_math) and has_digits:
            latex = text_to_latex(text)
            formulas.append({
                "index": i,
                "text": text,
                "latex": latex,
            })

    return formulas


def main():
    parser = argparse.ArgumentParser(description="Конвертация формул в LaTeX")
    parser.add_argument("--formula", help="Текстовая формула для конвертации")
    parser.add_argument("--scan", help="Путь к .docx файлу для сканирования формул")
    parser.add_argument("--fix-decimals", action="store_true",
                        help="Заменить точки на запятые в числах")
    args = parser.parse_args()

    if args.formula:
        formula = args.formula
        if args.fix_decimals:
            formula = fix_decimal_dots(formula)
        
        latex = text_to_latex(formula)
        print(format_dual_output(formula, latex))

    elif args.scan:
        from pathlib import Path
        if not Path(args.scan).exists():
            print(f"ОШИБКА: Файл не найден: {args.scan}", file=sys.stderr)
            sys.exit(1)

        formulas = scan_docx_for_formulas(args.scan)
        if not formulas:
            print("Формулы не найдены.")
        else:
            print(f"Найдено формул: {len(formulas)}\n")
            for f in formulas:
                print(f'--- Параграф [{f["index"]}] ---')
                print(format_dual_output(f["text"], f["latex"]))
                print()
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
