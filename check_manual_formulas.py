"""Проверить какие формулы не детектируются как MANUAL_FORMULA"""
from docx import Document
import re, sys, os
sys.path.insert(0, '/mnt/d/1. Project/Word/WORD/execution')
from format_docx import _is_manual_formula, _RE_FORMULA_NUM, _has_math, _heading_level, _is_heading_style, STRUCTURAL_KEYWORDS

doc = Document('/mnt/d/1. Project/Word/data/БР.docx')

# Параграфы, которые выглядят как формулы но не OMML
print('=== MANUAL FORMULAS (not OMML, have = sign) ===')
for i, p in enumerate(doc.paragraphs):
    text = p.text.strip()
    if not text or len(text) < 5:
        continue
    style = p.style.name if p.style else ''
    
    # Пропускаем заголовки
    if _is_heading_style(style):
        continue
    
    # Пропускаем OMML формулы
    if _has_math(p):
        continue
    
    # Пропускаем формулы с номерами (они обрабатываются отдельно)
    if _RE_FORMULA_NUM.search(text):
        continue
    
    upper = text.upper()
    if upper in STRUCTURAL_KEYWORDS:
        continue
    
    # Ищем строки с = и числами/переменными (возможные формулы)
    has_equals = '=' in text or '≈' in text or '≥' in text or '≤' in text
    is_short = len(text) < 80
    
    if has_equals and is_short:
        is_mf = _is_manual_formula(text, style)
        if not is_mf:
            print(f'P{i}: MISSED! style={style}, text=[{text[:70]}]')
        else:
            print(f'P{i}: OK      style={style}, text=[{text[:70]}]')
