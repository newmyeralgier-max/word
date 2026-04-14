"""Проверить реальный порядок элементов в body — где таблицы относительно параграфов"""
from docx import Document
from docx.oxml.ns import qn

doc = Document('/mnt/d/1. Project/Word/data/БР.docx')
body = doc._element.find(qn('w:body'))

# Ищем таблицы и их соседей
tables_found = 0
for i, elem in enumerate(body):
    if elem.tag == qn('w:tbl'):
        tables_found += 1
        # Печатаем 3 элемента ДО таблицы
        print(f'\n=== TABLE #{tables_found} (element index {i}) ===')
        for j in range(max(0, i-3), i):
            prev = list(body)[j]
            if prev.tag == qn('w:p'):
                texts = prev.findall('.//' + qn('w:t'))
                text = ''.join(t.text or '' for t in texts)[:80]
                print(f'  BEFORE [{j}]: P text=[{text}]')
            elif prev.tag == qn('w:tbl'):
                print(f'  BEFORE [{j}]: TABLE')
        
        # И 2 элемента ПОСЛЕ
        for j in range(i+1, min(i+3, len(list(body)))):
            nxt = list(body)[j]
            if nxt.tag == qn('w:p'):
                texts = nxt.findall('.//' + qn('w:t'))
                text = ''.join(t.text or '' for t in texts)[:80]
                print(f'  AFTER  [{j}]: P text=[{text}]')
            elif nxt.tag == qn('w:tbl'):
                print(f'  AFTER  [{j}]: TABLE')

# То же для GOST
print('\n\n========== GOST ==========')
doc_g = Document('/mnt/d/1. Project/Word/data/БР_GOST.docx')
body_g = doc_g._element.find(qn('w:body'))

tables_found = 0
for i, elem in enumerate(body_g):
    if elem.tag == qn('w:tbl'):
        tables_found += 1
        if tables_found <= 10:  # only first 10
            print(f'\n=== GOST TABLE #{tables_found} (element index {i}) ===')
            elems = list(body_g)
            for j in range(max(0, i-3), i):
                prev = elems[j]
                if prev.tag == qn('w:p'):
                    texts = prev.findall('.//' + qn('w:t'))
                    text = ''.join(t.text or '' for t in texts)[:80]
                    print(f'  BEFORE [{j}]: P text=[{text}]')
                elif prev.tag == qn('w:tbl'):
                    print(f'  BEFORE [{j}]: TABLE')
            for j in range(i+1, min(i+3, len(elems))):
                nxt = elems[j]
                if nxt.tag == qn('w:p'):
                    texts = nxt.findall('.//' + qn('w:t'))
                    text = ''.join(t.text or '' for t in texts)[:80]
                    print(f'  AFTER  [{j}]: P text=[{text}]')
                elif nxt.tag == qn('w:tbl'):
                    print(f'  AFTER  [{j}]: TABLE')
