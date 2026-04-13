import fitz
doc = fitz.open('/mnt/d/1. Project/Word/data/БР_GOST.pdf')
print(f'Pages: {len(doc)}')

empty = [i+1 for i,p in enumerate(doc) if len(p.get_text().strip()) < 5]
print(f'Empty pages: {empty}')

print('\n=== TABLE CAPTIONS ===')
for i,p in enumerate(doc):
    t = p.get_text()
    for line in t.split('\n'):
        s = line.strip()
        if s.startswith('Таблица'):
            print(f'  P{i+1}: [{s[:80]}]')

print('\n=== HEADINGS (digit start) ===')
for i in range(12, len(doc)):
    t = doc[i].get_text().strip()
    lines = t.split('\n')
    for line in lines:
        s = line.strip()
        if s and s[0].isdigit() and '.' in s[:4] and len(s) > 10:
            print(f'  P{i+1}: [{s[:80]}]')

print('\n=== БИБЛИОГРАФИЯ ===')
for i,p in enumerate(doc):
    t = p.get_text()
    if 'СПИСОК' in t and ('ЛИТЕРАТУРЫ' in t or 'ИСТОЧНИКОВ' in t):
        lines = [l.strip() for l in t.split('\n') if l.strip()][:8]
        for l in lines:
            print(f'  P{i+1}: [{l}]')
        break

print('\n=== СТРУКТУРНЫЕ ===')
for keyword in ['ВВЕДЕНИЕ', 'ЗАКЛЮЧЕНИЕ']:
    for i,p in enumerate(doc):
        t = p.get_text().strip()
        if t.startswith(keyword):
            print(f'  {keyword} P{i+1}: [{t[:60]}]')
            break
