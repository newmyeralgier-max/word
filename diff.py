import fitz

d1=fitz.open('d:/1. Project/Word/СРАВНЕНИЕ/БР.pdf')
d2=fitz.open('d:/1. Project/Word/СРАВНЕНИЕ/БР_GOST_v5.pdf')

print(f'PAGES Orig: {len(d1)} vs V5: {len(d2)}')
diffs=[]

for i in range(min(70, len(d1))):
    t1=d1[i].get_text('text').strip().replace('\n', ' ')
    t2=d2[i].get_text('text').strip().replace('\n', ' ')
    if t1!=t2:
        diffs.append(f"Pg {i+1}:\nOrig: {t1[:120]}\nV5  : {t2[:120]}")

if diffs:
    print(f'TOTAL SHIFTED PAGES: {len(diffs)}')
    print('\n'.join(diffs[:6]))
else:
    print('No text differences found! Absolute perfection!')
