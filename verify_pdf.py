import fitz

doc1 = fitz.open('d:/1. Project/Word/СРАВНЕНИЕ/БР.pdf')
doc2 = fitz.open('d:/1. Project/Word/СРАВНЕНИЕ/БР_GOST_v2.pdf')

print(f"Orig: {len(doc1)} pages, V2: {len(doc2)} pages\n")

diffs = []
for i in range(min(45, len(doc1))):
    t1 = doc1[i].get_text('text')[:300].replace('\n', ' ')
    t2 = doc2[i].get_text('text')[:300].replace('\n', ' ')
    if t1 != t2:
        diffs.append(f"Pg {i+1} Orig:\t{t1[:100]}...\nPg {i+1} V2:\t{t2[:100]}...\n")

if not diffs:
    print("First 45 pages start identical in content.")
else:
    for d in diffs[:8]:
        print(d)
        
print("Analysis complete.")
