"""
Скрипт для объединения всех секций переписанного гайда в один мастер-файл.
Порядок: section1 → section2 (parts 1-4) → section3 → section4.
"""
import os

TMP_DIR = os.path.dirname(os.path.abspath(__file__))

# Порядок файлов для сборки
FILES = [
    "rewritten_guide_section1.md",
    "section2_part1.md",
    "section2_part2.md",
    "section2_part3.md",
    "section2_part4.md",
    "section3.md",
    "section4.md",
]

OUTPUT = os.path.join(TMP_DIR, "full_guide_merged.md")

def merge():
    merged_lines = []
    for fname in FILES:
        fpath = os.path.join(TMP_DIR, fname)
        if not os.path.exists(fpath):
            print(f"[!] Not found: {fname}")
            continue
        print(f"[+] Reading: {fname}")
        with open(fpath, "r", encoding="utf-8") as f:
            content = f.read()
        merged_lines.append(content)
        merged_lines.append("\n\n")
    
    with open(OUTPUT, "w", encoding="utf-8") as f:
        f.write("".join(merged_lines))
    
    total_chars = sum(len(c) for c in merged_lines)
    total_lines = sum(c.count('\n') for c in merged_lines)
    print(f"\n[OK] Master file: {OUTPUT}")
    print(f"     Chars: {total_chars:,}")
    print(f"     Lines: ~{total_lines}")

if __name__ == "__main__":
    merge()
