import sys
from docx import Document

def analyze(file_path):
    doc = Document(file_path)
    for p in doc.paragraphs:
        text = p.text.strip()
        if text.startswith('6.') or text.startswith('6 '):
            print(f"TEXT: {text[:50]}...")
            print(f"  Style: {p.style.name}")
            print(f"  Alignment: {p.alignment}")
            pf = p.paragraph_format
            print(f"  Indent: {pf.first_line_indent}")
            print(f"  Space Before: {pf.space_before}, After: {pf.space_after}")
            print(f"  Line Spacing: {pf.line_spacing}, Line Spacing Rule: {pf.line_spacing_rule}")
            if p.runs:
                font = p.runs[0].font
                print(f"  Font: {font.name}, Size: {font.size}, Bold: {font.bold}, AllCaps: {font.all_caps}")
            print("-" * 40)

if __name__ == "__main__":
    analyze(r"d:\1. Project\Word\БР.docx")
