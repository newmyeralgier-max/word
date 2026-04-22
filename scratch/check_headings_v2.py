import docx
import sys
import glob

sys.stdout.reconfigure(encoding='utf-8')

def check_headings():
    files = glob.glob(r"d:\1. Project\Word\data\Уч пособ Матлаб ПРАВИЛЬНО2.docx")
    if not files: return
    doc = docx.Document(files[0])
    for para in doc.paragraphs:
        if para.style.name.startswith('Heading'):
            print(f"{para.style.name}: {para.text}")

if __name__ == "__main__":
    check_headings()
