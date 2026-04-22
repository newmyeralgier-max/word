import docx
import sys
import glob

sys.stdout.reconfigure(encoding='utf-8')

def find_last_lines():
    files = glob.glob(r"d:\1. Project\Word\data\Уч пособ Матлаб ПРАВИЛЬНО2_MODIFIED.docx")
    if not files:
        print("File not found")
        return
    file_path = files[0]
    doc = docx.Document(file_path)
    
    # Get last 10 paragraphs with their styles
    paras = [p for p in doc.paragraphs if p.text.strip()]
    for p in paras[-10:]:
        print(f"Style: {p.style.name} | Text: {p.text}")

if __name__ == "__main__":
    find_last_lines()
