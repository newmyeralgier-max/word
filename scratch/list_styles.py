import docx
import sys
import glob

sys.stdout.reconfigure(encoding='utf-8')

def list_styles():
    files = glob.glob(r"d:\1. Project\Word\data\Уч пособ Матлаб ПРАВИЛЬНО2.docx")
    if not files: return
    doc = docx.Document(files[0])
    for style in doc.styles:
        if style.type == docx.enum.style.WD_STYLE_TYPE.PARAGRAPH:
            print(f"Paragraph Style: {style.name}")

if __name__ == "__main__":
    list_styles()
