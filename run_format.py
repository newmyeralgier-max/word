import os, sys
sys.path.insert(0, r'D:\1. Project\Word\WORD\execution')
from format_docx import process_document

data_dir = r'D:\1. Project\Word\data'
for f in os.listdir(data_dir):
    if f.endswith('.docx') and 'GOST' not in f and not f.startswith('~'):
        input_path = os.path.join(data_dir, f)
        print(f"=== Processing: {f} ===")
        try:
            process_document(input_path, fast=True)
        except Exception as e:
            print(f"ERROR: {e}")
            import traceback
            traceback.print_exc()