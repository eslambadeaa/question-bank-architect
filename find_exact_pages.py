import docx
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

for name in ['مرجع الضبع الاسود.docx', 'مرجع الايجلا.docx']:
    path = os.path.join(r'C:\Users\MaximuM-Tech\Downloads', name)
    doc = docx.Document(path)
    
    pb = 0
    lrpb = 0
    for p in doc.paragraphs:
        for r in p.runs:
            xml = r._r.xml
            if 'w:type="page"' in xml:
                pb += 1
            if 'lastRenderedPageBreak' in xml:
                lrpb += 1
    print(name, f"PageBreaks={pb}, LastRenderedPageBreaks={lrpb}")

# Check if Microsoft Word is available via COM
try:
    import win32com.client
    print("win32com is installed, testing Word application...")
    word = win32com.client.Dispatch("Word.Application")
    print("Word.Application successfully dispatched!")
    word.Quit()
except Exception as e:
    print(f"Word COM check: {e}")
