import docx
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

for fname in ['مرجع الضبع الاسود.docx', 'مرجع الايجلا.docx']:
    fpath = os.path.join(r'C:\Users\MaximuM-Tech\Downloads', fname)
    doc = docx.Document(fpath)
    print(f"\n==================== {fname} ====================")
    cur_p = 1
    for i, p in enumerate(doc.paragraphs, 1):
        for r in p.runs:
            xml = r._r.xml
            if 'lastRenderedPageBreak' in xml or 'w:type="page"' in xml:
                cur_p += 1
        t = p.text.strip()
        # print headings or short structural paragraphs
        if t and (len(t) < 60 or p.style.name.startswith('Heading')):
            print(f"P{i:3d} [Page {cur_p:2d}]: {t}")
