import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

fpath = r'C:\Users\MaximuM-Tech\Downloads\مرجع الضبع الاسود.docx'
doc = docx.Document(fpath)
cur_p = 1
for i, p in enumerate(doc.paragraphs, 1):
    for r in p.runs:
        xml = r._r.xml
        if 'lastRenderedPageBreak' in xml or 'w:type="page"' in xml:
            cur_p += 1
    t = p.text.strip()
    if t and (len(t) < 65 or p.style.name.startswith('Heading')):
        print(f"P{i:3d} [Page {cur_p:2d}]: {t}")
