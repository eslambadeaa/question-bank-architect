import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

for f_path in [r'C:\Users\MaximuM-Tech\Downloads\ضبع اسود\مرجع الضبع الاسود.docx', r'C:\Users\MaximuM-Tech\Downloads\ضبع اسود\مرجع الايجلا.docx']:
    print('========================================')
    print('FILE:', f_path)
    print('========================================')
    doc = docx.Document(f_path)
    # Check paragraphs and any headings
    headings = []
    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            continue
        # Check style or run bold
        is_bold = any(r.bold for r in p.runs)
        if len(text) < 100 and (p.style.name.startswith('Heading') or is_bold):
            headings.append((i, text))
    print(f'Total headings found: {len(headings)}')
    for h in headings[:30]:
        print(f"[{h[0]}] {h[1]}")
