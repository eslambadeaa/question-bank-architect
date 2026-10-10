import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

def extract_detailed_page_info(doc_path):
    doc = docx.Document(doc_path)
    pages = {}
    current_page = 1
    pages[current_page] = []
    
    for p in doc.paragraphs:
        for r in p.runs:
            xml = r._r.xml
            if 'w:lastRenderedPageBreak' in xml or ('w:br' in xml and 'type="page"' in xml):
                current_page += 1
                pages[current_page] = []
        text = p.text.strip()
        if text:
            # check if bold
            is_bold = any(r.bold for r in p.runs)
            pages[current_page].append((text, is_bold))
    return pages

pages_h = extract_detailed_page_info(r'C:\Users\MaximuM-Tech\Downloads\ضبع اسود\مرجع الضبع الاسود.docx')
print(f"=== Black Hyena Pages: {len(pages_h)} ===")
for p_num, items in sorted(pages_h.items()):
    bolds = [t for t, b in items if b and len(t) < 80]
    first_text = items[0][0] if items else ""
    print(f"Page {p_num:02d}: Bolds: {bolds[:3]} | First: {first_text[:60]}")

pages_i = extract_detailed_page_info(r'C:\Users\MaximuM-Tech\Downloads\ضبع اسود\مرجع الايجلا.docx')
print(f"\n=== Igla Pages: {len(pages_i)} ===")
for p_num, items in sorted(pages_i.items()):
    bolds = [t for t, b in items if b and len(t) < 80]
    first_text = items[0][0] if items else ""
    print(f"Page {p_num:02d}: Bolds: {bolds[:3]} | First: {first_text[:60]}")
