import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

def print_all_pages(doc_path):
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
            pages[current_page].append(text)
            
    print(f"File: {doc_path} -> Total Pages: {len(pages)}")
    for p_num in range(1, len(pages) + 1):
        content = ' | '.join(pages.get(p_num, []))[:120]
        print(f"P{p_num:02d}: {content}")

print("=== مرجع الضبع الاسود ===")
print_all_pages(r'C:\Users\MaximuM-Tech\Downloads\ضبع اسود\مرجع الضبع الاسود.docx')
