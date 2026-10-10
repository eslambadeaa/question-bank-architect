import docx
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')

def extract_paragraphs_with_exact_pages(docx_path):
    doc = docx.Document(docx_path)
    current_page = 1
    para_page_map = []
    
    for p_idx, p in enumerate(doc.paragraphs, start=1):
        text = p.text.strip()
        p_start_page = current_page
        
        # Check all runs in paragraph for page breaks
        for r in p.runs:
            xml = r._r.xml
            if 'w:lastRenderedPageBreak' in xml or 'w:type="page"' in xml:
                current_page += 1
                
        p_end_page = current_page
        if text:
            para_page_map.append({
                'para_idx': p_idx,
                'text': text,
                'start_page': p_start_page,
                'end_page': p_end_page
            })
            
    print(f"Total parsed pages for {os.path.basename(docx_path)}: {current_page} (expected 33)")
    return para_page_map, current_page

for fname in ['مرجع الضبع الاسود.docx', 'مرجع الايجلا.docx']:
    fpath = os.path.join(r'C:\Users\MaximuM-Tech\Downloads', fname)
    pmap, total_p = extract_paragraphs_with_exact_pages(fpath)
    print(f"\nFirst 5 paragraphs of {fname}:")
    for item in pmap[:5]:
        print(f"  P{item['para_idx']} (Page {item['start_page']}-{item['end_page']}): {item['text'][:50]}")
    print(f"\nLast 5 paragraphs of {fname}:")
    for item in pmap[-5:]:
        print(f"  P{item['para_idx']} (Page {item['start_page']}-{item['end_page']}): {item['text'][:50]}")
