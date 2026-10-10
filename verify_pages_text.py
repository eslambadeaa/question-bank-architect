import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

def build_page_text(docx_path):
    doc = docx.Document(docx_path)
    cur_p = 1
    page_map = {}
    for p in doc.paragraphs:
        p_cur = cur_p
        for r in p.runs:
            xml = r._r.xml
            if 'lastRenderedPageBreak' in xml or 'w:type="page"' in xml:
                cur_p += 1
        t = p.text.strip()
        if t:
            for pg in range(p_cur, cur_p + 1):
                if pg not in page_map:
                    page_map[pg] = []
                page_map[pg].append(t)
    return page_map

dhab_pages = build_page_text(r'C:\Users\MaximuM-Tech\Downloads\مرجع الضبع الاسود.docx')
ejla_pages = build_page_text(r'C:\Users\MaximuM-Tech\Downloads\مرجع الايجلا.docx')

print(f"Dhab pages: {min(dhab_pages.keys())} to {max(dhab_pages.keys())}")
print(f"Ejla pages: {min(ejla_pages.keys())} to {max(ejla_pages.keys())}")

# Let's inspect pages 1 to 5 of Dhab
for p in range(1, 6):
    text = " ".join(dhab_pages.get(p, []))
    print(f"\n--- DHAB PAGE {p} ---")
    print(text[:200] + "...")
