import sys, io, docx
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

doc_hyena = docx.Document(r'C:\Users\MaximuM-Tech\Downloads\مرجع الضبع الاسود.docx')
doc_igla = docx.Document(r'C:\Users\MaximuM-Tech\Downloads\مرجع الايجلا.docx')

def search_text(doc, name, terms):
    print(f'=== SEARCH IN {name} ===')
    for term in terms:
        found = []
        for i, p in enumerate(doc.paragraphs):
            if term in p.text:
                found.append((i, p.text.strip()))
        print(f'Term: "{term}" -> {len(found)} matches')
        for idx, t in found[:2]:
            print(f'   P[{idx}]: {t[:100]}...')

search_text(doc_hyena, 'ضبع اسود', ['تتك', 'التتك', 'تكتيك', 'تخزين', 'صيانة', 'أعطال', 'القتالي', 'إشارة'])
search_text(doc_igla, 'ايجلا', ['تخزين', 'صيانة', 'أعطال', 'ناشنكاه', 'تسديد', 'تشويش'])
