import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Load doc paragraphs with start and end pages
def load_doc(docx_path):
    doc = docx.Document(docx_path)
    cur_p = 1
    paras = []
    for i, p in enumerate(doc.paragraphs, 1):
        p_start = cur_p
        for r in p.runs:
            xml = r._r.xml
            if 'lastRenderedPageBreak' in xml or 'w:type="page"' in xml:
                cur_p += 1
        p_end = cur_p
        t = p.text.strip()
        if t:
            paras.append({'idx': i, 'start_page': p_start, 'end_page': p_end, 'text': t})
    return paras

dhab = load_doc(r'C:\Users\MaximuM-Tech\Downloads\مرجع الضبع الاسود.docx')
ejla = load_doc(r'C:\Users\MaximuM-Tech\Downloads\مرجع الايجلا.docx')

prep_ranges = [
    (1, 3, "الخواص الفنية والتكتيكية"),
    (1, 1, "فكره عامه عن المعده"),
    (3, 4, "احتياطات الأمان"),
    (5, 5, "لم  يخرج الصاروخ"),
    (6, 6, "مكونات الصاروخ"),
    (6, 7, "رأس التوجيه الذاتي"),
    (7, 8, "الجهاز العدسي الجيروسكوبي"),
    (8, 9, "قرص التعديل"),
    (9, 9, "المقاومة الفتوغرافية"),
    (10, 10, "ملفات رأس التوجيه"),
    (10, 11, "نظام أخذ العضو الدوار لسرعته"),
    (11, 12, "نظام التثبيت الكهربي"),
    (13, 14, "نظام التوجيه و التتبع"),
    (14, 15, "الدليل الالي"),
    (15, 15, "جهاز تحديد السرعة الزاوية"),
    (15, 16, "الخاص بالدفات"),
    (16, 17, "مستودع الغاز"),
    (17, 18, "مصدر التغذية المحمول"),
    (18, 19, "جهاز التفجير"),
    (18, 20, "وسائل الامان"),
    (20, 20, "راس التدمير"),
    (20, 21, "المحركات"),
    (21, 22, "المحرك الرئيسي"),
    (22, 23, "القاذف"),
    (23, 25, "مصدر التغذية الارضي"),
    (25, 27, "مجموعة الا طلاق"),
    (27, 29, "التعاون بين عناصر المعدة"),
    (29, 31, "وحدة التاخير")
]

print("Validating Preparatory against Dhab text:")
for i, (sp, ep, kw) in enumerate(prep_ranges, 1):
    found = False
    found_pages = []
    for p in dhab:
        if kw in p['text']:
            found = True
            found_pages.append(f"P{p['start_page']}")
    print(f"Lesson {i:2d}: Target P{sp}-P{ep} | Keyword '{kw}' found in {found_pages}")
