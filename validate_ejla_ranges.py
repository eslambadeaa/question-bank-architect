import docx
import sys

sys.stdout.reconfigure(encoding='utf-8')

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

ejla = load_doc(r'C:\Users\MaximuM-Tech\Downloads\مرجع الايجلا.docx')

ejla_ranges = [
    (1, 2, "الخواص الفنية والتكتيكية للنظام الصاروخي ايجلا"),
    (2, 3, "القدرات التكتيكية"),
    (3, 4, "دائرة الإزاحة"),
    (4, 5, "مناطق الإطلاق"),
    (5, 6, "مكونات الصاروخ الرئيسية"),
    (6, 7, "المسمار الايروديناميكي"),
    (7, 8, "منسق التتبع"),
    (8, 9, "المقاومة الفوتوغرافية"),
    (9, 10, "وحدة التحكم"),
    (10, 11, "المشغل الميكانيكي"),
    (11, 12, "مصدر التغذية المحمول"),
    (12, 13, "جهاز الإحساس بالتغير الزاوي"),
    (13, 14, "وحدة الإمداد بالغاز"),
    (14, 15, "موتور توجيه الصاروخ فى المرحلة الابتدائية"),
    (15, 16, "وحدة التسليح"),
    (16, 17, "رأس التدمير"),
    (17, 19, "وسائل الأمان"),
    (20, 21, "المحركات"),
    (21, 22, "المحرك الدافع"),
    (22, 23, "المحرك ثنائي الوظيفة"),
    (23, 23, "مؤخر عمل الشحنة"),
    (23, 24, "وحدة الأجنحة الخلفية"),
    (24, 25, "قبل مغادرة الصاروخ للقاذف"),
    (25, 27, "عمل مكونات الصاروخ أثناء الطيران"),
    (28, 29, "مجموعة الاطلاق"),
    (29, 30, "الخطوة الاولي"),
    (30, 31, "مصدر التغذية الارضي"),
    (32, 33, "القاذف")
]

print("Validating Final against Ejla text:")
for i, (sp, ep, kw) in enumerate(ejla_ranges, 1):
    found = False
    found_pages = []
    for p in ejla:
        if kw in p['text']:
            found = True
            found_pages.append(f"P{p['start_page']}")
    print(f"Lesson {i:2d}: Target P{sp}-P{ep} | Keyword '{kw}' found in {found_pages}")
