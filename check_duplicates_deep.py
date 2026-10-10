import openpyxl
import collections
import sys

sys.stdout.reconfigure(encoding='utf-8')

files = [
    ('القسم المتوسط', r'C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_المتوسط_الضبع_الاسود_LMS.xlsx'),
    ('القسم النهائي', r'C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_النهائي_الضبع_الاسود_LMS.xlsx')
]

for sec_name, fpath in files:
    print(f"\n=======================================================")
    print(f"فحص التكرارات التفصيلي لـ: {sec_name}")
    print(f"=======================================================")
    wb = openpyxl.load_workbook(fpath, data_only=True)
    ws = wb['Questions']
    
    # نفحص التكرار داخل نفس الدرس، والتكرار عبر الملف ككل
    questions_by_text = collections.defaultdict(list)
    questions_by_lesson_and_text = collections.defaultdict(list)
    
    for r in range(2, ws.max_row + 1):
        q_type = ws.cell(r, 1).value
        q_text = str(ws.cell(r, 2).value).strip() if ws.cell(r, 2).value else ''
        q_ans = ws.cell(r, 5).value
        q_les = ws.cell(r, 10).value
        
        questions_by_text[q_text].append((r, q_les, q_type, q_ans))
        questions_by_lesson_and_text[(q_les, q_text)].append((r, q_type, q_ans))
        
    # تكرار داخل نفس الدرس
    same_lesson_dups = {k: v for k, v in questions_by_lesson_and_text.items() if len(v) > 1}
    print(f"عدد الأسئلة المكررة داخل نفس الدرس: {len(same_lesson_dups)}")
    if same_lesson_dups:
        for (les, txt), v in list(same_lesson_dups.items())[:10]:
            print(f"  - في درس '{les}': السؤال '{txt[:40]}...' تكرر {len(v)} مرات في الصفوف {[x[0] for x in v]}")
            
    # تكرار بين الدروس المختلفة
    cross_lesson_dups = {k: v for k, v in questions_by_text.items() if len(v) > 1}
    print(f"عدد الأسئلة المكررة عبر الملف ككل: {len(cross_lesson_dups)}")
    if cross_lesson_dups:
        print("عينة من الأسئلة المكررة عبر الدروس:")
        for txt, v in list(cross_lesson_dups.items())[:5]:
            lessons_involved = set(x[1] for x in v)
            print(f"  - السؤال: '{txt[:45]}...' تكرر {len(v)} مرات في الدروس: {lessons_involved}")
