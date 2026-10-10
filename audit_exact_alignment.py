import openpyxl, sys
sys.stdout.reconfigure(encoding='utf-8')

src_file = r'C:\Users\MaximuM-Tech\Downloads\برنامج تدريب تخصص الضبع الاسود + بنوك_محدث_LMS.xlsx'
med_bank = r'C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_المتوسط_الضبع_الاسود_LMS.xlsx'
fin_bank = r'C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_النهائي_الضبع_الاسود_LMS.xlsx'

wb_prog = openpyxl.load_workbook(src_file, data_only=True)

def audit_alignment(sec_name, prog_sheet, bank_path):
    print(f"\n=======================================================")
    print(f"AUDITING: {sec_name}")
    print(f"Program Sheet: {prog_sheet}")
    print(f"Bank File: {bank_path}")
    print(f"=======================================================")
    
    # 1. Lessons from Training Program
    ws_prog = wb_prog[prog_sheet]
    prog_lessons = []
    for r in range(5, ws_prog.max_row + 1):
        lec = ws_prog.cell(r, 3).value
        name = ws_prog.cell(r, 4).value
        if lec in ['ف1', 'ف3'] and name:
            prog_lessons.append(str(name).strip())
            
    print(f"Program Lessons count: {len(prog_lessons)}")
    
    # 2. Lessons from Bank file Questions sheet
    wb_b = openpyxl.load_workbook(bank_path, data_only=True)
    ws_q = wb_b['Questions']
    bank_lessons_sequence = []
    cur_les = None
    les_counts = {}
    for r in range(2, ws_q.max_row + 1):
        les = ws_q.cell(r, 10).value
        les = str(les).strip() if les else ''
        if les != cur_les:
            bank_lessons_sequence.append(les)
            cur_les = les
        les_counts[les] = les_counts.get(les, 0) + 1
        
    print(f"Bank Lessons in Questions sheet sequence: {len(bank_lessons_sequence)}")
    
    # 3. Lessons from Bank Lookups sheet
    ws_l = wb_b['Lookups']
    lookups_lessons = []
    for r in range(2, ws_l.max_row + 1):
        val = ws_l.cell(r, 4).value
        if val:
            lookups_lessons.append(str(val).strip())
            
    print(f"Bank Lessons in Lookups sheet: {len(lookups_lessons)}")
    
    # 4. Compare 1-to-1
    mismatches = []
    for idx, (p_les, b_les, l_les) in enumerate(zip(prog_lessons, bank_lessons_sequence, lookups_lessons), start=1):
        if p_les != b_les or p_les != l_les:
            mismatches.append((idx, p_les, b_les, l_les))
            
    if not mismatches:
        print(f">> 100% PERFECT 1-TO-1 MATCH ACROSS ALL {len(prog_lessons)} LESSONS!")
        print("   First 3 lessons match:")
        for i in range(3):
            print(f"     {i+1}. {prog_lessons[i]}")
        print("   Last 3 lessons match:")
        for i in range(len(prog_lessons)-3, len(prog_lessons)):
            print(f"     {i+1}. {prog_lessons[i]}")
    else:
        print(f">> FOUND {len(mismatches)} MISMATCHES!")
        for m in mismatches[:5]:
            print(f"   Lesson {m[0]}: Prog='{m[1]}' vs Bank='{m[2]}'")

audit_alignment('القسم المتوسط', 'برنامج تدريب - القسم المتوسط', med_bank)
audit_alignment('القسم النهائي', 'برنامج تدريب - القسم النهائي', fin_bank)
