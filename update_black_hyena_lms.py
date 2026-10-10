import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
import sys

sys.stdout.reconfigure(encoding='utf-8')

backup_path = r'C:\Users\MaximuM-Tech\Downloads\برنامج تدريب تخصص الضبع الاسود + بنوك_backup.xlsx'
target_path = r'C:\Users\MaximuM-Tech\Downloads\برنامج تدريب تخصص الضبع الاسود + بنوك.xlsx'

wb = openpyxl.load_workbook(backup_path, data_only=False)
print('Workbook loaded from backup. Sheets:', wb.sheetnames)

# 1. Ensure Lookups sheet exists and is properly populated
if 'Lookups' not in wb.sheetnames:
    ws_lookups = wb.create_sheet(title='Lookups')
else:
    ws_lookups = wb['Lookups']

lookups_data = [
    ['أنواع الأسئلة', 'مستويات الصعوبة', 'نماذج الإجابة الصحيحة', 'دروس المقرر'],
    ['1 - اختيار من متعدد', '1 - سهل', 'A', None],
    ['2 - صح/خطأ', '2 - متوسط', 'B', None],
    ['3 - مقالي', '3 - صعب', 'C', None],
    ['4 - مطابقة', '4 - صعب جدا', 'D', None],
    ['5 - أكمل الفراغ', '5 - تفوق', 'A;B', None],
    ['6 - ترتيب', None, 'A;C', None],
    ['7 - متعدد الإجابات', None, 'True', None],
    [None, None, 'False', None],
]

for r_idx, row in enumerate(lookups_data, start=1):
    for c_idx, val in enumerate(row, start=1):
        ws_lookups.cell(r_idx, c_idx, val)

# Set Lookups defined names matching LMS template
wb.defined_names['L_1'] = DefinedName('L_1', attr_text="Lookups!$A$2:$A$8")
wb.defined_names['L_2'] = DefinedName('L_2', attr_text="Lookups!$B$2:$B$6")
wb.defined_names['L_3'] = DefinedName('L_3', attr_text="Lookups!$C$2:$C$9")

# Difficulty sequence per 25-question lesson:
# 30% easy (7), 30% medium (8), 15% hard (4), 15% very hard (4), 10% excellence (2) = 25
DIFF_MAP = {
    1: '1 - سهل', 2: '1 - سهل', 3: '1 - سهل', 4: '1 - سهل', 5: '1 - سهل', 6: '1 - سهل', 7: '1 - سهل',
    8: '2 - متوسط', 9: '2 - متوسط', 10: '2 - متوسط', 11: '2 - متوسط', 12: '2 - متوسط', 13: '2 - متوسط', 14: '2 - متوسط', 15: '2 - متوسط',
    16: '3 - صعب', 17: '3 - صعب', 18: '3 - صعب', 19: '3 - صعب',
    20: '4 - صعب جدا', 21: '4 - صعب جدا', 22: '4 - صعب جدا', 23: '4 - صعب جدا',
    24: '5 - تفوق', 25: '5 - تفوق'
}

bank_sheets = [s for s in wb.sheetnames if s.startswith('بنك')]
print('Bank sheets to update:', bank_sheets)

for bs in bank_sheets:
    ws = wb[bs]
    max_r = ws.max_row
    print(f'\n--- Updating {bs} ({max_r - 1} questions) ---')
    
    # Group rows by lesson (25 questions each)
    cur_lesson = None
    lesson_rows = []
    all_lessons = []
    for r in range(2, max_r + 1):
        les = ws.cell(r, 10).value
        if les != cur_lesson:
            if lesson_rows:
                all_lessons.append((cur_lesson, lesson_rows))
            cur_lesson = les
            lesson_rows = [r]
        else:
            lesson_rows.append(r)
    if lesson_rows:
        all_lessons.append((cur_lesson, lesson_rows))
        
    print(f'Found {len(all_lessons)} lessons.')
    
    for les_idx, (les_name, rows) in enumerate(all_lessons):
        assert len(rows) == 25, f'Lesson {les_name} has {len(rows)} rows instead of 25!'
        
        # Read the 25 questions data
        mcq_items = []
        tf_items = []
        
        for r in rows:
            optA = str(ws.cell(r, 6).value).strip() if ws.cell(r, 6).value is not None else ''
            optB = str(ws.cell(r, 7).value).strip() if ws.cell(r, 7).value is not None else ''
            optC = ws.cell(r, 8).value
            curr_type = str(ws.cell(r, 1).value).strip() if ws.cell(r, 1).value is not None else ''
            curr_ans = str(ws.cell(r, 5).value).strip() if ws.cell(r, 5).value is not None else ''
            q_text = ws.cell(r, 2).value
            q_exp = ws.cell(r, 3).value
            q_optC_val = ws.cell(r, 8).value
            q_optD_val = ws.cell(r, 9).value
            
            is_tf = (optA in ['True', 'صواب', 'صح'] and optB in ['False', 'خطأ', 'خطا'] and (optC is None or str(optC).strip() in ['', 'None', '-'])) or (curr_type in ['صواب أو خطأ', 'صح/خطأ', '2 - صح/خطأ'])
            
            if is_tf:
                # Determine normalized TF answer: True or False
                if curr_ans in ['الخيار أ (A)', 'A', 'أ', 'صواب', 'True', 'true', 'صح']:
                    norm_ans = 'True'
                elif curr_ans in ['الخيار ب (B)', 'B', 'ب', 'خطأ', 'False', 'false', 'خطا']:
                    norm_ans = 'False'
                else:
                    raise ValueError(f'Unknown TF answer in {bs} row {r}: {curr_ans}')
                
                tf_items.append({
                    'text': q_text,
                    'exp': q_exp,
                    'ans': norm_ans,
                    'optA': 'True',
                    'optB': 'False',
                    'optC': None,
                    'optD': None,
                    'lesson': les_name
                })
            else:
                # Determine normalized MCQ answer: A, B, C, D
                if curr_ans in ['الخيار أ (A)', 'A', 'أ']:
                    norm_ans = 'A'
                elif curr_ans in ['الخيار ب (B)', 'B', 'ب']:
                    norm_ans = 'B'
                elif curr_ans in ['الخيار ج (C)', 'C', 'ج']:
                    norm_ans = 'C'
                elif curr_ans in ['الخيار د (D)', 'D', 'د']:
                    norm_ans = 'D'
                else:
                    raise ValueError(f'Unknown MCQ answer in {bs} row {r}: {curr_ans}')
                    
                mcq_items.append({
                    'text': q_text,
                    'exp': q_exp,
                    'ans': norm_ans,
                    'optA': ws.cell(r, 6).value,
                    'optB': ws.cell(r, 7).value,
                    'optC': q_optC_val,
                    'optD': q_optD_val,
                    'lesson': les_name
                })
                
        assert len(mcq_items) == 15, f'Lesson {les_name} has {len(mcq_items)} MCQs (expected 15)'
        assert len(tf_items) == 10, f'Lesson {les_name} has {len(tf_items)} TFs (expected 10)'
        
        # Write back in clean order: 15 MCQ (idx 1..15) followed by 10 TF (idx 16..25)
        ordered_items = [(item, False) for item in mcq_items] + [(item, True) for item in tf_items]
        
        for idx_in_les, (item, is_tf) in enumerate(ordered_items, start=1):
            target_r = rows[idx_in_les - 1]
            
            # Col A: نوع السؤال
            ws.cell(target_r, 1, '2 - صح/خطأ' if is_tf else '1 - اختيار من متعدد')
            # Col B: نص السؤال
            ws.cell(target_r, 2, item['text'])
            # Col C: الشرح / التفسير
            ws.cell(target_r, 3, item['exp'])
            # Col D: مستوى الصعوبة
            ws.cell(target_r, 4, DIFF_MAP[idx_in_les])
            # Col E: الإجابة الصحيحة
            ws.cell(target_r, 5, item['ans'])
            # Col F..I: الخيارات
            ws.cell(target_r, 6, item['optA'])
            ws.cell(target_r, 7, item['optB'])
            ws.cell(target_r, 8, item['optC'])
            ws.cell(target_r, 9, item['optD'])
            # Col J: الدرس
            ws.cell(target_r, 10, item['lesson'])

    # Clear previous validations and attach template DataValidations
    ws.data_validations.dataValidation.clear()

    dv_colA = DataValidation(type='list', formula1='L_1', allow_blank=True)
    dv_colA.error = 'اختر قيمة من القائمة.'
    dv_colA.errorTitle = 'قيمة غير صحيحة'
    dv_colA.showErrorMessage = True
    dv_colA.add(f'A2:A{max_r + 50}')
    ws.add_data_validation(dv_colA)

    dv_colD = DataValidation(type='list', formula1='L_2', allow_blank=True)
    dv_colD.error = 'اختر قيمة من القائمة.'
    dv_colD.errorTitle = 'قيمة غير صحيحة'
    dv_colD.showErrorMessage = True
    dv_colD.add(f'D2:D{max_r + 50}')
    ws.add_data_validation(dv_colD)

    dv_colE = DataValidation(type='list', formula1='L_3', allow_blank=True)
    dv_colE.error = 'اختر قيمة من القائمة.'
    dv_colE.errorTitle = 'قيمة غير صحيحة'
    dv_colE.showErrorMessage = True
    dv_colE.add(f'E2:E{max_r + 50}')
    ws.add_data_validation(dv_colE)

updated_path = r'C:\Users\MaximuM-Tech\Downloads\برنامج تدريب تخصص الضبع الاسود + بنوك_محدث_LMS.xlsx'
wb.save(updated_path)
print('Successfully saved to:', updated_path)

