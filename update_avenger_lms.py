import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
import sys

sys.stdout.reconfigure(encoding='utf-8')

file_path = r'C:\Users\MaximuM-Tech\Downloads\بنوك\برنامج_تدريب_وبنك_أسئلة_معدة_الإطلاق_أفنجر_القسم_المتوسط.xlsx'
wb = openpyxl.load_workbook(file_path, data_only=False)
print('Avenger Master loaded. Sheets:', wb.sheetnames)

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

bank_sheets = ['بنك المتوسط - ترم أول', 'بنك المتوسط - ترم ثاني']

for bs in bank_sheets:
    ws = wb[bs]
    max_r = ws.max_row
    print(f'\n--- Updating {bs} ({max_r - 1} questions) ---')
    
    for r in range(2, max_r + 1):
        idx_in_lesson = (r - 2) % 25 + 1
        is_tf = (idx_in_lesson > 15)
        curr_ans = str(ws.cell(r, 5).value).strip() if ws.cell(r, 5).value is not None else ''
        
        # 1. Col A: نوع السؤال
        if is_tf:
            ws.cell(r, 1, '2 - صح/خطأ')
        else:
            ws.cell(r, 1, '1 - اختيار من متعدد')
            
        # 2. Col D: مستوى الصعوبة
        ws.cell(r, 4, DIFF_MAP[idx_in_lesson])
        
        # 3. Col E: الإجابة الصحيحة
        if is_tf:
            if curr_ans in ['الخيار أ (A)', 'A', 'أ', 'صواب', 'True', 'true', 'صح']:
                ws.cell(r, 5, 'True')
            elif curr_ans in ['الخيار ب (B)', 'B', 'ب', 'خطأ', 'False', 'false', 'خطا']:
                ws.cell(r, 5, 'False')
            else:
                raise ValueError(f'Unknown TF ans in {bs} row {r}: {curr_ans}')
            
            # Standardize options for TF
            ws.cell(r, 6, 'True')
            ws.cell(r, 7, 'False')
            ws.cell(r, 8, None)
            ws.cell(r, 9, None)
        else:
            if curr_ans in ['الخيار أ (A)', 'A', 'أ']:
                ws.cell(r, 5, 'A')
            elif curr_ans in ['الخيار ب (B)', 'B', 'ب']:
                ws.cell(r, 5, 'B')
            elif curr_ans in ['الخيار ج (C)', 'C', 'ج']:
                ws.cell(r, 5, 'C')
            elif curr_ans in ['الخيار د (D)', 'D', 'د']:
                ws.cell(r, 5, 'D')
            else:
                raise ValueError(f'Unknown MCQ ans in {bs} row {r}: {curr_ans}')

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

# Save directly to file (or retry if locked)
try:
    wb.save(file_path)
    print('\nSuccessfully saved to:', file_path)
except PermissionError:
    alt_path = r'C:\Users\MaximuM-Tech\Downloads\بنوك\برنامج_تدريب_وبنك_أسئلة_معدة_الإطلاق_أفنجر_القسم_المتوسط_محدث_LMS.xlsx'
    wb.save(alt_path)
    print('\nFile locked, saved to alternative path:', alt_path)
