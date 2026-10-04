import sys, io, openpyxl
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

path = r'C:\Users\MaximuM-Tech\Downloads\بنوك\برنامج تدريب تخصص الضبع الاسود_النسخة_الرسمية_المعتمدة.xlsx'
wb = openpyxl.load_workbook(path, data_only=True)

# 1. Prep
ws_prog = wb['برنامج تدريب - القسم الإعدادي']
prep_prog_lessons = []
for r in range(1, ws_prog.max_row + 1):
    c4 = ws_prog.cell(r, 4).value
    c5 = ws_prog.cell(r, 5).value
    if c4 and str(c4).strip() not in ['اسم الموضوع', 'None', ''] and (c5 == 2 or c5 == '2'):
        prep_prog_lessons.append(str(c4).strip())

ws_bank = wb['بنك القسم الإعدادي']
prep_bank_lessons = []
for r in range(2, ws_bank.max_row + 1):
    val = ws_bank.cell(r, 10).value
    if val and val not in prep_bank_lessons:
        prep_bank_lessons.append(val)

print('=== PREP ===')
print('Prog lessons count:', len(prep_prog_lessons))
print('Bank lessons count:', len(prep_bank_lessons))
for i, (p, b) in enumerate(zip(prep_prog_lessons, prep_bank_lessons)):
    if p != b:
        print(f'Mismatch {i+1}: Prog="{p}" vs Bank="{b}"')
if prep_prog_lessons == prep_bank_lessons:
    print('PREP: 100% MATCH!')

# 2. Medium
ws_prog_med = wb['برنامج تدريب - القسم المتوسط']
med_t1_prog = []
med_t2_prog = []
current_term = ''
for r in range(1, ws_prog_med.max_row + 1):
    c1 = ws_prog_med.cell(r, 1).value
    c4 = ws_prog_med.cell(r, 4).value
    c5 = ws_prog_med.cell(r, 5).value
    if c1 and 'الترم الأول' in str(c1): current_term = 't1'
    elif c1 and 'الترم الثاني' in str(c1): current_term = 't2'
    if c4 and str(c4).strip() not in ['اسم الموضوع', 'None', ''] and (c5 == 2 or c5 == '2'):
        if current_term == 't1': med_t1_prog.append(str(c4).strip())
        elif current_term == 't2': med_t2_prog.append(str(c4).strip())

med_t1_bank = []
ws_b_m1 = wb['بنك المتوسط - ترم أول']
for r in range(2, ws_b_m1.max_row + 1):
    val = ws_b_m1.cell(r, 10).value
    if val and val not in med_t1_bank: med_t1_bank.append(val)

med_t2_bank = []
ws_b_m2 = wb['بنك المتوسط - ترم ثاني']
for r in range(2, ws_b_m2.max_row + 1):
    val = ws_b_m2.cell(r, 10).value
    if val and val not in med_t2_bank: med_t2_bank.append(val)

print('\n=== MED T1 ===')
print('Prog count:', len(med_t1_prog), 'Bank count:', len(med_t1_bank))
for i, (p, b) in enumerate(zip(med_t1_prog, med_t1_bank)):
    if p != b:
        print(f'Mismatch {i+1}: Prog="{p}" vs Bank="{b}"')
if med_t1_prog == med_t1_bank: print('MED T1: 100% MATCH!')

print('\n=== MED T2 ===')
print('Prog count:', len(med_t2_prog), 'Bank count:', len(med_t2_bank))
for i, (p, b) in enumerate(zip(med_t2_prog, med_t2_bank)):
    if p != b:
        print(f'Mismatch {i+1}: Prog="{p}" vs Bank="{b}"')
if med_t2_prog == med_t2_bank: print('MED T2: 100% MATCH!')

# 3. Final
ws_prog_fin = wb['برنامج تدريب - القسم النهائي']
fin_t1_prog = []
fin_t2_prog = []
current_term = ''
for r in range(1, ws_prog_fin.max_row + 1):
    c1 = ws_prog_fin.cell(r, 1).value
    c4 = ws_prog_fin.cell(r, 4).value
    c5 = ws_prog_fin.cell(r, 5).value
    if c1 and 'الترم الأول' in str(c1): current_term = 't1'
    elif c1 and 'الترم الثاني' in str(c1): current_term = 't2'
    if c4 and str(c4).strip() not in ['اسم الموضوع', 'None', ''] and (c5 == 2 or c5 == '2'):
        if current_term == 't1': fin_t1_prog.append(str(c4).strip())
        elif current_term == 't2': fin_t2_prog.append(str(c4).strip())

fin_t1_bank = []
ws_b_f1 = wb['بنك النهائي - ترم أول']
for r in range(2, ws_b_f1.max_row + 1):
    val = ws_b_f1.cell(r, 10).value
    if val and val not in fin_t1_bank: fin_t1_bank.append(val)

fin_t2_bank = []
ws_b_f2 = wb['بنك النهائي - ترم ثاني']
for r in range(2, ws_b_f2.max_row + 1):
    val = ws_b_f2.cell(r, 10).value
    if val and val not in fin_t2_bank: fin_t2_bank.append(val)

print('\n=== FINAL T1 ===')
print('Prog count:', len(fin_t1_prog), 'Bank count:', len(fin_t1_bank))
for i, (p, b) in enumerate(zip(fin_t1_prog, fin_t1_bank)):
    if p != b:
        print(f'Mismatch {i+1}: Prog="{p}" vs Bank="{b}"')
if fin_t1_prog == fin_t1_bank: print('FINAL T1: 100% MATCH!')

print('\n=== FINAL T2 ===')
print('Prog count:', len(fin_t2_prog), 'Bank count:', len(fin_t2_bank))
for i, (p, b) in enumerate(zip(fin_t2_prog, fin_t2_bank)):
    if p != b:
        print(f'Mismatch {i+1}: Prog="{p}" vs Bank="{b}"')
if fin_t2_prog == fin_t2_bank: print('FINAL T2: 100% MATCH!')
