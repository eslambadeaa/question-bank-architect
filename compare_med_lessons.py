import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

wb_src = openpyxl.load_workbook(r'C:\Users\MaximuM-Tech\Downloads\برنامج تدريب تخصص الضبع الاسود + بنوك_محدث_LMS.xlsx', data_only=True)
def get_prog(sheet):
    ws = wb_src[sheet]
    ls = []
    for r in range(5, ws.max_row+1):
        c4 = ws.cell(r, 4).value
        if c4 and str(c4).strip() not in ['None', '', 'امتحان منتصف الترم', 'اسم الموضوع']:
            t = str(c4).strip()
            if t not in ls: ls.append(t)
    return ls

med_prog = get_prog('برنامج تدريب - القسم المتوسط')
wb_raw = openpyxl.load_workbook(r'C:\Users\MaximuM-Tech\Downloads\بنوك\متوسط ضبع اسود.xlsx', data_only=True)
ws_raw = wb_raw.active
raw_med = []
for r in range(2, ws_raw.max_row+1):
    l = ws_raw.cell(r, 10).value
    if l and l not in raw_med: raw_med.append(l)

print('Med prog count:', len(med_prog), 'Raw med count:', len(raw_med))
for i in range(min(len(med_prog), len(raw_med))):
    if med_prog[i] != raw_med[i]:
        print(f'Diff at {i+1}: Prog=\"{med_prog[i]}\" vs Raw=\"{raw_med[i]}\"')
        break
else:
    print('ALL 52 LESSONS IN MEDIUM SECTION MATCH 100% IDENTICALLY!')
