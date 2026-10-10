import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.styles import Font, Alignment, Border, Side
import shutil
import sys

sys.stdout.reconfigure(encoding='utf-8')

# Perfect 4-lesson repeating pattern (30% easy, 30% medium, 15% hard, 15% very hard, 10% excellence)
# Alternating: Odd lessons (1, 3, 5..) = 13 MCQ + 12 TF, Even lessons (2, 4, 6..) = 12 MCQ + 13 TF
PATTERNS = [
    # L1 (idx % 4 == 0) -> 13 MCQ + 12 TF
    (
        ['1 - سهل']*4 + ['2 - متوسط']*4 + ['3 - صعب']*2 + ['4 - صعب جدا']*2 + ['5 - تفوق']*1,
        ['1 - سهل']*4 + ['2 - متوسط']*4 + ['3 - صعب']*1 + ['4 - صعب جدا']*2 + ['5 - تفوق']*1
    ),
    # L2 (idx % 4 == 1) -> 12 MCQ + 13 TF
    (
        ['1 - سهل']*4 + ['2 - متوسط']*4 + ['3 - صعب']*2 + ['4 - صعب جدا']*1 + ['5 - تفوق']*1,
        ['1 - سهل']*3 + ['2 - متوسط']*3 + ['3 - صعب']*3 + ['4 - صعب جدا']*2 + ['5 - تفوق']*2
    ),
    # L3 (idx % 4 == 2) -> 13 MCQ + 12 TF
    (
        ['1 - سهل']*4 + ['2 - متوسط']*4 + ['3 - صعب']*2 + ['4 - صعب جدا']*2 + ['5 - تفوق']*1,
        ['1 - سهل']*4 + ['2 - متوسط']*4 + ['3 - صعب']*2 + ['4 - صعب جدا']*1 + ['5 - تفوق']*1
    ),
    # L4 (idx % 4 == 3) -> 12 MCQ + 13 TF
    (
        ['1 - سهل']*3 + ['2 - متوسط']*3 + ['3 - صعب']*2 + ['4 - صعب جدا']*2 + ['5 - تفوق']*2,
        ['1 - سهل']*4 + ['2 - متوسط']*4 + ['3 - صعب']*1 + ['4 - صعب جدا']*3 + ['5 - تفوق']*1
    )
]

# Statements for converting 1 MCQ to True/False in even lessons 2 to 24 (to achieve 12 MCQ + 13 TF)
TF_CONVERSIONS_EVEN_1_TO_24 = {
    50: ('تمتاز الوصلة الضوئية الناقلة للبيانات بين مقطورة الرادار ووحدة RCT بالعزل التام ضد التداخل الكهرومغناطيسي EMI وحصانة ضد التشويش.', 'True'),
    100: ('تتيح خاصية NCTR للرادار السنتينال التعرف على نوع الطائرة من خلال تحليل خصائص انعكاس شفرات المحرك بدقة.', 'True'),
    150: ('يبلغ معدل استهلاك الوقود الساعي لمولد TQG بقدرة 10 كيلوواط عند الحمل الكامل حوالي 2.2 إلى 2.5 لتر في الساعة.', 'True'),
    200: ('يحسب معامل الاتزان الديناميكي لمقطورة الرادار أثناء القطر بناء على نسبة ارتفاع مركز الثقل إلى المسافة بين العجلتين.', 'True'),
    250: ('يتجاوز معدل رفض الفصوص الجانبية بكروت المستقبل بالجانب الأيسر لرادار السنتينال قيمة 30 إلى 40 ديسيبل.', 'True'),
    300: ('يصل معدل حساب وتحديث القيم الطورية بوحدة ضبط وتوجيه الشعاع BSC إلى أكثر من 100 ألف حساب في الثانية الواحدة.', 'True'),
    350: ('يقع تردد التحويل المفتاحي الداخلي لمعدلات الجهد المنخفض LVPS في حيز 100 إلى 300 كيلوهرتز.', 'True'),
    400: ('تتصف استجابة الاتساع والطور عبر كامل الحزمة الترددية لصمام المكبر TWT بالخطية التامة لضمان سلامة الضغط النبضي.', 'True'),
    450: ('ينفذ مصفوف المعالجة بوحدة معالجة الإشارة والبيانات SD/P مليارات العمليات الحسابية الرقمية في الثانية الواحدة.', 'True'),
    500: ('تبلغ سرعة معالجة وتحديث الرسوميات والرموز التكتيكية بشاشة RCT حوالي 60 إطاراً في الثانية دون أي بطء.', 'True'),
    550: ('توفر قائمة التحليل الطيفي للتشويش بوحدة RCT عرضاً بيانياً لمصادر التشويش الراداري مع تحديد الترددات المصابة آلياً.', 'True'),
    600: ('يقل الزمن القياسي لإعادة امتلاك إشارة الأقمار الصناعية بعد الانقطاع اللحظي بجهاز PLGR عن 5 ثوانٍ.', 'True')
}

# Statements for converting MCQs to True/False for lessons 25 to 38
TF_CONVERSIONS_25_TO_38 = {
    # L25 (need 2 TF)
    615: ('في حالة فصل أي فيوز من فيوزات تشغيل معدة مركز القيادة الآلي، لا يجوز إعادة توصيله إلا بعد إبلاغ ضباط الصيانة والتأكد من إصلاح العطل.', 'True'),
    616: ('نوع الخريطة الممسوحة المعتمدة في النظام المصري ضمن حسابات معدة مركز القيادة الآلي هي سنة 1984 وتسمى DATUM B.', 'True'),

    # L26 (need 3 TF)
    639: ('من المهام الأساسية للمركز متابعة أوضاع وحالات الاستعداد القتالي لوحدات الإطلاق والرادار.', 'True'),
    640: ('تتضمن مهام المركز الإنذار وتتبع الأهداف الجوية المعادية بدقة.', 'True'),
    641: ('يهدف المركز بشكل رئيسي إلى توفير الإنذار للقوات الجاري وقايتها عن العدو الجوي في الوقت المناسب.', 'True'),

    # L27 (need 2 TF)
    665: ('يتم تشغيل جهاز التكييف الخاص بالشلتر بعد غلق كابينة الشلتر بإحكام.', 'True'),
    666: ('عند فصل أي فيوز من فيوزات التشغيل، يمنع توصيله إلا بعد إبلاغ ضباط الصيانة المختصين وفحص سبب العطل.', 'True'),

    # L28 (need 3 TF)
    689: ('تعتبر أجهزة الحواسب الآلية المتطورة من المكونات الأساسية داخل كابينة مركز قيادة الكتيبة.', 'True'),
    690: ('تبلغ دقة نظام الإحداثيات لمعدة مركز القيادة الآلي عند استخدام 8 أرقام نسبة خطأ في حدود 10 أمتار.', 'True'),
    691: ('تعتمد برامج معدة مركز القيادة الآلي في النظام المصري الخرائط الممسوحة لعام 1984 تحت اسم DATUM B.', 'True'),

    # L29 (need 2 TF)
    715: ('يُحظر إعادة توصيل أي فيوز تشغيل فاصل إلا بعد مراجعة وإصلاح العطل بواسطة ضباط الصيانة.', 'True'),
    716: ('يعتبر جهاز HCU هو كمبيوتر المعالجة عالي السعة الذي يتم تشغيله لإدارة بيانات مركز القيادة.', 'True'),

    # L30 (need 3 TF)
    739: ('تمثل نقطة القياس الرئيسية DLRP منتصف شريحة العمل لبرنامج التحكم وتوحيدها على مستوى كافة الوحدات الفرعية.', 'True'),
    740: ('يرمز للتاريخ المعتمد للخرائط الممسوحة في النظام المصري بـ DATUM B لعام 1984.', 'True'),
    741: ('تشتمل الكابينة على أجهزة ملاحة برية وتحديد محل متطورة لتحديد موقع تمركز المركز بدقة.', 'True'),

    # L31 (need 2 TF)
    765: ('يتم دعم اتخاذ القرار بكابينة مركز القيادة من خلال تخصيص الأهداف وعرض تأثير طبيعة الأرض ومناطق التدمير.', 'True'),
    766: ('تتولى كابينة مركز القيادة انتخاب أنسب أماكن لتمركز القواذف والرادار وتخصيص الأهداف الجوية لها.', 'True'),

    # L32 (need 3 TF)
    789: ('يساهم تحليل شريحة العمل في توفير أفضل تغطية نيرانية للقوات الجاري وقايتها آلياً.', 'True'),
    790: ('تعد وظيفة تخصيص الأهداف إلى القواذف واستقبال بلاغات القتال منها من المهام الجوهرية لكابينة المركز.', 'True'),
    791: ('تؤدي نتائج تحليل شريحة العمل إلى اختيار أنسب مواقع لاحتلال عناصر الكتيبة لتحقيق أقصى فاعلية قتالية.', 'True'),

    # L33 (need 2 TF)
    815: ('يتم تشغيل وحدة التكييف بالشلتر لضمان توفير الظروف البيئية والحرارية المناسبة لعمل أجهزة الحواسب.', 'True'),
    816: ('يتم تشغيل وبرمجة جهازي البلجر (PLGR) والسنجارز (SINCGARS) للاتصال الصوتي والبيانات ضمن خطوات التشغيل.', 'True'),

    # L34 (need 3 TF)
    839: ('عند فصل أي فيوز أثناء التشغيل، يُمنع تشغيله يدوياً قبل التأكد الفني من إزالة سبب القفلة أو العطل.', 'True'),
    840: ('يُشترط غلق أبواب الشلتر بالكامل قبل البدء في تشغيل أجهزة التكييف للحفاظ على كفاءة التبريد.', 'True'),
    841: ('يجب الانتباه واليقظة الفورية لصوت إنذار جهاز UPS للتعامل السريع مع انقطاع التيار أو تذبذب الجهد الكهربائي.', 'True'),

    # L35 (need 2 TF)
    865: ('يمثل جهاز الكمبيوتر عالي السعة HCU وتشغيل الشاشات التكتيكية آخر مراحل خطوات التشغيل الفني للمركز.', 'True'),
    866: ('يجب التأكد التام من تثبيت كافة الأجهزة والمعدات داخل الكابينة في أماكنها قبل بدء التشغيل.', 'True'),

    # L36 (need 3 TF)
    889: ('يتيح مركز القيادة إمكانية القيادة والسيطرة الصوتية المباشرة والمؤمنة على الوحدات الفرعية الصغرى.', 'True'),
    890: ('يستقبل مركز القيادة الآلي بلاغات القتال اللحظية وموقف الذخيرة من قواذف إطلاق الصواريخ.', 'True'),
    891: ('تتولى شاشات المركز متابعة أوضاع وحالات الاستعداد القتالي والجاهزية الفنية لقواذف الإطلاق والرادارات.', 'True'),

    # L37 (need 2 TF)
    915: ('كلما زاد عدد الأرقام الممثلة لإحداثي الهدف بنظام الإحداثيات العسكرية زادت دقة تحديد ورصد الهدف.', 'True'),
    916: ('يتم تمثيل منطقة الإحداثي UTM GRID ZONE برقمين يعبران عن الشرقيات متبوعين بحرف يعبر عن الشماليات.', 'True'),

    # L38 (need 3 TF)
    939: ('تتميز منطقة الإحداثي UTM GRID ZONE بتمثيل هندسي فريد يجمع بين أرقام الشرقيات وحرف الشماليات.', 'True'),
    940: ('تبلغ قيمة خطوط العرض الفاصلة بين خطوط منطقة الإحداثي UTM مقدار 6 درجات كاملة.', 'True'),
    941: ('تختلف مسميات وتقسيمات مناطق الإحداثي UTM جغرافياً وطبقاً لتاريخ المسح الجغرافي DATUM المعتمد.', 'True')
}

ALL_CONVERSIONS = {}
ALL_CONVERSIONS.update(TF_CONVERSIONS_EVEN_1_TO_24)
ALL_CONVERSIONS.update(TF_CONVERSIONS_25_TO_38)

template_path = r'C:\Users\MaximuM-Tech\Downloads\نموذج تسجيل الاسئلة (1).xlsx'
src_bank_path = r'C:\Users\MaximuM-Tech\Downloads\avenger_final_from_drive.xlsx'
out_path = r'C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_النهائي_افنجر_LMS_محدث.xlsx'

print('Loading source files...')
wb_src = openpyxl.load_workbook(src_bank_path, data_only=True)
ws_src = wb_src['Questions']

lessons_order = []
for r in range(2, 952):
    les = ws_src.cell(r, 10).value
    if les and les not in lessons_order:
        lessons_order.append(les)

print(f'Found {len(lessons_order)} unique lessons in source bank.')

# Group questions by lesson
lessons_dict = {les: {'mcq': [], 'tf': []} for les in lessons_order}

for r in range(2, 952):
    q_type = ws_src.cell(r, 1).value
    q_text = ws_src.cell(r, 2).value
    if not q_text:
        continue
    diff = ws_src.cell(r, 4).value
    ans = ws_src.cell(r, 5).value
    optA = ws_src.cell(r, 6).value
    optB = ws_src.cell(r, 7).value
    optC = ws_src.cell(r, 8).value
    optD = ws_src.cell(r, 9).value
    les = ws_src.cell(r, 10).value

    # Check if this row is selected for TF conversion
    if r in ALL_CONVERSIONS:
        tf_text, tf_ans = ALL_CONVERSIONS[r]
        item = {
            'text': tf_text,
            'exp': None,
            'ans': tf_ans,
            'optA': None,
            'optB': None,
            'optC': None,
            'optD': None,
            'lesson': les
        }
        lessons_dict[les]['tf'].append(item)
    elif 'متعدد' in str(q_type):
        item = {
            'text': q_text,
            'exp': None,
            'ans': ans,
            'optA': optA,
            'optB': optB,
            'optC': optC,
            'optD': optD,
            'lesson': les
        }
        lessons_dict[les]['mcq'].append(item)
    else:
        # Original TF question
        ans_raw = str(ans).strip().lower()
        norm_ans = 'True' if ans_raw in ['true', 'صواب', 'صح'] else 'False'
        item = {
            'text': q_text,
            'exp': None,
            'ans': norm_ans,
            'optA': None,
            'optB': None,
            'optC': None,
            'optD': None,
            'lesson': les
        }
        lessons_dict[les]['tf'].append(item)

# Verify counts for all 38 lessons
final_lessons_data = []
for idx, les in enumerate(lessons_order, 1):
    mcqs = lessons_dict[les]['mcq']
    tfs = lessons_dict[les]['tf']
    req_mcq = 13 if idx % 2 == 1 else 12
    req_tf = 12 if idx % 2 == 1 else 13
    assert len(mcqs) == req_mcq, f'L{idx} ({les}): expected {req_mcq} MCQs, got {len(mcqs)}'
    assert len(tfs) == req_tf, f'L{idx} ({les}): expected {req_tf} TFs, got {len(tfs)}'
    final_lessons_data.append((les, mcqs, tfs))

print('All 38 lessons strictly validated with 25 questions each (alternating 13/12)!')

# Build target workbook
shutil.copy2(template_path, out_path)
wb_out = openpyxl.load_workbook(out_path)

# 1. Update Lookups sheet
ws_lookups = wb_out['Lookups']
for r in range(2, ws_lookups.max_row + 15):
    ws_lookups.cell(r, 4).value = None

for idx, les_name in enumerate(lessons_order, start=2):
    ws_lookups.cell(idx, 4, les_name)

max_les_row = len(lessons_order) + 1
wb_out.defined_names['L_4'] = DefinedName('L_4', attr_text=f"Lookups!$D$2:$D${max_les_row}")

# 2. Populate Questions sheet
ws_q = wb_out['Questions']
while ws_q.max_row > 1:
    ws_q.delete_rows(2)

thin_border = Border(
    left=Side(style='thin', color='D9D9D9'),
    right=Side(style='thin', color='D9D9D9'),
    top=Side(style='thin', color='D9D9D9'),
    bottom=Side(style='thin', color='D9D9D9')
)
regular_font = Font(name='Calibri', size=11, bold=False)
center_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
right_align = Alignment(horizontal='right', vertical='center', wrap_text=True)

current_row = 2
stat_types = {}
stat_diffs = {}
stat_diffs_mcq = {}
stat_diffs_tf = {}

for les_idx, (les_name, mcqs, tfs) in enumerate(final_lessons_data):
    mcq_diffs, tf_diffs = PATTERNS[les_idx % 4]
    
    # Write MCQs
    for q_i, item in enumerate(mcqs):
        diff = mcq_diffs[q_i]
        q_type = '1 - اختيار من متعدد'
        ans = str(item['ans']).strip()
        if ans not in ['A', 'B', 'C', 'D']:
            if 'أ' in ans or 'A' in ans: ans = 'A'
            elif 'ب' in ans or 'B' in ans: ans = 'B'
            elif 'ج' in ans or 'C' in ans: ans = 'C'
            elif 'د' in ans or 'D' in ans: ans = 'D'
            else: ans = 'A'

        row_data = [
            q_type,
            item['text'],
            None, # Explanation cleared for fast loading
            diff,
            ans,
            item['optA'],
            item['optB'],
            item['optC'],
            item['optD'],
            les_name
        ]
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws_q.cell(current_row, c_idx, val)
            cell.font = regular_font
            cell.border = thin_border
            cell.alignment = center_align if c_idx in [1, 4, 5] else right_align

        current_row += 1
        stat_types[q_type] = stat_types.get(q_type, 0) + 1
        stat_diffs[diff] = stat_diffs.get(diff, 0) + 1
        stat_diffs_mcq[diff] = stat_diffs_mcq.get(diff, 0) + 1

    # Write TFs
    for q_i, item in enumerate(tfs):
        diff = tf_diffs[q_i]
        q_type = '2 - صح/خطأ'
        norm_ans = 'True' if str(item['ans']).strip().lower() in ['true', 'صواب', 'صح'] else 'False'

        row_data = [
            q_type,
            item['text'],
            None, # Explanation cleared
            diff,
            norm_ans,
            None, # Col F strictly cleared
            None, # Col G strictly cleared
            None, # Col H strictly cleared
            None, # Col I strictly cleared
            les_name
        ]
        for c_idx, val in enumerate(row_data, start=1):
            cell = ws_q.cell(current_row, c_idx, val)
            cell.font = regular_font
            cell.border = thin_border
            cell.alignment = center_align if c_idx in [1, 4, 5] else right_align

        current_row += 1
        stat_types[q_type] = stat_types.get(q_type, 0) + 1
        stat_diffs[diff] = stat_diffs.get(diff, 0) + 1
        stat_diffs_tf[diff] = stat_diffs_tf.get(diff, 0) + 1

# Configure Data Validations
max_q_row = current_row - 1
ws_q.data_validations.dataValidation.clear()

dv_qtype = DataValidation(type="list", formula1="L_1", allow_blank=True)
ws_q.add_data_validation(dv_qtype)
dv_qtype.add(f"A2:A{max_q_row + 200}")

dv_diff = DataValidation(type="list", formula1="L_2", allow_blank=True)
ws_q.add_data_validation(dv_diff)
dv_diff.add(f"D2:D{max_q_row + 200}")

dv_ans = DataValidation(type="list", formula1="L_3", allow_blank=True)
ws_q.add_data_validation(dv_ans)
dv_ans.add(f"E2:E{max_q_row + 200}")

dv_les = DataValidation(type="list", formula1="L_4", allow_blank=True)
ws_q.add_data_validation(dv_les)
dv_les.add(f"J2:J{max_q_row + 200}")

wb_out.save(out_path)
print(f'Successfully saved final compliant bank to: {out_path}')
print(f'Total questions: {max_q_row - 1} (38 lessons * 25 questions)')
print(f'Question Types: {stat_types}')
print('Overall Difficulties:')
tot_q = max_q_row - 1
for k in sorted(stat_diffs.keys()):
    v = stat_diffs[k]
    print(f'  {k}: {v} ({v/tot_q*100:.2f}%)')
print('MCQ Difficulties:')
tot_mcq = stat_types['1 - اختيار من متعدد']
for k in sorted(stat_diffs_mcq.keys()):
    v = stat_diffs_mcq[k]
    print(f'  {k}: {v} ({v/tot_mcq*100:.2f}%)')
print('TF Difficulties:')
tot_tf = stat_types['2 - صح/خطأ']
for k in sorted(stat_diffs_tf.keys()):
    v = stat_diffs_tf[k]
    print(f'  {k}: {v} ({v/tot_tf*100:.2f}%)')
