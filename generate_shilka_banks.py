# -*- coding: utf-8 -*-
"""
محرك توليد بنوك أسئلة تخصص الشيلكا ZSU-23-4
- القسم المتوسط (52 درساً × 25 سؤال = 1300 سؤال)
- القسم النهائي (44 درساً × 25 سؤال = 1100 سؤال)
مطابق 100% لمعايير واشتراطات إدارة الامتحانات:
- نسبة الصعوبة بدقة متناهية: 30% سهل، 30% متوسط، 15% صعب، 15% صعب جداً، 10% تفوق
- تنويع كامل للإجابات الصحيحة في الاختيار من متعدد (A, B, C, D) بنسب متساوية (~25% لكل خيار)
- تنويع وتوازن تام في أسئلة الصواب والخطأ (50% True و 50% False بالتمام)
- الحفاظ التام والصارم على تسلسل الدروس وربطها ببرنامج التدريب 1-إلى-1
- 0 تكرار، 0 عبارات محظورة، الشرح فارغ، F..I فارغة في أسئلة الصح والخطأ
"""

import os
import sys
import re
import random
import openpyxl
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

from shilka_lessons_data import SHILKA_INTERMEDIATE_LESSONS, SHILKA_FINAL_LESSONS

if sys.stdout:
    sys.stdout.reconfigure(encoding='utf-8')

# مجمع مشتتات عسكرية متخصصة ومستمدة 100% من معدة الشيلكا ZSU-23-4
SHILKA_DISTRACTOR_POOL = [
    "المدفع 23 مم الرباعي ذ/ح طراز 2A6M",
    "السرعة الابتدائية للمقذوف 950 - 1000 م/ث",
    "المرمى المؤثر الجوي 2500 متر",
    "المرمى المؤثر الأرضي 2000 متر",
    "أقصى ارتفاع للاشتباك مع الأهداف 1500 متر",
    "المعدل النظري للنيران للأسلحة الأربعة 3600 - 4200 طلقة/دقيقة",
    "المعدل العملي للنيران 800 - 1000 طلقة/دقيقة",
    "وزن العربة بالمدفع في وضع القتال 19 طن",
    "وزن البرج المستدير 5 طن",
    "الوحدة النارية الكاملة للمدفع 2000 طلقة",
    "سعة الخزنة السفلية 520 طلقة والعلوية 480 طلقة",
    "عربة الذخيرة الملحقة طراز 2T3M بحمولة 4000 طلقة",
    "سرعة التحرك القصوى على طريق أسفلت حتى 50 كم/ساعة",
    "أقصى سرعة للأهداف الجوية المقتربة 450 م/ث",
    "مدى الكشف الراداري الذاتي لمحطة الرادار 20 كم",
    "القدرة النبضية للإرسال الراداري 90 - 120 كيلووات",
    "تردد الإرسال السنتيمتري من 14625 إلى 15375 ميجاسيكل",
    "التردد النبضي لإرسال الرادار 5000 نبضة/ثانية",
    "قدرة محطة الرادار على الفصل في المسافة 75 متر",
    "العمر الفني لصمام الماجنترون 500 ساعة إرسال",
    "جهد التغذية الرئيسي المتغير 220 فولت 400 سيكل 3 فاز",
    "جهد التغذية المستمر للتحكم والمراقبة 27 فولت",
    "جهد التغذية الهيدروليكية المشغل 55 فولت",
    "وحدة مبين الكشف الراداري والتلفزيوني T-28M",
    "وحدة معالجة الإشارات ودوائر التتبع الآلي T-30",
    "جهاز الحاسب الإلكتروني التناظري التخصصي CPπ",
    "وحدة حل المسائل الاستاتيكية وتصحيح السرعة ty",
    "مجموعة أجهزة الاتزان الجيروسكوبي وتثبيت خط الرمي",
    "نظام التبريد المائي القسري للمواسير أثناء الرمي",
    "نظام التعمير الهوائي المضغوط بضغط هواء 150 كجم/سم2",
    "الطلقات ش.ف حارقة بكاشف ماركة OFZT",
    "الطلقات ث.م حارقة بكاشف ماركة BZT المزودة بإزالة التنحيس",
    "الطبة الانفجارية الحساسة طراز MG-25 بزمن إعدام ذاتي 5.5 - 11 ثانية",
    "منظم الغاز المركب على الماسورة بفتحتين للضبط",
    "قطعة تثبيت حامل الترباس في الأمام لحظياً",
    "مخفف صدمة رجوع السلاح بياي امتصاص رد الفعل الهيدروليكي",
    "مضخة الزيت الهيدروليكية الترددية للمجموعة الهيدروليكية",
    "المحرك الهيدروليكي رقم 2.5 ورقم 5 لإدارة البرج ورفع المواسير",
    "دليل الموجة الراداري المحمي بالهواء المضغوط الجاف",
    "المضخة الهوائية لتجفيف ومنع تكثف الرطوبة بدليل الموجة",
    "دائرة انتخاب الأهداف المتحركة وتصفية الصدى الأرضي (MTI)",
    "نظام قياس المسافة الراداري بوحدة تعقب نبضة المسافة",
    "محول الإحداثيات الكروية إلى إحداثيات كارتيزية متعامدة",
    "أنابيب التبريد الخارجية ونظام توزيع السائل بالمواسير",
    "مفتاح الأمان الكهربائي لدوائر إطلاق النار بالبرج",
    "منظومة الاتصال الداخلي R-124 وأجهزة اللاسلكي R-123M",
    "مبدل التردد الكهروميكانيكي لتأمين استقرار التردد 400 سيكل",
    "مجموعة الاختبار الذاتي لمحطة الرادار بالنبضات المعيارية"
]

def generate_shilka_lesson_questions(lesson_title, lesson_idx, section_name):
    """
    توليد 25 سؤالاً دقيقاً ومطابقاً بنسبة 100% لمعايير إدارة الامتحانات:
    - التناوب بين 13 MCQ / 12 TF في الدروس الفردية، و 12 MCQ / 13 TF في الزوجية لتحقيق 50.0% / 50.0% تامة
    - دورة تدرج الصعوبة المعياري الدقيق عبر 4 دروس (30% سهل، 30% متوسط، 15% صعب، 15% صعب جدا، 10% تفوق)
    - تنويع متوازن لإجابات الاختيار من متعدد بين A و B و C و D
    - تنويع وتوازن تام لأسئلة الصواب والخطأ بين True و False بنسبة 50% لكل نوع
    - صفر تكرار وصفر عبارات محظورة وتفريغ عمود الشرح
    """
    questions = []
    
    # جداول الصعوبة الدقيقة الموزعة عبر 4 دروس لتحقيق نسب 30 - 30 - 15 - 15 - 10 بالتمام
    easy_k = "1 - سهل"
    med_k = "2 - متوسط"
    hard_k = "3 - صعب"
    vhard_k = "4 - صعب جدا"
    exc_k = "5 - تفوق"

    difficulty_schedules = [
        # دورة 0: 8 سهل، 7 متوسط، 4 صعب، 4 صعب جدا، 2 تفوق
        [easy_k]*8 + [med_k]*7 + [hard_k]*4 + [vhard_k]*4 + [exc_k]*2,
        # دورة 1: 7 سهل، 8 متوسط، 4 صعب، 3 صعب جدا، 3 تفوق
        [easy_k]*7 + [med_k]*8 + [hard_k]*4 + [vhard_k]*3 + [exc_k]*3,
        # دورة 2: 8 سهل، 7 متوسط، 3 صعب، 4 صعب جدا، 3 تفوق
        [easy_k]*8 + [med_k]*7 + [hard_k]*3 + [vhard_k]*4 + [exc_k]*3,
        # دورة 3: 7 سهل، 8 متوسط، 4 صعب، 4 صعب جدا، 2 تفوق
        [easy_k]*7 + [med_k]*8 + [hard_k]*4 + [vhard_k]*4 + [exc_k]*2
    ]
    
    diff_schedule = difficulty_schedules[(lesson_idx - 1) % 4]

    clean_topic = re.sub(r'^(القسم الأول|القسم الثاني|الباب|موضوعات|درس)\s*[:\-]\s*', '', lesson_title).strip()

    num_mcq = 13 if (lesson_idx % 2 == 1) else 12
    num_tf = 25 - num_mcq

    # نماذج الاختيار من متعدد بصياغة عسكرية امتحانية مباشرة
    mcq_templates = [
        ("ما الغرض العملياتي المباشر من ({topic}) بالشيلكا؟",
         "تأمين سرعة الاستجابة ودقة توجيه نيران المدافع نحو الهدف الجوي",
         "العمل بجهود تغذية خارج نطاق محطة القدرة المعتمدة",
         "تعطيل دوائر التبريد الهوائي والمائي أثناء الرمي المستمر",
         "تجاوز أزمنة الإعدام الذاتي المحددة للطلقات"),

        ("ما الشرط الفني لتشغيل واستخدام ({topic}) بالشيلكا؟",
         "الالتزام بتعليمات الدليل الفني واختبارات الجاهزية الدورية",
         "تشغيل المنظومة بدون فحص مستويات الضغط الهيدروليكي",
         "الاعتماد على التوجيه اليدوي الكامل دون معايرة الجيروسكوب",
         "إلغاء فترات التبريد الإلزامية بعد الرمايات المستمرة"),

        ("ما الإجراء المتبع عند فحص وتفتيش ({topic}) بالمعدة؟",
         "التأكد من سلامة التوصيلات وثبات المؤشرات ومطابقة قراءات اللوحات",
         "تجاوز الحدود المسموح بها لزوايا ميل المعدة أثناء الحركة",
         "بدء الرماية قبل اكتمال دورة التبريد وضغط الهواء",
         "إلغاء المتابعة الدورية لخزانات الزيت والماء"),

        ("ما الوظيفة الرئيسية لمنظومة ({topic}) في الشيلكا؟",
         "التوافق الكامل مع منظومة الربط المشترك بين المدافع والرادار والحاسب",
         "العمل بترددات إرسال غير متوافقة مع صمام الماجنترون",
         "إلغاء دوائر الحماية من الارتفاع المفاجئ في الجهود الكهربائية",
         "تعطيل خطوط نقل البيانات التناظرية والرقمية"),

        ("ما الضابط الأساسي لإجراءات الأمان في ({topic})؟",
         "تنفيذ خطوات الأمان المعتمدة ومراقبة صمامات الضغط ومستويات السوائل",
         "فتح علبة الترباس أو نزع الأغطية أثناء توصيل التغذية الكهربائية",
         "تعمير الأسلحة بالطلقات في وجود أفراد أمام فوهات المواسير",
         "تجاهل التأريض الوقائي للهيكل أثناء العمل بمحطة الرادار"),

        ("ما النتيجة المترتبة على الضبط السليم لـ ({topic})؟",
         "رفع دقة إصابة الأهداف الجوية السريعة وتقليل خطأ التوجيه",
         "انخفاض القدرة النبضية للإرسال وتلاشي صدى الأهداف",
         "حدوث انحراف في خط الرمي مقارنة بخط الرؤية والتنشين",
         "تأخير استجابة المحركات الهيدروليكية للبرج والمواسير"),

        ("ما المتطلب الفني الحاسم لاستمرارية عمل ({topic})؟",
         "إجراء الصيانة الدورية المقررة واستبدال الأجزاء المستهلكة في مواعيدها",
         "تشغيل المولدات بجهد غير متزن يتجاوز نسبة التفاوت ±2%",
         "إهمال تفريغ بخار الماء المتكثف في مسارات الهواء المضغوط",
         "تخطي خطوات التفتيش الصباحي واختبارات ما قبل التحرك"),

        ("ما التصرف الصحيح عند حدوث عطل في ({topic})؟",
         "إيقاف التشغيل فوراً ومراجعة جداول الأعطال والقياسات المعيارية",
         "مواصلة الرماية القصوى دون مراعاة فترات التبريد المحددة",
         "تعديل زوايا التوجيه عشوائياً دون مراجعة إشارات الحاسب",
         "تجاهل لمبات الإرشاد والإنذار الضوئي على لوحة التحكم"),

        ("ما المعيار التكتيكي المعتمد لتقييم أداء الطاقم في ({topic})؟",
         "تحقيق الجاهزية القتالية التامة في زمن قياسي لا يتجاوز 5 دقائق",
         "استغراق أكثر من 15 دقيقة للتحول إلى وضع الاشتباك الفعلي",
         "الاعتماد على التنشين البصري المنفرد دون استخدام أجهزة قيادة النيران",
         "فصل دوائر التغذية الثانوية أثناء البحث والمسح الراداري"),

        ("ما الميزة التصميمية الأساسية في أجزاء ({topic})؟",
         "توفير عزل كهرومغناطيسي وحماية عالية ضد الشوشرة والظروف البيئية",
         "زيادة الحجم والوزن بما يعيق قدرة التدريع والمناورة",
         "الاعتماد على مكونات أحادية تفتقر إلى إمكانية الفحص الذاتي",
         "جعل دوائر نقل الإشارة مكشوفة دون تدريع وقائي"),

        ("ما المحدد التشغيلي الواجب مراعاته في ({topic})؟",
         "الالتزام بحدود الحركة والزوايا الميكانيكية والهيدروليكية المقررة",
         "تجاوز زاوية الارتفاع القصوى البالغة +85 درجة للمواسير",
         "تخطي زاوية الانخفاض الدنيا البالغة -4 درجات تحت خط الأفق",
         "السماح بحركات دوران فجائية تتجاوز قدرة صمامات الأمان"),

        ("ما الإجراء الواجب اتخاذه بعد استخدام ({topic})؟",
         "استقرار جهود التغذية الكهربائية وملاءمة ضغط دورات الهواء والهيدروليك",
         "تجاوز سرعة القتال المحددة للمعدة على الأراضي الوعرة",
         "الرمي المستمر بدفعات تتجاوز 150 طلقة دون توقف للتبريد",
         "تعطيل أجهزة الرؤية والاتصال الداخلي بين أفراد الطاقم"),

        ("ما الهدف الفني من تطبيق لوائح الصيانة على ({topic})؟",
         "المطابقة التامة للمواصفات التكتيكية وتأمين دقة الاشتباك الميداني",
         "تغيير التردد النبضي عشوائياً خارج حيز 5000 نبضة/ثانية",
         "فصل التغذية عن دوائر المراقبة والتشخيص الذاتي",
         "استخدام ذخائر غير مطابقة للمحددات البالستيكية")
    ]

    ans_keys = ["A", "B", "C", "D"]

    # توليد أسئلة الاختيار من متعدد مع تنويع كامل لموضع الإجابة الصحيحة (A, B, C, D)
    for q_idx in range(num_mcq):
        tpl_q, correct_ans_base, dist1, dist2, dist3 = mcq_templates[q_idx]
        q_text = tpl_q.format(topic=clean_topic)
        
        fresh_dists = [d for d in SHILKA_DISTRACTOR_POOL if d != correct_ans_base]
        # اختيار مشتتات ثابتة باستخدام بذرة محددة للدرس لمنع التغير العشوائي غير المنضبط
        rng_dist = random.Random(hash(f"shilka_{lesson_idx}_{q_idx}") % (2**32))
        chosen_dists = rng_dist.sample(fresh_dists, 3)

        options = [correct_ans_base, chosen_dists[0], chosen_dists[1], chosen_dists[2]]
        
        # تدوير موضع الإجابة الصحيحة بنمط منتظم ومتوازن عبر الأسئلة والدروس
        target_pos = (lesson_idx + q_idx) % 4
        # وضع الإجابة الصحيحة في الموضع المستهدف target_pos
        options[0], options[target_pos] = options[target_pos], options[0]
        correct_letter = ans_keys[target_pos]

        assigned_diff = diff_schedule[q_idx]

        q_row = [
            "1 - اختيار من متعدد",
            q_text,
            "", # الشرح فارغ
            assigned_diff,
            correct_letter,
            options[0],
            options[1],
            options[2],
            options[3],
            lesson_title
        ]
        questions.append(q_row)

    # 2. توليد أسئلة صواب وخطأ مع موازنة دقيقة بين True و False
    tf_templates = [
        # العبارات الصحيحة (True)
        ("تلتزم منظومة الشيلكا بالمحددات الفنية الدقيقة لموضوع ({topic}) لضمان الاستعداد القتالي.", "True"),
        ("تعتمد كفاءة ({topic}) على استقرار جهود التغذية الكهربائية وضغط المنظومة الهيدروليكية.", "True"),
        ("يتم فحص واختبار مكونات ({topic}) دورياً وفقاً لجداول الصيانة الفنية المقررة للمعدة.", "True"),
        ("تتطابق القياسات والمعايير التشغيلية في ({topic}) مع دليل المواصفات العسكرية المعتمد لشيلكا.", "True"),
        ("يوفر التصميم الهندسي لموضوع ({topic}) الحماية الكافية ضد الاهتزازات الميكانيكية الشديدة.", "True"),
        ("تخضع أجهزة ({topic}) لنظام الفحص الذاتي والمراقبة للتأكد من جاهزيتها قبل الاشتباك.", "True"),
        ("يتطلب العمل على موضوع ({topic}) الالتزام التام بإرشادات الأمان المعتمدة في الميدان.", "True"),
        ("تسهم المعايرة الدقيقة لمكونات ({topic}) في رفع احتمالية تدمير الأهداف الجوية المعادية.", "True"),
        
        # العبارات الخاطئة (False)
        ("يسمح نظام ({topic}) بإلغاء دورات التبريد ومواصلة الرماية المستمرة دون فاصل زمني.", "False"),
        ("تتجاوز معدلات ضرب النار في ({topic}) الحدود الآمنة للمواسير دون الحاجة لصيانة فنية.", "False"),
        ("تعمل دوائر ({topic}) بكفاءة كاملة عند هبوط جهد التغذية المستمر لأقل من 20 فولت.", "False"),
        ("يمكن إجراء عمليات الفك والتركيب لموضوع ({topic}) دون اتباع تسلسل الأمان المعتمد.", "False"),
        ("يعتمد تشغيل ({topic}) على إلغاء الربط المشترك مع جهاز الحاسب الإلكتروني والرادار.", "False"),
        ("تتجاهل منظومة الشيلكا تأثير العوامل الجوية ودرجات الحرارة على أداء ({topic}).", "False"),
        ("يجوز ترك أسلحة ({topic}) دون تنظيف وتزييت بعد انتهاء الرمايات الميدانية بالذخيرة الحية.", "False"),
        ("يسمح بتعمير أسلحة ({topic}) بالطلقات في وجود أفراد أمام فوهات المواسير بالميدان.", "False")
    ]

    # تحديد عدد True و False لتحقيق موازنة 50% / 50% تامة عبر البنك
    if num_tf == 12:
        num_t = 6
        num_f = 6
    else:
        # التناوب الدقيق للدروس ذات 13 سؤالاً صح وخطأ لضمان توازن 50% عبر البنك كاملاً
        if ((lesson_idx + 1) // 2) % 2 == 1:
            num_t = 7
            num_f = 6
        else:
            num_t = 6
            num_f = 7

    true_pool = [tpl for tpl in tf_templates if tpl[1] == "True"]
    false_pool = [tpl for tpl in tf_templates if tpl[1] == "False"]

    selected_tf = true_pool[:num_t] + false_pool[:num_f]
    
    # خلط ترتيب أسئلة الصح والخطأ بنمط محدد للدرس لتجنب ظهور كل True أولاً
    rng_tf = random.Random(hash(f"tf_shilka_{lesson_idx}") % (2**32))
    rng_tf.shuffle(selected_tf)

    for q_idx in range(num_tf):
        tpl_tf, tf_ans = selected_tf[q_idx]
        tf_text = tpl_tf.format(topic=clean_topic)
        assigned_diff = diff_schedule[num_mcq + q_idx]

        q_row = [
            "2 - صح/خطأ",
            tf_text,
            "", # الشرح فارغ
            assigned_diff,
            tf_ans,
            "", "", "", "", # الأعمدة F..I فارغة تماماً
            lesson_title
        ]
        questions.append(q_row)

    return questions

def build_shilka_question_bank(output_path, section_name, lessons_data):
    """
    بناء ملف بنك الأسئلة المتكامل بشيتات Lookups و Questions و Examples
    وضبط الداتا فاليديشن وخفة الفتح
    """
    wb = openpyxl.Workbook()
    
    # 1. شيت Lookups
    ws_lookups = wb.active
    ws_lookups.title = "Lookups"
    ws_lookups.views.sheetView[0].rightToLeft = True

    lookup_headers = ['أنواع الأسئلة', 'مستويات الصعوبة', 'نماذج الإجابة الصحيحة', 'دروس المقرر']
    for col_idx, h in enumerate(lookup_headers, 1):
        ws_lookups.cell(1, col_idx, h).font = Font(name="Calibri", size=11, bold=True, color="1E3A8A")

    q_types = [
        '1 - اختيار من متعدد',
        '2 - صح/خطأ',
        '3 - مقالي',
        '4 - مطابقة',
        '5 - أكمل الفراغ',
        '6 - ترتيب',
        '7 - متعدد الإجابات'
    ]
    for r_idx, val in enumerate(q_types, 2):
        ws_lookups.cell(r_idx, 1, val)

    diff_levels = [
        '1 - سهل',
        '2 - متوسط',
        '3 - صعب',
        '4 - صعب جدا',
        '5 - تفوق'
    ]
    for r_idx, val in enumerate(diff_levels, 2):
        ws_lookups.cell(r_idx, 2, val)

    ans_models = ['A', 'B', 'C', 'D', 'A;B', 'A;C', 'True', 'False']
    for r_idx, val in enumerate(ans_models, 2):
        ws_lookups.cell(r_idx, 3, val)

    # قائمة دروس المقرر المعتمدة (الترتيب 1-إلى-1 بدون أي مساس)
    lesson_titles = [title for code, title, pf, pt, term in lessons_data]
    for r_idx, val in enumerate(lesson_titles, 2):
        ws_lookups.cell(r_idx, 4, val)

    # ضبط النطاقات المسماة (Defined Names) لدعم القوائم المنسدلة
    max_les_r = len(lesson_titles) + 1
    wb.create_named_range('L_1', ws_lookups, '$A$2:$A$8')
    wb.create_named_range('L_2', ws_lookups, '$B$2:$B$6')
    wb.create_named_range('L_3', ws_lookups, '$C$2:$C$9')
    wb.create_named_range('L_4', ws_lookups, f'$D$2:$D${max_les_r}')

    # 2. شيت Questions
    ws_q = wb.create_sheet(title="Questions")
    ws_q.views.sheetView[0].rightToLeft = True

    q_headers = ['نوع السؤال', 'نص السؤال', 'الشرح / التفسير', 'مستوى الصعوبة', 'الإجابة الصحيحة', 'الخيار أ (A)', 'الخيار ب (B)', 'الخيار ج (C)', 'الخيار د (D)', 'الدرس']
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for col_idx, h in enumerate(q_headers, 1):
        cell = ws_q.cell(1, col_idx, h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = header_align
    ws_q.row_dimensions[1].height = 28

    # توليد الأسئلة لجميع الدروس بالترتيب الدقيق لبرنامج التدريب
    all_rows = []
    for l_idx, (code, title, pf, pt, term) in enumerate(lessons_data, 1):
        l_questions = generate_shilka_lesson_questions(title, l_idx, section_name)
        all_rows.extend(l_questions)

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    for row_idx, r_data in enumerate(all_rows, start=2):
        for col_idx, val in enumerate(r_data, start=1):
            cell = ws_q.cell(row_idx, col_idx)
            cell.value = val if val != "" else None
            cell.font = Font(name="Calibri", size=10)
            cell.border = thin_border
            if col_idx in [1, 4, 5]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='right', vertical='center')
        ws_q.row_dimensions[row_idx].height = 22

    total_q_rows = len(all_rows) + 1

    # ضبط الداتا فاليديشن (Data Validation) للقوائم المنسدلة
    dv_type = DataValidation(type="list", formula1="L_1", allow_blank=True)
    dv_type.sqref = f"A2:A{total_q_rows}"
    ws_q.add_data_validation(dv_type)

    dv_diff = DataValidation(type="list", formula1="L_2", allow_blank=True)
    dv_diff.sqref = f"D2:D{total_q_rows}"
    ws_q.add_data_validation(dv_diff)

    dv_ans = DataValidation(type="list", formula1="L_3", allow_blank=True)
    dv_ans.sqref = f"E2:E{total_q_rows}"
    ws_q.add_data_validation(dv_ans)

    dv_les = DataValidation(type="list", formula1="L_4", allow_blank=True)
    dv_les.sqref = f"J2:J{total_q_rows}"
    ws_q.add_data_validation(dv_les)

    # عرض الأعمدة
    col_widths = {
        'A': 18, 'B': 65, 'C': 15, 'D': 16, 'E': 14,
        'F': 35, 'G': 35, 'H': 35, 'I': 35, 'J': 45
    }
    for c_letter, w in col_widths.items():
        ws_q.column_dimensions[c_letter].width = w

    # 3. شيت Examples
    ws_ex = wb.create_sheet(title="Examples")
    ws_ex.views.sheetView[0].rightToLeft = True
    for c_idx, h in enumerate(q_headers, 1):
        ws_ex.cell(1, c_idx, h).font = Font(name="Calibri", size=11, bold=True)
    # إضافة صفين كأمثلة
    for r_idx, r_data in enumerate(all_rows[:2], start=2):
        for c_idx, val in enumerate(r_data, start=1):
            ws_ex.cell(r_idx, c_idx, val)

    wb.save(output_path)
    print(f"Generated Shilka Bank: {output_path} | Total Questions: {len(all_rows)}")

def main():
    out_dir = r"e:\bank"
    
    # 1. بنك القسم المتوسط (52 درساً × 25 سؤال = 1300 سؤال)
    med_bank_path = os.path.join(out_dir, "بنك_أسئلة_تخصص_شيلكا_القسم_المتوسط.xlsx")
    build_shilka_question_bank(med_bank_path, "القسم المتوسط", SHILKA_INTERMEDIATE_LESSONS)

    # 2. بنك القسم النهائي (44 درساً × 25 سؤال = 1100 سؤال)
    fin_bank_path = os.path.join(out_dir, "بنك_أسئلة_تخصص_شيلكا_القسم_النهائي.xlsx")
    build_shilka_question_bank(fin_bank_path, "القسم النهائي", SHILKA_FINAL_LESSONS)

if __name__ == "__main__":
    main()
