import sys
sys.stdout.reconfigure(encoding='utf-8')

# Define syllabus schedule with Theory and Practical split for each section

# PREPARATORY (28 periods, 7 days, 56 hours: 28h Theory + 28h Practical)
PREP_SCHEDULE = [
    # Day 1
    {"period": "ف1", "type": "theory", "theory": "التركيب العام والخواص الفنية والتكتيكية للمعدة", "prac": "-", "p_from": 1, "p_to": 3},
    {"period": "ف2", "type": "theory", "theory": "فكرة عامة والغرض من المعدة وأوضاع الإطلاق", "prac": "-", "p_from": 1, "p_to": 2},
    {"period": "ف3", "type": "prac", "theory": "-", "prac": "تدريب عملي على فحص وتمييز المكونات الرئيسية للمعدة وأوضاع حملها", "p_from": 1, "p_to": 3},
    {"period": "ف4", "type": "prac", "theory": "-", "prac": "تدريب عملي على تطبيق احتياطات الأمان وإجراءات التجهيز", "p_from": 3, "p_to": 4},
    # Day 2
    {"period": "ف1", "type": "theory", "theory": "أوضاع الإطلاق واحتياطات الأمان والتحذيرات الفنية", "prac": "-", "p_from": 3, "p_to": 4},
    {"period": "ف2", "type": "theory", "theory": "إجراءات التعامل مع عدم خروج الصاروخ بعد الضغطة الثانية", "prac": "-", "p_from": 4, "p_to": 5},
    {"period": "ف3", "type": "prac", "theory": "-", "prac": "تدريب عملي على إجراءات التعامل مع حالات عدم خروج الصاروخ من القاذف", "p_from": 4, "p_to": 5},
    {"period": "ف4", "type": "prac", "theory": "-", "prac": "تدريب عملي متقدم على مواقف الطوارئ والأعطال للرامي", "p_from": 4, "p_to": 5},
    # Day 3
    {"period": "ف1", "type": "theory", "theory": "الغرض من الصاروخ ومكوناته ورأس التوجيه الذاتي", "prac": "-", "p_from": 5, "p_to": 7},
    {"period": "ف2", "type": "theory", "theory": "الجهاز العدسي الجيروسكوبي وقرص التعديل والمقاومة الفوتوغرافية", "prac": "-", "p_from": 7, "p_to": 9},
    {"period": "ف3", "type": "prac", "theory": "-", "prac": "تدريب عملي على رأس التوجيه الذاتي والجهاز العدسي الجيروسكوبي", "p_from": 6, "p_to": 8},
    {"period": "ف4", "type": "prac", "theory": "-", "prac": "تدريب عملي على معايرة واختبار قرص التعديل والمقاومة الفوتوغرافية", "p_from": 8, "p_to": 9},
    # Day 4
    {"period": "ف1", "type": "theory", "theory": "ملفات رأس التوجيه الذاتي ونظام أخذ السرعة والتثبيت الكهربي", "prac": "-", "p_from": 10, "p_to": 12},
    {"period": "ف2", "type": "prac", "theory": "-", "prac": "تدريب عملي على منظومة التثبيت الكهربي وتشغيل الجيروسكوب", "p_from": 10, "p_to": 12},
    # Midterm here
    {"period": "ف3", "type": "theory", "theory": "نظام التوجيه والتتبع الآلي والدليل الآلي وقنطرة هيوستن", "prac": "-", "p_from": 13, "p_to": 15},
    {"period": "ف4", "type": "theory", "theory": "الجزء الخاص بالدفات ومستودع الغاز ومصدر التغذية المحمول", "prac": "-", "p_from": 15, "p_to": 18},
    # Day 5
    {"period": "ف1", "type": "theory", "theory": "الجزء القتالي وجهاز التفجير ورأس التدمير ووسائل الأمان", "prac": "-", "p_from": 18, "p_to": 20},
    {"period": "ف2", "type": "theory", "theory": "منظومة المحركات (الدافع والرئيسي) والأجنحة وحلقات الاتزان", "prac": "-", "p_from": 20, "p_to": 22},
    {"period": "ف3", "type": "theory", "theory": "القاذف ومواصفاته الفنية وحماية الرامي", "prac": "-", "p_from": 22, "p_to": 23},
    {"period": "ف4", "type": "theory", "theory": "مصدر التغذية الأرضي ومكوناته وفكرة عمله الكيميائية", "prac": "-", "p_from": 23, "p_to": 25},
    # Day 6 (8 hours practical)
    {"period": "ف1", "type": "prac", "theory": "-", "prac": "تدريب عملي على فحص القاذف ومصدر التغذية الأرضي", "p_from": 22, "p_to": 25},
    {"period": "ف2", "type": "prac", "theory": "-", "prac": "تدريب عملي على تركيب وتأمين مصدر التغذية الأرضي بالقاذف", "p_from": 23, "p_to": 25},
    {"period": "ف3", "type": "prac", "theory": "-", "prac": "تدريب عملي على مجموعة الإطلاق وضبط ذراع الأمان والتتك", "p_from": 25, "p_to": 27},
    {"period": "ف4", "type": "prac", "theory": "-", "prac": "تدريب عملي على تجميع عناصر المعدة واختبار التلامسات الكهربية", "p_from": 25, "p_to": 27},
    # Day 7
    {"period": "ف1", "type": "theory", "theory": "مجموعة الإطلاق والتعاون بين عناصر المعدة ووحدة التأخير", "prac": "-", "p_from": 25, "p_to": 29},
    {"period": "ف2", "type": "prac", "theory": "-", "prac": "تدريب عملي على خطوات التقاط الهدف وسماع الإشارة الصوتية", "p_from": 27, "p_to": 29},
    {"period": "ف3", "type": "prac", "theory": "-", "prac": "تدريب عملي على أخذ الضغطة الأولى والثانية وتسلسل الإطلاق", "p_from": 27, "p_to": 30},
    {"period": "ف4", "type": "prac", "theory": "-", "prac": "تدريب عملي شامل على سيناريو اشتباك متكامل من الرصد حتى التدمير", "p_from": 29, "p_to": 31},
]

th_count = sum(1 for x in PREP_SCHEDULE if x['type'] == 'theory')
pr_count = sum(1 for x in PREP_SCHEDULE if x['type'] == 'prac')
print(f"PREP: Theory periods = {th_count} ({th_count*2}h), Practical periods = {pr_count} ({pr_count*2}h), Total = {len(PREP_SCHEDULE)} periods ({len(PREP_SCHEDULE)*2}h)")
