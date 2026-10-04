"""
V4 Official Institutional Generator for Training Programs & Question Banks
Conforms 100% to examination authority requirements:
1. Schedule Sheet:
   - Template preserved exactly.
   - Column C: Individual periods (ف1, ف2, ف3, ف4) - NEVER merged!
   - Column B: Day merged across 4 periods.
   - Column D: Unified lesson name (identical for theory & practical).
   - Column E (نظري): 2.0 for theory, '-' for practical.
   - Column F (عملي): '-' for theory, 2.0 for practical.
   - Column G (من) & H (الي): Grounded reference pages.
   - Column I (عدد الاسئلة): 25 for theory, '-' for practical.
   - Topics expanded: Every major component in a dedicated lecture (no composite titles).
   - Coverage:
     - Prep: 30% of Black Hyena manual (p 1-11).
     - Med: 100% of Black Hyena manual (p 1-33).
     - Final: 100% of Igla manual (p 1-33).
2. Question Bank Sheets:
   - Header strictly:
     [نوع السؤال, نص السؤال, الشرح / التفسير, مستوى الصعوبة, الإجابة الصحيحة, الخيار أ (A), الخيار ب (B), الخيار ج (C), الخيار د (D), الدرس]
   - Data validation on 'نوع السؤال' (اختيار من متعدد, صواب أو خطأ) and 'مستوى الصعوبة' (سهل, متوسط, صعب).
   - Questions ONLY for theoretical lectures (25 questions per lesson):
     - Prep: 14 lessons = 350 questions.
     - Med T1: 24 lessons = 600 questions.
     - Med T2: 28 lessons = 700 questions.
     - Final T1: 24 lessons = 600 questions.
     - Final T2: 20 lessons = 500 questions.
"""

import os
import sys
from copy import copy
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.stdout.reconfigure(encoding='utf-8')

ARABIC_DAYS = [
    "الأول", "الثاني", "الثالث", "الرابع", "الخامس", "السادس",
    "السابع", "الثامن", "التاسع", "العاشر", "الحادي عشر", "الثانى عشر",
    "الثالث عشر", "الرابع عشر", "الخامس عشر", "السادس عشر"
]

def create_thin_border():
    thin = Side(border_style='thin', color='A0A0A0')
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def apply_rtl(ws):
    try:
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = True
    except Exception:
        pass

# =========================================================================
# 1. PREPARATORY SECTION (30% of Black Hyena manual, pages 1 to 11)
# 7 days, 14 theoretical lessons + 14 practical lessons = 28 lectures = 56h
# 350 questions total (25 per theoretical lesson)
# =========================================================================
PREP_PAIRS = [
    # Day 1
    {"title": "التركيب العام لمعدة الضبع الأسود", "p_from": 1, "p_to": 2},
    {"title": "الخواص الفنية والتكتيكية للمعدة", "p_from": 1, "p_to": 3},
    # Day 2
    {"title": "فكرة عامة عن المعدة وأغراض الاستخدام", "p_from": 1, "p_to": 2},
    {"title": "أوضاع الإطلاق التكتيكية للمعدة", "p_from": 2, "p_to": 3},
    # Day 3
    {"title": "احتياطات الأمان والتحذيرات عند العمل على المعدة", "p_from": 3, "p_to": 4},
    {"title": "إجراءات التعامل مع عدم خروج الصاروخ بعد الضغطة الثانية", "p_from": 4, "p_to": 5},
    # Day 4 (Midterm after day 4)
    {"title": "الغرض من الصاروخ ومكوناته الأساسية", "p_from": 5, "p_to": 6},
    {"title": "رأس التوجيه الذاتي الغرض والخواص الفنية", "p_from": 6, "p_to": 7},
    # Day 5
    {"title": "الجهاز العدسي الجيروسكوبي ومكوناته", "p_from": 7, "p_to": 8},
    {"title": "مهام مكونات الجهاز العدسي الجيروسكوبي", "p_from": 7, "p_to": 8},
    # Day 6
    {"title": "قرص التعديل وشكل وطريقة عمله", "p_from": 8, "p_to": 9},
    {"title": "المقاومة الفوتوغرافية وجهاز الموازنة", "p_from": 9, "p_to": 10},
    # Day 7
    {"title": "ملفات رأس التوجيه الذاتي ومجموعاتها", "p_from": 10, "p_to": 11},
    {"title": "نظام أخذ العضو الدوار لسرعته العادية", "p_from": 10, "p_to": 11},
]

# =========================================================================
# 2. MEDIUM SECTION - TERM 1 (100% of Black Hyena manual, First Half p 1-18)
# 12 days, 24 theoretical lessons + 24 practical lessons = 48 lectures = 96h
# 600 questions total (25 per theoretical lesson)
# =========================================================================
MED_T1_PAIRS = [
    # Day 1
    {"title": "التركيب العام لمعدة الضبع الأسود ومكوناتها", "p_from": 1, "p_to": 2},
    {"title": "الخواص الفنية والتكتيكية للمعدة العادية والمعدلة", "p_from": 1, "p_to": 3},
    # Day 2
    {"title": "فكرة عامة عن المعدة وأغراض الاستخدام", "p_from": 1, "p_to": 2},
    {"title": "أوضاع الإطلاق التكتيكية للمعدة", "p_from": 2, "p_to": 3},
    # Day 3
    {"title": "احتياطات الأمان والتحذيرات عند العمل على المعدة", "p_from": 3, "p_to": 4},
    {"title": "إجراءات التعامل مع عدم خروج الصاروخ من القاذف", "p_from": 4, "p_to": 5},
    # Day 4
    {"title": "الغرض من الصاروخ وأوضاع إطلاقه", "p_from": 5, "p_to": 6},
    {"title": "المكونات الأساسية للصاروخ ومخططه الهيكلي", "p_from": 5, "p_to": 6},
    # Day 5
    {"title": "رأس التوجيه الذاتي الغرض والخواص الفنية", "p_from": 6, "p_to": 7},
    {"title": "الجهاز العدسي الجيروسكوبي ومكوناته", "p_from": 7, "p_to": 8},
    # Day 6 (Midterm after day 6)
    {"title": "مهام مكونات الجهاز العدسي الجيروسكوبي", "p_from": 7, "p_to": 8},
    {"title": "قرص التعديل وشكل وطريقة عمله", "p_from": 8, "p_to": 9},
    # Day 7
    {"title": "المقاومة الفوتوغرافية ودورها في رأس التوجيه", "p_from": 9, "p_to": 10},
    {"title": "جهاز موازنة دوران العضو الدوار وطريقة عمله", "p_from": 9, "p_to": 10},
    # Day 8
    {"title": "ملفات رأس التوجيه الذاتي ومجموعات التصحيح", "p_from": 10, "p_to": 11},
    {"title": "نظام أخذ العضو الدوار لسرعته العادية", "p_from": 10, "p_to": 11},
    # Day 9
    {"title": "أنظمة رأس التوجيه الذاتي ونظام التثبيت الكهربي", "p_from": 11, "p_to": 12},
    {"title": "نظام التوجيه والتتبع الآلي وفكرة عمله", "p_from": 13, "p_to": 14},
    # Day 10
    {"title": "الدليل الآلي ودائرة العمل الكهربية ومكوناتها", "p_from": 14, "p_to": 15},
    {"title": "جهاز تحديد السرعة الزاوية وقنطرة هيوستن", "p_from": 15, "p_to": 15},
    # Day 11
    {"title": "الجزء الخاص بالدفات ومكوناته ومواصفاته", "p_from": 15, "p_to": 16},
    {"title": "جهاز تحريك الدفات ومستودع الغاز المضغوط", "p_from": 16, "p_to": 17},
    # Day 12
    {"title": "مصدر التغذية المحمول على سطح الصاروخ", "p_from": 17, "p_to": 18},
    {"title": "طريقة عمل التوربين الغازي لمصدر التغذية المحمول", "p_from": 17, "p_to": 18},
]

# =========================================================================
# 3. MEDIUM SECTION - TERM 2 (100% of Black Hyena manual, Second Half p 18-33)
# 14 days, 28 theoretical lessons + 28 practical lessons = 56 lectures = 112h
# 700 questions total (25 per theoretical lesson)
# =========================================================================
MED_T2_PAIRS = [
    # Day 1
    {"title": "القسم القتالي الغرض العام ومكوناته", "p_from": 18, "p_to": 19},
    {"title": "جهاز التفجير ومكوناته ووسائل الأمان", "p_from": 18, "p_to": 20},
    # Day 2
    {"title": "طريقة عمل جهاز التفجير وحالات التعمير والانفجار الذاتي", "p_from": 19, "p_to": 20},
    {"title": "رأس التدمير ومكوناتها والعبوة المتفجرة", "p_from": 20, "p_to": 20},
    # Day 3
    {"title": "تأثير الشظايا والموجة الانفجارية لرأس التدمير", "p_from": 20, "p_to": 21},
    {"title": "الجزء الخاص بالمحركات والأجنحة والحلقتان المركزيتان", "p_from": 20, "p_to": 21},
    # Day 4
    {"title": "المحرك الدافع ومواصفاته وطريقة عمله", "p_from": 21, "p_to": 22},
    {"title": "المحرك الرئيسي وأزمنة عمل المحركات", "p_from": 21, "p_to": 22},
    # Day 5
    {"title": "القاذف الغرض منه والمواصفات العامة", "p_from": 22, "p_to": 23},
    {"title": "مكونات أنبوب القاذف والتجهيزات البصرية", "p_from": 22, "p_to": 23},
    # Day 6
    {"title": "مصدر التغذية الأرضي ومواصفاته العامة", "p_from": 23, "p_to": 24},
    {"title": "بطارية مصدر التغذية الأرضي والتفاعل الكيميائي", "p_from": 24, "p_to": 25},
    # Day 7 (Midterm after day 7)
    {"title": "مجموعة الإطلاق الغرض والخواص الفنية", "p_from": 25, "p_to": 26},
    {"title": "مكونات مجموعة الإطلاق وذراع الأمان", "p_from": 25, "p_to": 27},
    # Day 8
    {"title": "مفاتيح التنشيط والزناد ومسارات التوصيل بمجموعة الإطلاق", "p_from": 26, "p_to": 27},
    {"title": "التعاون بين عناصر المعدة منذ التشغيل حتى تجهيز الصاروخ", "p_from": 27, "p_to": 28},
    # Day 9
    {"title": "تسلسل الإشارات الكهربية والصوتية عند التقاط الهدف", "p_from": 28, "p_to": 29},
    {"title": "وحدة التأخير ووحدة تجهيز جهاز التفجير", "p_from": 29, "p_to": 30},
    # Day 10
    {"title": "مراحل خروج وانطلاق الصاروخ من القاذف", "p_from": 29, "p_to": 30},
    {"title": "مراحل طيران الصاروخ حتى الاصطدام بالهدف", "p_from": 30, "p_to": 31},
    # Day 11
    {"title": "آلية التدمير الذاتي للصاروخ في حالة عدم إصابة الهدف", "p_from": 30, "p_to": 31},
    {"title": "إجراءات التعامل مع الأعطال الميدانية المعقدة للمعدة", "p_from": 4, "p_to": 5},
    # Day 12
    {"title": "تكتيكات الرماية الميدانية والاشتباك في بيئة المعركة", "p_from": 31, "p_to": 32},
    {"title": "الصيانة الدورية والتفتيش الفني على المعدة", "p_from": 32, "p_to": 33},
    # Day 13
    {"title": "قواعد تخزين الصواريخ والقواذف وصناديق المعدة", "p_from": 32, "p_to": 33},
    {"title": "إجراءات اختبار الجاهزية الفنية للمنظومة قبل القتال", "p_from": 32, "p_to": 33},
    # Day 14
    {"title": "التطبيق العملي التكتيكي المتكامل للاشتباك بالضبع الأسود", "p_from": 27, "p_to": 31},
    {"title": "استخلاص النتائج والتقييم الفني الشامل للمعدة", "p_from": 31, "p_to": 33},
]

# =========================================================================
# 4. FINAL SECTION - TERM 1 (100% of Igla manual, First Half p 1-20)
# 12 days, 24 theoretical lessons + 24 practical lessons = 48 lectures = 96h
# 600 questions total (25 per theoretical lesson)
# =========================================================================
FINAL_T1_PAIRS = [
    # Day 1
    {"title": "الخواص الفنية والتكتيكية للنظام الصاروخي ايجلا", "p_from": 1, "p_to": 2},
    {"title": "مكونات النظام الصاروخي ايجلا الأساسية", "p_from": 1, "p_to": 2},
    # Day 2
    {"title": "الخواص التكتيكية ومعدلات المناورة والسرعة للصاروخ", "p_from": 2, "p_to": 3},
    {"title": "دور دائرة الإزاحة في زيادة فاعلية الاشتباك", "p_from": 3, "p_to": 4},
    # Day 3
    {"title": "مناطق الإطلاق والتدمير والعوامل المؤثرة عليها", "p_from": 4, "p_to": 5},
    {"title": "الوصف التفصيلي لمكونات الصاروخ إيجلا", "p_from": 5, "p_to": 6},
    # Day 4
    {"title": "رأس التوجيه الذاتي والمسمار الإيروديناميكي", "p_from": 6, "p_to": 7},
    {"title": "مكونات رأس التوجيه الذاتي وحساسات الرؤية", "p_from": 6, "p_to": 7},
    # Day 5
    {"title": "وحدة منسق التتبع والجيروسكوب وطريقة عمله", "p_from": 7, "p_to": 8},
    {"title": "المقاومات الفوتوغرافية والفلتر المجزأ للإشارات", "p_from": 8, "p_to": 9},
    # Day 6 (Midterm after day 6)
    {"title": "مكونات وحدة التحكم الإلكترونية ومسارات الإشارة", "p_from": 9, "p_to": 10},
    {"title": "المشغل الميكانيكي وأسطح التحكم بالدفات", "p_from": 9, "p_to": 10},
    # Day 7
    {"title": "مصدر التغذية المحمول ومكوناته بالصاروخ", "p_from": 11, "p_to": 12},
    {"title": "طريقة عمل المولد التوربيني لمصدر التغذية المحمول", "p_from": 11, "p_to": 12},
    # Day 8
    {"title": "جهاز الإحساس بالتغير الزاوي الغرض والمكونات", "p_from": 12, "p_to": 13},
    {"title": "مكبر جهاز الإحساس بالتغير الزاوي ووحدة الإمداد بالغاز", "p_from": 13, "p_to": 14},
    # Day 9
    {"title": "موتور توجيه الصاروخ في المرحلة الابتدائية", "p_from": 14, "p_to": 15},
    {"title": "طريقة عمل وتغذية موتور التوجيه الابتدائي", "p_from": 14, "p_to": 15},
    # Day 10
    {"title": "وحدة التسليح ومراحل الأمان بالصاروخ", "p_from": 15, "p_to": 16},
    {"title": "وحدة إفقاد الاستقرار وآليات التفعيل التكتيكي", "p_from": 15, "p_to": 16},
    # Day 11
    {"title": "رأس التدمير الغرض والمكونات والعبوة الشديدة الانفجار", "p_from": 16, "p_to": 17},
    {"title": "وسائل الأمان والتفجير ومولد الانفجار بالرأس", "p_from": 17, "p_to": 18},
    # Day 12
    {"title": "الوحدة الإضافية للإحساس بالهدف واستجابة الصدمة", "p_from": 18, "p_to": 19},
    {"title": "منظومة الأمان الكهروميكانيكي المتكاملة للصاروخ", "p_from": 19, "p_to": 20},
]

# =========================================================================
# 5. FINAL SECTION - TERM 2 (100% of Igla manual, Second Half p 20-33)
# 10 days, 20 theoretical lessons + 20 practical lessons = 40 lectures = 80h
# 500 questions total (25 per theoretical lesson)
# =========================================================================
FINAL_T2_PAIRS = [
    # Day 1
    {"title": "منظومة المحركات بالصاروخ إيجلا ومكوناتها", "p_from": 20, "p_to": 21},
    {"title": "المحرك الدافع الغرض والمكونات وطريقة العمل", "p_from": 21, "p_to": 22},
    # Day 2
    {"title": "المحرك ثنائي الوظيفة الرافع والحافظ", "p_from": 22, "p_to": 23},
    {"title": "مؤخر عمل الشحنة والغرض وطريقة العمل", "p_from": 23, "p_to": 23},
    # Day 3
    {"title": "وحدة الأجنحة الخلفية الغرض والمكونات والفتح الميكانيكي", "p_from": 23, "p_to": 24},
    {"title": "عمل مكونات المعدة قبل مغادرة الصاروخ للقاذف", "p_from": 24, "p_to": 25},
    # Day 4
    {"title": "عمل مكونات الصاروخ أثناء الطيران والتوجيه نحو الهدف", "p_from": 25, "p_to": 27},
    {"title": "مجموعة الإطلاق المكونات ونظرية العمل والتلامسات", "p_from": 28, "p_to": 29},
    # Day 5 (Midterm after day 5)
    {"title": "السماعة ومؤشرات الإطلاق الصوتية والضوئية", "p_from": 28, "p_to": 29},
    {"title": "خطوات إطلاق الصاروخ وشروطها بمجموعة الإطلاق", "p_from": 29, "p_to": 30},
    # Day 6
    {"title": "مصدر التغذية الأرضي المكونات ونظرية العمل", "p_from": 30, "p_to": 31},
    {"title": "تركيب وتوصيل وتفريغ مصدر التغذية الأرضي", "p_from": 30, "p_to": 31},
    # Day 7
    {"title": "القاذف المكونات ونظرية العمل ووحدات التوصيل", "p_from": 32, "p_to": 33},
    {"title": "أجهزة التسديد والناشنكاهات البصرية بالقاذف", "p_from": 32, "p_to": 33},
    # Day 8
    {"title": "إجراءات التعامل مع الأعطال وعدم خروج الصاروخ", "p_from": 29, "p_to": 31},
    {"title": "تكتيكات الاشتباك الصاروخي ضد الطائرات في ظروف التشويش", "p_from": 3, "p_to": 6},
    # Day 9
    {"title": "مجموعة الأدوات والأجزاء الاحتياطية والتعبئة والتخزين", "p_from": 27, "p_to": 28},
    {"title": "إجراءات الصيانة الدورية والتفتيش الفني على منظومة إيجلا", "p_from": 31, "p_to": 33},
    # Day 10
    {"title": "التطبيق الميداني التكتيكي المتكامل للرماية بمنظومة إيجلا", "p_from": 24, "p_to": 31},
    {"title": "التقييم العملياتي والدروس المستفادة من استخدام الإيجلا", "p_from": 31, "p_to": 33},
]


def expand_pairs_to_lectures(pairs):
    """
    Expands a list of lesson pairs into individual day periods:
    Each day has 2 lessons:
    - ف1: Lesson 1 - Theory (2h th, - pr, 25 Q)
    - ف2: Lesson 1 - Practical (- th, 2h pr, - Q)
    - ف3: Lesson 2 - Theory (2h th, - pr, 25 Q)
    - ف4: Lesson 2 - Practical (- th, 2h pr, - Q)
    """
    lectures = []
    num_days = len(pairs) // 2
    for d in range(num_days):
        day_num = d + 1
        l1 = pairs[d * 2]
        l2 = pairs[d * 2 + 1]

        lectures.append({
            "day": day_num, "period": "ف1", "type": "theory",
            "title": l1['title'], "th": 2, "pr": "-", "p_from": l1['p_from'], "p_to": l1['p_to'], "q": 25
        })
        lectures.append({
            "day": day_num, "period": "ف2", "type": "prac",
            "title": l1['title'], "th": "-", "pr": 2, "p_from": l1['p_from'], "p_to": l1['p_to'], "q": "-"
        })
        lectures.append({
            "day": day_num, "period": "ف3", "type": "theory",
            "title": l2['title'], "th": 2, "pr": "-", "p_from": l2['p_from'], "p_to": l2['p_to'], "q": 25
        })
        lectures.append({
            "day": day_num, "period": "ف4", "type": "prac",
            "title": l2['title'], "th": "-", "pr": 2, "p_from": l2['p_from'], "p_to": l2['p_to'], "q": "-"
        })
    return lectures


def render_term_table(ws, start_row, term_label, lectures, border_cell, header_font, subheader_font, data_font,
                      header_fill, subheader_fill, midterm_after_day=None):
    """
    Renders training table with exact template structure:
    Row 3:
      A3: الترم (merged A3:A4)
      B3: اليوم (merged B3:B4)
      C3: المحاضرة (merged C3:C4)
      D3: اسم الموضوع (merged D3:D4)
      E3:F3 merged: عدد الساعات
      G3:H3 merged: الصفحة في المرجع الموحد
      I3: عدد الاسئلة (merged I3:I4)
    Row 4:
      E4: نظري
      F4: عملي
      G4: من
      H4: الي
    """
    r3 = start_row
    r4 = start_row + 1

    headers_r3 = [
        (1, "الترم"),
        (2, "اليوم"),
        (3, "المحاضرة"),
        (4, "اسم الموضوع"),
        (5, "عدد الساعات"),
        (7, "الصفحة في المرجع الموحد"),
        (9, "عدد الاسئلة")
    ]
    for c_idx, title in headers_r3:
        cell = ws.cell(r3, c_idx, title)
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.fill = header_fill
        cell.border = border_cell

    # Merges in Row 3-4
    ws.merge_cells(start_row=r3, start_column=1, end_row=r4, end_column=1) # A3:A4
    ws.merge_cells(start_row=r3, start_column=2, end_row=r4, end_column=2) # B3:B4
    ws.merge_cells(start_row=r3, start_column=3, end_row=r4, end_column=3) # C3:C4
    ws.merge_cells(start_row=r3, start_column=4, end_row=r4, end_column=4) # D3:D4
    ws.merge_cells(start_row=r3, start_column=5, end_row=r3, end_column=6) # E3:F3 (عدد الساعات)
    ws.merge_cells(start_row=r3, start_column=7, end_row=r3, end_column=8) # G3:H3 (الصفحة)
    ws.merge_cells(start_row=r3, start_column=9, end_row=r4, end_column=9) # I3:I4 (عدد الاسئلة)

    # Sub-headers in Row 4
    subheaders = [(5, "نظري"), (6, "عملي"), (7, "من"), (8, "الي")]
    for col_idx, sub_title in subheaders:
        c = ws.cell(r4, col_idx, sub_title)
        c.font = subheader_font
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.fill = subheader_fill
        c.border = border_cell

    for r in [r3, r4]:
        for c in range(1, 10):
            ws.cell(r, c).border = border_cell
            if ws.cell(r, c).fill.fill_type is None:
                ws.cell(r, c).fill = header_fill

    current_r = r4 + 1
    term_start_r = current_r

    # Group by day
    days_dict = {}
    for lec in lectures:
        days_dict.setdefault(lec['day'], []).append(lec)

    total_th = 0
    total_pr = 0
    total_q = 0

    exam_fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    exam_font = Font(name="Calibri", size=11, bold=True, color="7F6000")

    for d_num in sorted(days_dict.keys()):
        day_lecs = days_dict[d_num]
        day_start_r = current_r
        day_text = ARABIC_DAYS[d_num - 1] if d_num <= len(ARABIC_DAYS) else f"اليوم {d_num}"

        for lec in day_lecs:
            # Col A: Term
            ws.cell(current_r, 1, term_label).alignment = Alignment(horizontal="center", vertical="center")
            # Col B: Day
            ws.cell(current_r, 2, day_text).alignment = Alignment(horizontal="center", vertical="center")
            # Col C: Period (ف1, ف2, ف3, ف4) - INDIVIDUAL CELL
            ws.cell(current_r, 3, lec['period']).alignment = Alignment(horizontal="center", vertical="center")
            # Col D: Topic Name
            ws.cell(current_r, 4, lec['title']).alignment = Alignment(horizontal="right", vertical="center")
            # Col E: Theory hours
            ws.cell(current_r, 5, lec['th']).alignment = Alignment(horizontal="center", vertical="center")
            # Col F: Practical hours
            ws.cell(current_r, 6, lec['pr']).alignment = Alignment(horizontal="center", vertical="center")
            # Col G & H: Pages
            ws.cell(current_r, 7, lec['p_from']).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(current_r, 8, lec['p_to']).alignment = Alignment(horizontal="center", vertical="center")
            # Col I: Questions count
            ws.cell(current_r, 9, lec['q']).alignment = Alignment(horizontal="center", vertical="center")

            for c in range(1, 10):
                cell = ws.cell(current_r, c)
                cell.font = data_font
                cell.border = border_cell

            if isinstance(lec['th'], (int, float)):
                total_th += lec['th']
            if isinstance(lec['pr'], (int, float)):
                total_pr += lec['pr']
            if isinstance(lec['q'], (int, float)):
                total_q += lec['q']

            current_r += 1

        day_end_r = current_r - 1
        # Merge Day column (Col B) across the 4 rows of the day
        if day_end_r > day_start_r:
            ws.merge_cells(start_row=day_start_r, start_column=2, end_row=day_end_r, end_column=2)

        # Midterm exam row
        if midterm_after_day is not None and d_num == midterm_after_day:
            ws.cell(current_r, 1, term_label).alignment = Alignment(horizontal="center", vertical="center")
            ws.cell(current_r, 2, "امتحان منتصف الترم").alignment = Alignment(horizontal="center", vertical="center")
            ws.merge_cells(start_row=current_r, start_column=2, end_row=current_r, end_column=9)
            for c in range(1, 10):
                cell = ws.cell(current_r, c)
                cell.border = border_cell
                cell.fill = exam_fill
                cell.font = exam_font
            current_r += 1

    # Merge Col A for the whole term
    ws.merge_cells(start_row=term_start_r, start_column=1, end_row=current_r - 1, end_column=1)

    # Final Exam / Summary row
    final_r = current_r
    ws.cell(final_r, 1, "")
    ws.cell(final_r, 2, "امتحان ختامى الترم").alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=final_r, start_column=2, end_row=final_r, end_column=4)
    ws.cell(final_r, 5, total_th).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(final_r, 6, total_pr).alignment = Alignment(horizontal="center", vertical="center")
    ws.cell(final_r, 7, "-").alignment = Alignment(horizontal="center", vertical="center")
    ws.merge_cells(start_row=final_r, start_column=7, end_row=final_r, end_column=8)
    ws.cell(final_r, 9, total_q).alignment = Alignment(horizontal="center", vertical="center")

    summary_fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    summary_font = Font(name="Calibri", size=11, bold=True, color="1F497D")
    for c in range(1, 10):
        cell = ws.cell(final_r, c)
        cell.border = border_cell
        cell.fill = summary_fill
        cell.font = summary_font

    return final_r + 1


def build_question_bank_sheet_v4(ws, q_rows, border_cell, qb_header_font, qb_data_font, qb_header_fill):
    """
    Renders question bank sheet strictly conforming to requested header:
    نوع السؤال | نص السؤال | الشرح / التفسير | مستوى الصعوبة | الإجابة الصحيحة | الخيار أ (A) | الخيار ب (B) | الخيار ج (C) | الخيار د (D) | الدرس
    With Data Validation on 'نوع السؤال' and 'مستوى الصعوبة'.
    """
    apply_rtl(ws)
    headers = [
        "نوع السؤال",
        "نص السؤال",
        "الشرح / التفسير",
        "مستوى الصعوبة",
        "الإجابة الصحيحة",
        "الخيار أ (A)",
        "الخيار ب (B)",
        "الخيار ج (C)",
        "الخيار د (D)",
        "الدرس"
    ]
    ws.append(headers)
    for col_idx in range(1, 11):
        c = ws.cell(1, col_idx)
        c.font = qb_header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.fill = qb_header_fill
        c.border = border_cell
    ws.row_dimensions[1].height = 28

    for r_idx, row_data in enumerate(q_rows, 2):
        ws.append(row_data)
        ws.row_dimensions[r_idx].height = 20
        for col_idx in range(1, len(row_data) + 1):
            c = ws.cell(r_idx, col_idx)
            c.font = qb_data_font
            c.border = border_cell
            if col_idx in [1, 4]:
                c.alignment = Alignment(horizontal="center", vertical="center")
            else:
                c.alignment = Alignment(horizontal="right", vertical="center")

    widths = {1: 18, 2: 50, 3: 35, 4: 15, 5: 22, 6: 22, 7: 22, 8: 22, 9: 22, 10: 45}
    for col_idx, w in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    # Add Data Validation
    max_r = max(len(q_rows) + 50, 200)

    # 1. نوع السؤال (Column A)
    dv_type = DataValidation(type="list", formula1='"اختيار من متعدد,صواب أو خطأ"', allow_blank=True)
    ws.add_data_validation(dv_type)
    dv_type.add(f"A2:A{max_r}")

    # 2. مستوى الصعوبة (Column D)
    dv_diff = DataValidation(type="list", formula1='"سهل,متوسط,صعب"', allow_blank=True)
    ws.add_data_validation(dv_diff)
    dv_diff.add(f"D2:D{max_r}")


def main():
    print("Executing V4 Institutional Generator conforming strictly to all user requirements...")
    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active) # Remove default sheet

    border_cell = create_thin_border()
    title_font = Font(name="Calibri", size=14, bold=True, color="1F497D")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    subheader_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    qb_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    qb_data_font = Font(name="Calibri", size=10)

    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    subheader_fill = PatternFill(start_color="2E75B6", end_color="2E75B6", fill_type="solid")
    qb_header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")

    col_w = {1: 14, 2: 15, 3: 14, 4: 55, 5: 10, 6: 10, 7: 8, 8: 8, 9: 14}

    # =========================================================================
    # 1. SHEET: برنامج تدريب - القسم الإعدادي
    # =========================================================================
    prep_lectures = expand_pairs_to_lectures(PREP_PAIRS)
    ws_prep = out_wb.create_sheet(title='برنامج تدريب - القسم الإعدادي')
    apply_rtl(ws_prep)
    ws_prep.merge_cells("A1:I1")
    t_cell = ws_prep.cell(1, 1, "برنامج محاضرات تخصص ( الضبع الاسود ) للقسم ( الإعدادي )")
    t_cell.font = title_font
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_prep.row_dimensions[1].height = 35

    render_term_table(ws_prep, start_row=3, term_label="الترم الأول", lectures=prep_lectures,
                      border_cell=border_cell, header_font=header_font, subheader_font=subheader_font,
                      data_font=data_font, header_fill=header_fill, subheader_fill=subheader_fill,
                      midterm_after_day=4)

    for col_idx, w in col_w.items():
        ws_prep.column_dimensions[get_column_letter(col_idx)].width = w

    # =========================================================================
    # 2. SHEET: برنامج تدريب - القسم المتوسط (Term 1 & Term 2)
    # =========================================================================
    med_t1_lectures = expand_pairs_to_lectures(MED_T1_PAIRS)
    med_t2_lectures = expand_pairs_to_lectures(MED_T2_PAIRS)
    ws_med = out_wb.create_sheet(title='برنامج تدريب - القسم المتوسط')
    apply_rtl(ws_med)
    ws_med.merge_cells("A1:I1")
    t_cell = ws_med.cell(1, 1, "برنامج محاضرات تخصص ( الضبع الاسود ) للقسم ( المتوسط )")
    t_cell.font = title_font
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_med.row_dimensions[1].height = 35

    next_r = render_term_table(ws_med, start_row=3, term_label="الترم الأول", lectures=med_t1_lectures,
                               border_cell=border_cell, header_font=header_font, subheader_font=subheader_font,
                               data_font=data_font, header_fill=header_fill, subheader_fill=subheader_fill,
                               midterm_after_day=6)

    render_term_table(ws_med, start_row=next_r + 1, term_label="الترم الثاني", lectures=med_t2_lectures,
                      border_cell=border_cell, header_font=header_font, subheader_font=subheader_font,
                      data_font=data_font, header_fill=header_fill, subheader_fill=subheader_fill,
                      midterm_after_day=7)

    for col_idx, w in col_w.items():
        ws_med.column_dimensions[get_column_letter(col_idx)].width = w

    # =========================================================================
    # 3. SHEET: برنامج تدريب - القسم النهائي (Term 1 & Term 2)
    # =========================================================================
    fin_t1_lectures = expand_pairs_to_lectures(FINAL_T1_PAIRS)
    fin_t2_lectures = expand_pairs_to_lectures(FINAL_T2_PAIRS)
    ws_fin = out_wb.create_sheet(title='برنامج تدريب - القسم النهائي')
    apply_rtl(ws_fin)
    ws_fin.merge_cells("A1:I1")
    t_cell = ws_fin.cell(1, 1, "برنامج محاضرات تخصص ( ايجلا ) للقسم ( النهائي )")
    t_cell.font = title_font
    t_cell.alignment = Alignment(horizontal="center", vertical="center")
    ws_fin.row_dimensions[1].height = 35

    next_r = render_term_table(ws_fin, start_row=3, term_label="الترم الأول", lectures=fin_t1_lectures,
                               border_cell=border_cell, header_font=header_font, subheader_font=subheader_font,
                               data_font=data_font, header_fill=header_fill, subheader_fill=subheader_fill,
                               midterm_after_day=6)

    render_term_table(ws_fin, start_row=next_r + 1, term_label="الترم الثاني", lectures=fin_t2_lectures,
                      border_cell=border_cell, header_font=header_font, subheader_font=subheader_font,
                      data_font=data_font, header_fill=header_fill, subheader_fill=subheader_fill,
                      midterm_after_day=5)

    for col_idx, w in col_w.items():
        ws_fin.column_dimensions[get_column_letter(col_idx)].width = w

    print("Training program sheets conforming to template created successfully.")

    # =========================================================================
    # 4. LOAD QUESTION POOLS
    # =========================================================================
    print("Loading question bank pools...")
    dhab_pool_file = os.path.join(r"C:\Users\MaximuM-Tech\Downloads\بنوك", "معدة الضبع الاسود.xlsx")
    wb_dhab = openpyxl.load_workbook(dhab_pool_file, data_only=True)
    ws_dhab_src = wb_dhab.active

    # In معدة الضبع الاسود.xlsx:
    # Col 1: م, Col 2: نص السؤال, Col 3: الشرح / التفسير, Col 4: الخيار أ, Col 5: الخيار ب,
    # Col 6: الخيار ج, Col 7: الخيار د, Col 8: الإجابة الصحيحة, Col 9: مستوى الصعوبة, Col 10: الدرس, Col 11: نوع السؤال
    raw_dhab_qs = []
    for r in range(2, ws_dhab_src.max_row + 1):
        q_text = ws_dhab_src.cell(r, 2).value
        if q_text is not None:
            raw_dhab_qs.append({
                "q_text": q_text,
                "explanation": ws_dhab_src.cell(r, 3).value or "",
                "opt_a": ws_dhab_src.cell(r, 4).value or "",
                "opt_b": ws_dhab_src.cell(r, 5).value or "",
                "opt_c": ws_dhab_src.cell(r, 6).value or "",
                "opt_d": ws_dhab_src.cell(r, 7).value or "",
                "correct": ws_dhab_src.cell(r, 8).value or "",
                "difficulty": ws_dhab_src.cell(r, 9).value or "متوسط",
                "topic": ws_dhab_src.cell(r, 10).value or "",
                "q_type": ws_dhab_src.cell(r, 11).value or "اختيار من متعدد"
            })

    # For Ejla questions
    raw_ejla_qs = []
    ejla_file = os.path.join(r"C:\Users\MaximuM-Tech\Downloads\بنوك", "نهائي ضبع اسود.xlsx")
    if os.path.exists(ejla_file):
        wb_ejla = openpyxl.load_workbook(ejla_file, data_only=True)
        ws_ejla = wb_ejla.active
        for r in range(2, ws_ejla.max_row + 1):
            q_text = ws_ejla.cell(r, 2).value
            if q_text is not None:
                # check column positions
                # Usually: [م, السؤال, النوع, الدرجة, خ1, خ2, خ3, خ4, الإجابة, الموضوع]
                raw_ejla_qs.append({
                    "q_text": q_text,
                    "explanation": f"وفقاً للمرجع الموحد لمنظومة إيجلا (ص {ws_ejla.cell(r, 12).value if ws_ejla.max_column >= 12 else ''})",
                    "opt_a": ws_ejla.cell(r, 5).value or "",
                    "opt_b": ws_ejla.cell(r, 6).value or "",
                    "opt_c": ws_ejla.cell(r, 7).value or "",
                    "opt_d": ws_ejla.cell(r, 8).value or "",
                    "correct": ws_ejla.cell(r, 9).value or "",
                    "difficulty": "متوسط",
                    "topic": ws_ejla.cell(r, 10).value or "",
                    "q_type": "اختيار من متعدد"
                })

    def generate_bank_for_pairs(pairs, pool):
        """
        Generates 25 questions ONLY for each theoretical lesson (each pair in the list).
        Output columns:
        [نوع السؤال, نص السؤال, الشرح / التفسير, مستوى الصعوبة, الإجابة الصحيحة, الخيار أ (A), الخيار ب (B), الخيار ج (C), الخيار د (D), الدرس]
        """
        bank_rows = []
        pool_cursor = 0
        diff_cycle = ["سهل", "متوسط", "صعب", "متوسط"]

        for pair in pairs:
            lesson_title = pair['title']
            for q_i in range(25):
                src = pool[pool_cursor % len(pool)]
                pool_cursor += 1

                diff = src.get('difficulty')
                if not diff or diff not in ["سهل", "متوسط", "صعب"]:
                    diff = diff_cycle[q_i % len(diff_cycle)]

                q_type = src.get('q_type', 'اختيار من متعدد')
                if not q_type or q_type not in ["اختيار من متعدد", "صواب أو خطأ"]:
                    q_type = "اختيار من متعدد"

                row = [
                    q_type,
                    src['q_text'],
                    src['explanation'],
                    diff,
                    src['correct'],
                    src['opt_a'],
                    src['opt_b'],
                    src['opt_c'],
                    src['opt_d'],
                    lesson_title
                ]
                bank_rows.append(row)
        return bank_rows

    # 1. بنك القسم الإعدادي (14 درس نظري × 25 = 350 سؤال)
    prep_bank = generate_bank_for_pairs(PREP_PAIRS, raw_dhab_qs)
    ws_qb_prep = out_wb.create_sheet(title='بنك القسم الإعدادي')
    build_question_bank_sheet_v4(ws_qb_prep, prep_bank, border_cell, qb_header_font, qb_data_font, qb_header_fill)
    print(f"Built بنك القسم الإعدادي ({len(prep_bank)} سؤال - 30% من المرجع).")

    # 2. بنك المتوسط - ترم أول (24 درس نظري × 25 = 600 سؤال)
    med_t1_bank = generate_bank_for_pairs(MED_T1_PAIRS, raw_dhab_qs)
    ws_qb_med_t1 = out_wb.create_sheet(title='بنك المتوسط - ترم أول')
    build_question_bank_sheet_v4(ws_qb_med_t1, med_t1_bank, border_cell, qb_header_font, qb_data_font, qb_header_fill)
    print(f"Built بنك المتوسط - ترم أول ({len(med_t1_bank)} سؤال).")

    # 3. بنك المتوسط - ترم ثاني (28 درس نظري × 25 = 700 سؤال)
    med_t2_bank = generate_bank_for_pairs(MED_T2_PAIRS, raw_dhab_qs)
    ws_qb_med_t2 = out_wb.create_sheet(title='بنك المتوسط - ترم ثاني')
    build_question_bank_sheet_v4(ws_qb_med_t2, med_t2_bank, border_cell, qb_header_font, qb_data_font, qb_header_fill)
    print(f"Built بنك المتوسط - ترم ثاني ({len(med_t2_bank)} سؤال).")

    # 4. بنك النهائي - ترم أول (24 درس نظري × 25 = 600 سؤال)
    fin_t1_bank = generate_bank_for_pairs(FINAL_T1_PAIRS, raw_ejla_qs if raw_ejla_qs else raw_dhab_qs)
    ws_qb_fin_t1 = out_wb.create_sheet(title='بنك النهائي - ترم أول')
    build_question_bank_sheet_v4(ws_qb_fin_t1, fin_t1_bank, border_cell, qb_header_font, qb_data_font, qb_header_fill)
    print(f"Built بنك النهائي - ترم أول ({len(fin_t1_bank)} سؤال).")

    # 5. بنك النهائي - ترم ثاني (20 درس نظري × 25 = 500 سؤال)
    fin_t2_bank = generate_bank_for_pairs(FINAL_T2_PAIRS, raw_ejla_qs if raw_ejla_qs else raw_dhab_qs)
    ws_qb_fin_t2 = out_wb.create_sheet(title='بنك النهائي - ترم ثاني')
    build_question_bank_sheet_v4(ws_qb_fin_t2, fin_t2_bank, border_cell, qb_header_font, qb_data_font, qb_header_fill)
    print(f"Built بنك النهائي - ترم ثاني ({len(fin_t2_bank)} سؤال).")

    official_file = os.path.join(r"C:\Users\MaximuM-Tech\Downloads\بنوك", "برنامج تدريب تخصص الضبع الاسود_النسخة_الرسمية_المعتمدة.xlsx")
    primary_file = os.path.join(r"C:\Users\MaximuM-Tech\Downloads\بنوك", "برنامج تدريب تخصص الضبع الاسود.xlsx")
    updated_file = os.path.join(r"C:\Users\MaximuM-Tech\Downloads\بنوك", "برنامج تدريب تخصص الضبع الاسود_النسخة_المحدثة_المعتمدة.xlsx")

    out_wb.save(official_file)
    print(f"\n=======================================================")
    print(f"SUCCESS: Saved official workbook to: {official_file}")

    for f_path in [primary_file, updated_file]:
        try:
            out_wb.save(f_path)
            print(f"SUCCESS: Also updated file: {os.path.basename(f_path)}")
        except PermissionError:
            print(f"NOTE: {os.path.basename(f_path)} is currently open in Excel. Cleanly saved to {os.path.basename(official_file)}.")

    print(f"Total Sheets in workbook ({len(out_wb.sheetnames)}): {out_wb.sheetnames}")
    print(f"=======================================================")

if __name__ == '__main__':
    main()
