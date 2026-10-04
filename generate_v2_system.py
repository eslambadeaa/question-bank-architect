"""
V2 Institutional Generator for Training Programs & Question Banks
Implements:
1. Unified Lesson Title in Column D (اسم الموضوع) - exactly identical for theory and practical sessions.
2. Split "عدد الساعات" into two sub-columns in rows 3-4:
   - Column E: نظري
   - Column F: عملي
3. Page numbers in Columns G & H (من / الي) grounded in docx text.
4. Total quota preserved:
   - Prep: 56h (14 Th + 14 Pr), 700 questions
   - Med T1: 96h (24 Th + 24 Pr), 1200 questions
   - Med T2: 112h (28 Th + 28 Pr), 1400 questions
   - Final T1: 96h (24 Th + 24 Pr), 1200 questions
   - Final T2: 80h (20 Th + 20 Pr), 1000 questions
"""

import os
import sys
from copy import copy
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

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
# 1. PREPARATORY SCHEDULE (28 lectures: 14 Th + 14 Pr = 56h)
# Unified lesson names from مرجع الضبع الاسود
# =========================================================================
PREP_V2_SCHEDULE = [
    # Day 1
    {"period": "ف1", "type": "theory", "title": "التركيب العام والخواص الفنية والتكتيكية للمعدة", "p_from": 1, "p_to": 3},
    {"period": "ف2", "type": "prac",   "title": "التركيب العام والخواص الفنية والتكتيكية للمعدة", "p_from": 1, "p_to": 3},
    {"period": "ف3", "type": "theory", "title": "فكرة عامة والغرض من المعدة وأوضاع الإطلاق", "p_from": 1, "p_to": 2},
    {"period": "ف4", "type": "prac",   "title": "فكرة عامة والغرض من المعدة وأوضاع الإطلاق", "p_from": 1, "p_to": 2},
    # Day 2
    {"period": "ف1", "type": "theory", "title": "احتياطات الأمان والتحذيرات الواجب مراعاتها عند العمل على المعدة", "p_from": 3, "p_to": 4},
    {"period": "ف2", "type": "prac",   "title": "احتياطات الأمان والتحذيرات الواجب مراعاتها عند العمل على المعدة", "p_from": 3, "p_to": 4},
    {"period": "ف3", "type": "theory", "title": "إجراءات التعامل مع عدم خروج الصاروخ بعد الضغطة الثانية", "p_from": 4, "p_to": 5},
    {"period": "ف4", "type": "prac",   "title": "إجراءات التعامل مع عدم خروج الصاروخ بعد الضغطة الثانية", "p_from": 4, "p_to": 5},
    # Day 3
    {"period": "ف1", "type": "theory", "title": "الغرض من الصاروخ وأوضاع إطلاقه ومكوناته الرئيسية", "p_from": 5, "p_to": 6},
    {"period": "ف2", "type": "theory", "title": "رأس التوجيه الذاتي الغرض والخواص الفنية", "p_from": 6, "p_to": 7},
    {"period": "ف3", "type": "prac",   "title": "رأس التوجيه الذاتي الغرض والخواص الفنية", "p_from": 6, "p_to": 7},
    {"period": "ف4", "type": "theory", "title": "الجهاز العدسي الجيروسكوبي ومكوناته ومهامه", "p_from": 7, "p_to": 8},
    # Day 4 (Midterm exam at day 4)
    {"period": "ف1", "type": "prac",   "title": "الجهاز العدسي الجيروسكوبي ومكوناته ومهامه", "p_from": 7, "p_to": 8},
    {"period": "ف2", "type": "theory", "title": "قرص التعديل والمقاومة الفوتوغرافية وجهاز الموازنة", "p_from": 8, "p_to": 10},
    {"period": "ف3", "type": "theory", "title": "ملفات رأس التوجيه الذاتي ونظام أخذ السرعة العادية", "p_from": 10, "p_to": 11},
    {"period": "ف4", "type": "theory", "title": "نظام التثبيت الكهربي في رأس التوجيه الذاتي", "p_from": 11, "p_to": 12},
    # Day 5
    {"period": "ف1", "type": "prac",   "title": "نظام التثبيت الكهربي في رأس التوجيه الذاتي", "p_from": 11, "p_to": 12},
    {"period": "ف2", "type": "theory", "title": "نظام التوجيه والتتبع الآلي والدليل الآلي وقنطرة هيوستن", "p_from": 13, "p_to": 15},
    {"period": "ف3", "type": "theory", "title": "الجزء الخاص بالدفات ومستودع الغاز ومصدر التغذية المحمول", "p_from": 15, "p_to": 18},
    {"period": "ف4", "type": "theory", "title": "الجزء القتالي وجهاز التفجير ورأس التدمير والمحركات", "p_from": 18, "p_to": 22},
    # Day 6
    {"period": "ف1", "type": "theory", "title": "القاذف ومصدر التغذية الأرضي ومواصفاته ومكوناته", "p_from": 22, "p_to": 25},
    {"period": "ف2", "type": "prac",   "title": "القاذف ومصدر التغذية الأرضي ومواصفاته ومكوناته", "p_from": 22, "p_to": 25},
    {"period": "ف3", "type": "prac",   "title": "القاذف ومصدر التغذية الأرضي ومواصفاته ومكوناته", "p_from": 22, "p_to": 25},
    {"period": "ف4", "type": "theory", "title": "مجموعة الإطلاق وخواصها الفنية ومكوناتها", "p_from": 25, "p_to": 27},
    # Day 7
    {"period": "ف1", "type": "prac",   "title": "مجموعة الإطلاق وخواصها الفنية ومكوناتها", "p_from": 25, "p_to": 27},
    {"period": "ف2", "type": "prac",   "title": "مجموعة الإطلاق وخواصها الفنية ومكوناتها", "p_from": 25, "p_to": 27},
    {"period": "ف3", "type": "theory", "title": "التعاون بين عناصر المعدة ووحدة التأخير ومراحل الطيران", "p_from": 27, "p_to": 31},
    {"period": "ف4", "type": "prac",   "title": "التعاون بين عناصر المعدة ووحدة التأخير ومراحل الطيران", "p_from": 27, "p_to": 31},
]

# =========================================================================
# 2. MEDIUM TERM 1 SCHEDULE (48 lectures: 24 Th + 24 Pr = 96h)
# Unified lesson names from مرجع الضبع الاسود
# =========================================================================
MED_T1_V2_SCHEDULE = [
    # Day 1
    {"period": "ف1", "type": "theory", "title": "التركيب العام والخواص الفنية والتكتيكية لمعدة الضبع الأسود", "p_from": 1, "p_to": 3},
    {"period": "ف2", "type": "prac",   "title": "التركيب العام والخواص الفنية والتكتيكية لمعدة الضبع الأسود", "p_from": 1, "p_to": 3},
    {"period": "ف3", "type": "theory", "title": "فكرة عامة عن المعدة وأغراضها وأوضاع الإطلاق", "p_from": 1, "p_to": 2},
    {"period": "ف4", "type": "prac",   "title": "فكرة عامة عن المعدة وأغراضها وأوضاع الإطلاق", "p_from": 1, "p_to": 2},
    # Day 2
    {"period": "ف1", "type": "theory", "title": "الخواص الفنية والتكتيكية التفصيلية للمعدة العادية والمعدلة", "p_from": 2, "p_to": 4},
    {"period": "ف2", "type": "prac",   "title": "الخواص الفنية والتكتيكية التفصيلية للمعدة العادية والمعدلة", "p_from": 2, "p_to": 4},
    {"period": "ف3", "type": "theory", "title": "احتياطات الأمان والتحذيرات الواجب مراعاتها عند العمل على المعدة", "p_from": 3, "p_to": 4},
    {"period": "ف4", "type": "prac",   "title": "احتياطات الأمان والتحذيرات الواجب مراعاتها عند العمل على المعدة", "p_from": 3, "p_to": 4},
    # Day 3
    {"period": "ف1", "type": "theory", "title": "إجراءات التعامل مع الأعطال وحالات عدم خروج الصاروخ من القاذف", "p_from": 4, "p_to": 5},
    {"period": "ف2", "type": "prac",   "title": "إجراءات التعامل مع الأعطال وحالات عدم خروج الصاروخ من القاذف", "p_from": 4, "p_to": 5},
    {"period": "ف3", "type": "prac",   "title": "إجراءات التعامل مع الأعطال وحالات عدم خروج الصاروخ من القاذف", "p_from": 4, "p_to": 5},
    {"period": "ف4", "type": "theory", "title": "الغرض من الصاروخ وأوضاع إطلاقه ومكوناته الرئيسية", "p_from": 5, "p_to": 6},
    # Day 4
    {"period": "ف1", "type": "prac",   "title": "الغرض من الصاروخ وأوضاع إطلاقه ومكوناته الرئيسية", "p_from": 5, "p_to": 6},
    {"period": "ف2", "type": "theory", "title": "الغرض والخواص الفنية لرأس التوجيه الذاتي", "p_from": 6, "p_to": 7},
    {"period": "ف3", "type": "prac",   "title": "الغرض والخواص الفنية لرأس التوجيه الذاتي", "p_from": 6, "p_to": 7},
    {"period": "ف4", "type": "theory", "title": "الجهاز العدسي الجيروسكوبي ومكوناته ومهامه", "p_from": 7, "p_to": 8},
    # Day 5
    {"period": "ف1", "type": "prac",   "title": "الجهاز العدسي الجيروسكوبي ومكوناته ومهامه", "p_from": 7, "p_to": 8},
    {"period": "ف2", "type": "theory", "title": "قرص التعديل وشكل وطريقة عمله وإنتاج إشارة الخطأ", "p_from": 8, "p_to": 9},
    {"period": "ف3", "type": "prac",   "title": "قرص التعديل وشكل وطريقة عمله وإنتاج إشارة الخطأ", "p_from": 8, "p_to": 9},
    {"period": "ف4", "type": "theory", "title": "المقاومة الفوتوغرافية وجهاز موازنة دوران العضو الدوار", "p_from": 9, "p_to": 10},
    # Day 6 (Midterm day for Term 1)
    {"period": "ف1", "type": "prac",   "title": "المقاومة الفوتوغرافية وجهاز موازنة دوران العضو الدوار", "p_from": 9, "p_to": 10},
    {"period": "ف2", "type": "theory", "title": "ملفات رأس التوجيه الذاتي وتصنيفاتها المختلفة", "p_from": 10, "p_to": 11},
    {"period": "ف3", "type": "prac",   "title": "ملفات رأس التوجيه الذاتي وتصنيفاتها المختلفة", "p_from": 10, "p_to": 11},
    {"period": "ف4", "type": "theory", "title": "نظام أخذ العضو الدوار لسرعته العادية وطريقة العمل", "p_from": 10, "p_to": 11},
    # Day 7
    {"period": "ف1", "type": "prac",   "title": "نظام أخذ العضو الدوار لسرعته العادية وطريقة العمل", "p_from": 10, "p_to": 11},
    {"period": "ف2", "type": "theory", "title": "نظام التثبيت الكهربي ومكوناته وطريقة عمله", "p_from": 11, "p_to": 12},
    {"period": "ف3", "type": "prac",   "title": "نظام التثبيت الكهربي ومكوناته وطريقة عمله", "p_from": 11, "p_to": 12},
    {"period": "ف4", "type": "theory", "title": "نظام التوجيه والتتبع الآلي وفكرة عمله والاقتراب التناسبي", "p_from": 13, "p_to": 14},
    # Day 8
    {"period": "ف1", "type": "prac",   "title": "نظام التوجيه والتتبع الآلي وفكرة عمله والاقتراب التناسبي", "p_from": 13, "p_to": 14},
    {"period": "ف2", "type": "theory", "title": "الدليل الآلي ودائرة العمل الكهربية ومكوناته", "p_from": 14, "p_to": 15},
    {"period": "ف3", "type": "prac",   "title": "الدليل الآلي ودائرة العمل الكهربية ومكوناته", "p_from": 14, "p_to": 15},
    {"period": "ف4", "type": "theory", "title": "جهاز تحديد السرعة الزاوية وفكرة عمله وقنطرة هيوستن", "p_from": 15, "p_to": 15},
    # Day 9
    {"period": "ف1", "type": "prac",   "title": "جهاز تحديد السرعة الزاوية وفكرة عمله وقنطرة هيوستن", "p_from": 15, "p_to": 15},
    {"period": "ف2", "type": "theory", "title": "الجزء الخاص بالدفات ومكوناته ومستودع الغاز المضغوط", "p_from": 15, "p_to": 17},
    {"period": "ف3", "type": "prac",   "title": "الجزء الخاص بالدفات ومكوناته ومستودع الغاز المضغوط", "p_from": 15, "p_to": 17},
    {"period": "ف4", "type": "theory", "title": "مصدر التغذية المحمول على سطح الصاروخ ومكوناته", "p_from": 17, "p_to": 18},
    # Day 10
    {"period": "ف1", "type": "prac",   "title": "مصدر التغذية المحمول على سطح الصاروخ ومكوناته", "p_from": 17, "p_to": 18},
    {"period": "ف2", "type": "theory", "title": "القسم القتالي وجهاز التفجير ووسائل الأمان ورأس التدمير", "p_from": 18, "p_to": 20},
    {"period": "ف3", "type": "prac",   "title": "القسم القتالي وجهاز التفجير ووسائل الأمان ورأس التدمير", "p_from": 18, "p_to": 20},
    {"period": "ف4", "type": "theory", "title": "الجزء الخاص بالمحركات والأجنحة والحلقتين المركزيتين", "p_from": 20, "p_to": 22},
    # Day 11
    {"period": "ف1", "type": "prac",   "title": "الجزء الخاص بالمحركات والأجنحة والحلقتين المركزيتين", "p_from": 20, "p_to": 22},
    {"period": "ف2", "type": "theory", "title": "القاذف ومصدر التغذية الأرضي ومواصفاته ومكوناته", "p_from": 22, "p_to": 25},
    {"period": "ف3", "type": "prac",   "title": "القاذف ومصدر التغذية الأرضي ومواصفاته ومكوناته", "p_from": 22, "p_to": 25},
    {"period": "ف4", "type": "theory", "title": "مجموعة الإطلاق وخواصها الفنية ومكوناتها وطريقة عملها", "p_from": 25, "p_to": 27},
    # Day 12
    {"period": "ف1", "type": "prac",   "title": "مجموعة الإطلاق وخواصها الفنية ومكوناتها وطريقة عملها", "p_from": 25, "p_to": 27},
    {"period": "ف2", "type": "theory", "title": "التعاون بين عناصر المعدة ووحدة التأخير ومراحل الطيران", "p_from": 27, "p_to": 31},
    {"period": "ف3", "type": "prac",   "title": "التعاون بين عناصر المعدة ووحدة التأخير ومراحل الطيران", "p_from": 27, "p_to": 31},
    {"period": "ف4", "type": "prac",   "title": "التعاون بين عناصر المعدة ووحدة التأخير ومراحل الطيران", "p_from": 27, "p_to": 31},
]

# =========================================================================
# 3. MEDIUM TERM 2 SCHEDULE (56 lectures: 28 Th + 28 Pr = 112h)
# Unified lesson names from مرجع الضبع الاسود
# =========================================================================
MED_T2_V2_SCHEDULE = [
    # Day 1
    {"period": "ف1", "type": "theory", "title": "التركيب العام المتقدم والخواص التكتيكية لمعدة الضبع الأسود", "p_from": 1, "p_to": 3},
    {"period": "ف2", "type": "prac",   "title": "التركيب العام المتقدم والخواص التكتيكية لمعدة الضبع الأسود", "p_from": 1, "p_to": 3},
    {"period": "ف3", "type": "theory", "title": "احتياطات الأمان والتحذيرات الصارمة في ظروف القتال والتحرك", "p_from": 3, "p_to": 4},
    {"period": "ف4", "type": "prac",   "title": "احتياطات الأمان والتحذيرات الصارمة في ظروف القتال والتحرك", "p_from": 3, "p_to": 4},
    # Day 2
    {"period": "ف1", "type": "theory", "title": "أوضاع الإطلاق التكتيكية ومكونات المعدة الأساسية", "p_from": 3, "p_to": 4},
    {"period": "ف2", "type": "prac",   "title": "أوضاع الإطلاق التكتيكية ومكونات المعدة الأساسية", "p_from": 3, "p_to": 4},
    {"period": "ف3", "type": "theory", "title": "الصاروخ والغرض منه وأوضاع إطلاقه ومكوناته الرئيسية", "p_from": 5, "p_to": 6},
    {"period": "ف4", "type": "prac",   "title": "الصاروخ والغرض منه وأوضاع إطلاقه ومكوناته الرئيسية", "p_from": 5, "p_to": 6},
    # Day 3
    {"period": "ف1", "type": "theory", "title": "رأس التوجيه الذاتي الغرض والخواص الفنية ومجال التتبع", "p_from": 6, "p_to": 7},
    {"period": "ف2", "type": "prac",   "title": "رأس التوجيه الذاتي الغرض والخواص الفنية ومجال التتبع", "p_from": 6, "p_to": 7},
    {"period": "ف3", "type": "theory", "title": "الجهاز العدسي الجيروسكوبي ومكوناته ومهامه وانعكاس الأشعة", "p_from": 7, "p_to": 8},
    {"period": "ف4", "type": "prac",   "title": "الجهاز العدسي الجيروسكوبي ومكوناته ومهامه وانعكاس الأشعة", "p_from": 7, "p_to": 8},
    # Day 4
    {"period": "ف1", "type": "theory", "title": "قرص التعديل وشكل وطريقة عمله وإنتاج إشارة الخطأ", "p_from": 8, "p_to": 9},
    {"period": "ف2", "type": "prac",   "title": "قرص التعديل وشكل وطريقة عمله وإنتاج إشارة الخطأ", "p_from": 8, "p_to": 9},
    {"period": "ف3", "type": "theory", "title": "المقاومة الفوتوغرافية وجهاز الموازنة ودوران العضو الدوار", "p_from": 9, "p_to": 10},
    {"period": "ف4", "type": "prac",   "title": "المقاومة الفوتوغرافية وجهاز الموازنة ودوران العضو الدوار", "p_from": 9, "p_to": 10},
    # Day 5
    {"period": "ف1", "type": "theory", "title": "ملفات رأس التوجيه الذاتي وتصنيفاتها والمجال المغناطيسي", "p_from": 10, "p_to": 11},
    {"period": "ف2", "type": "prac",   "title": "ملفات رأس التوجيه الذاتي وتصنيفاتها والمجال المغناطيسي", "p_from": 10, "p_to": 11},
    {"period": "ف3", "type": "theory", "title": "نظام أخذ العضو الدوار لسرعته العادية وطريقة العمل", "p_from": 10, "p_to": 11},
    {"period": "ف4", "type": "prac",   "title": "نظام أخذ العضو الدوار لسرعته العادية وطريقة العمل", "p_from": 10, "p_to": 11},
    # Day 6
    {"period": "ف1", "type": "theory", "title": "نظام التثبيت الكهربي ومكوناته وطريقة عمله", "p_from": 11, "p_to": 12},
    {"period": "ف2", "type": "prac",   "title": "نظام التثبيت الكهربي ومكوناته وطريقة عمله", "p_from": 11, "p_to": 12},
    {"period": "ف3", "type": "theory", "title": "نظام التوجيه والتتبع الآلي وفكرة عمله ونظرية الاقتراب التناسبي", "p_from": 13, "p_to": 14},
    {"period": "ف4", "type": "prac",   "title": "نظام التوجيه والتتبع الآلي وفكرة عمله ونظرية الاقتراب التناسبي", "p_from": 13, "p_to": 14},
    # Day 7 (Midterm day for Term 2)
    {"period": "ف1", "type": "theory", "title": "الدليل الآلي ودائرة العمل الكهربية ومكوناته والمعدل والمستخلص", "p_from": 14, "p_to": 15},
    {"period": "ف2", "type": "prac",   "title": "الدليل الآلي ودائرة العمل الكهربية ومكوناته والمعدل والمستخلص", "p_from": 14, "p_to": 15},
    {"period": "ف3", "type": "theory", "title": "جهاز تحديد السرعة الزاوية وفكرة عمله وقنطرة هيوستن", "p_from": 15, "p_to": 15},
    {"period": "ف4", "type": "prac",   "title": "جهاز تحديد السرعة الزاوية وفكرة عمله وقنطرة هيوستن", "p_from": 15, "p_to": 15},
    # Day 8
    {"period": "ف1", "type": "theory", "title": "الجزء الخاص بالدفات ومكوناته ومستودع الغاز المضغوط", "p_from": 15, "p_to": 17},
    {"period": "ف2", "type": "prac",   "title": "الجزء الخاص بالدفات ومكوناته ومستودع الغاز المضغوط", "p_from": 15, "p_to": 17},
    {"period": "ف3", "type": "theory", "title": "مصدر التغذية المحمول على سطح الصاروخ ومكوناته والتوربين", "p_from": 17, "p_to": 18},
    {"period": "ف4", "type": "prac",   "title": "مصدر التغذية المحمول على سطح الصاروخ ومكوناته والتوربين", "p_from": 17, "p_to": 18},
    # Day 9
    {"period": "ف1", "type": "theory", "title": "القسم القتالي وجهاز التفجير ووسائل الأمان", "p_from": 18, "p_to": 19},
    {"period": "ف2", "type": "prac",   "title": "القسم القتالي وجهاز التفجير ووسائل الأمان", "p_from": 18, "p_to": 19},
    {"period": "ف3", "type": "theory", "title": "رأس التدمير ومكوناتها والعبوة المتفجرة وتأثير الشظايا", "p_from": 20, "p_to": 20},
    {"period": "ف4", "type": "prac",   "title": "رأس التدمير ومكوناتها والعبوة المتفجرة وتأثير الشظايا", "p_from": 20, "p_to": 20},
    # Day 10
    {"period": "ف1", "type": "theory", "title": "الجزء الخاص بالمحركات والأجنحة والحلقتان المركزيتان", "p_from": 20, "p_to": 21},
    {"period": "ف2", "type": "prac",   "title": "الجزء الخاص بالمحركات والأجنحة والحلقتان المركزيتان", "p_from": 20, "p_to": 21},
    {"period": "ف3", "type": "theory", "title": "طريقة عمل المحرك الدافع والمحرك الرئيسي وأزمنة العمل", "p_from": 21, "p_to": 22},
    {"period": "ف4", "type": "prac",   "title": "طريقة عمل المحرك الدافع والمحرك الرئيسي وأزمنة العمل", "p_from": 21, "p_to": 22},
    # Day 11
    {"period": "ف1", "type": "theory", "title": "القاذف الغرض منه والمواصفات العامة وحماية الرامي", "p_from": 22, "p_to": 23},
    {"period": "ف2", "type": "prac",   "title": "القاذف الغرض منه والمواصفات العامة وحماية الرامي", "p_from": 22, "p_to": 23},
    {"period": "ف3", "type": "theory", "title": "مصدر التغذية الأرضي المواصفات والمكونات والتفاعل الكيميائي", "p_from": 23, "p_to": 25},
    {"period": "ف4", "type": "prac",   "title": "مصدر التغذية الأرضي المواصفات والمكونات والتفاعل الكيميائي", "p_from": 23, "p_to": 25},
    # Day 12
    {"period": "ف1", "type": "theory", "title": "مجموعة الإطلاق الغرض والخواص الفنية ومكوناتها وذراع الأمان", "p_from": 25, "p_to": 27},
    {"period": "ف2", "type": "prac",   "title": "مجموعة الإطلاق الغرض والخواص الفنية ومكوناتها وذراع الأمان", "p_from": 25, "p_to": 27},
    {"period": "ف3", "type": "theory", "title": "التعاون بين عناصر المعدة منذ التشغيل حتى التقاط الهدف", "p_from": 27, "p_to": 29},
    {"period": "ف4", "type": "prac",   "title": "التعاون بين عناصر المعدة منذ التشغيل حتى التقاط الهدف", "p_from": 27, "p_to": 29},
    # Day 13
    {"period": "ف1", "type": "theory", "title": "وحدة التأخير ووحدة تجهيز جهاز التفجير ومراحل الطيران", "p_from": 29, "p_to": 31},
    {"period": "ف2", "type": "prac",   "title": "وحدة التأخير ووحدة تجهيز جهاز التفجير ومراحل الطيران", "p_from": 29, "p_to": 31},
    {"period": "ف3", "type": "theory", "title": "مراحل طيران الصاروخ حتى الاصطدام بالهدف أو التدمير الذاتي", "p_from": 30, "p_to": 31},
    {"period": "ف4", "type": "prac",   "title": "مراحل طيران الصاروخ حتى الاصطدام بالهدف أو التدمير الذاتي", "p_from": 30, "p_to": 31},
    # Day 14
    {"period": "ف1", "type": "theory", "title": "إجراءات التعامل مع الأعطال المعقدة في ظروف القتال الجوي", "p_from": 4, "p_to": 5},
    {"period": "ف2", "type": "prac",   "title": "إجراءات التعامل مع الأعطال المعقدة في ظروف القتال الجوي", "p_from": 4, "p_to": 5},
    {"period": "ف3", "type": "theory", "title": "التطبيق التكتيكي المتكامل للاشتباك الصاروخي والرماية الميدانية", "p_from": 1, "p_to": 31},
    {"period": "ف4", "type": "prac",   "title": "التطبيق التكتيكي المتكامل للاشتباك الصاروخي والرماية الميدانية", "p_from": 1, "p_to": 31},
]

# =========================================================================
# 4. FINAL TERM 1 SCHEDULE (48 lectures: 24 Th + 24 Pr = 96h)
# Unified lesson names from مرجع الايجلا
# =========================================================================
FINAL_T1_V2_SCHEDULE = [
    # Day 1
    {"period": "ف1", "type": "theory", "title": "الخواص الفنية والتكتيكية للنظام الصاروخي ايجلا ومكونات النظام", "p_from": 1, "p_to": 2},
    {"period": "ف2", "type": "prac",   "title": "الخواص الفنية والتكتيكية للنظام الصاروخي ايجلا ومكونات النظام", "p_from": 1, "p_to": 2},
    {"period": "ف3", "type": "theory", "title": "القدرات التكتيكية لمعدة الايجلا ونظرية التوجيه والتحكم", "p_from": 2, "p_to": 3},
    {"period": "ف4", "type": "prac",   "title": "القدرات التكتيكية لمعدة الايجلا ونظرية التوجيه والتحكم", "p_from": 2, "p_to": 3},
    # Day 2
    {"period": "ف1", "type": "theory", "title": "دور دائرة الإزاحة في زيادة فاعلية الاشتباك مع الطائرات الحديثة", "p_from": 3, "p_to": 4},
    {"period": "ف2", "type": "prac",   "title": "دور دائرة الإزاحة في زيادة فاعلية الاشتباك مع الطائرات الحديثة", "p_from": 3, "p_to": 4},
    {"period": "ف3", "type": "theory", "title": "مناطق الإطلاق والتدمير والعوامل المؤثرة عليها", "p_from": 4, "p_to": 5},
    {"period": "ف4", "type": "prac",   "title": "مناطق الإطلاق والتدمير والعوامل المؤثرة عليها", "p_from": 4, "p_to": 5},
    # Day 3
    {"period": "ف1", "type": "theory", "title": "الوصف التفصيلي لمكونات الصاروخ ورأس التوجيه الذاتي", "p_from": 5, "p_to": 7},
    {"period": "ف2", "type": "prac",   "title": "الوصف التفصيلي لمكونات الصاروخ ورأس التوجيه الذاتي", "p_from": 5, "p_to": 7},
    {"period": "ف3", "type": "theory", "title": "وحدة منسق التتبع والجيروسزوب وطريقة عمله", "p_from": 7, "p_to": 8},
    {"period": "ف4", "type": "prac",   "title": "وحدة منسق التتبع والجيروسزوب وطريقة عمله", "p_from": 7, "p_to": 8},
    # Day 4
    {"period": "ف1", "type": "theory", "title": "وظيفة المقاومات الفوتوغرافية والفلتر المجزأ", "p_from": 8, "p_to": 9},
    {"period": "ف2", "type": "prac",   "title": "وظيفة المقاومات الفوتوغرافية والفلتر المجزأ", "p_from": 8, "p_to": 9},
    {"period": "ف3", "type": "theory", "title": "مكونات وحدة التحكم والمشغل الميكانيكي وأسطح التحكم", "p_from": 9, "p_to": 10},
    {"period": "ف4", "type": "prac",   "title": "مكونات وحدة التحكم والمشغل الميكانيكي وأسطح التحكم", "p_from": 9, "p_to": 10},
    # Day 5
    {"period": "ف1", "type": "theory", "title": "مصدر التغذية المحمول ومكوناته والمولد التوربيني", "p_from": 11, "p_to": 12},
    {"period": "ف2", "type": "prac",   "title": "مصدر التغذية المحمول ومكوناته والمولد التوربيني", "p_from": 11, "p_to": 12},
    {"period": "ف3", "type": "theory", "title": "جهاز الإحساس بالتغير الزاوي الغرض والمكونات", "p_from": 12, "p_to": 13},
    {"period": "ف4", "type": "prac",   "title": "جهاز الإحساس بالتغير الزاوي الغرض والمكونات", "p_from": 12, "p_to": 13},
    # Day 6 (Midterm day for Final Term 1)
    {"period": "ف1", "type": "theory", "title": "مكبر جهاز الإحساس بالتغير الزاوي ووحدة الإمداد بالغاز", "p_from": 13, "p_to": 14},
    {"period": "ف2", "type": "prac",   "title": "مكبر جهاز الإحساس بالتغير الزاوي ووحدة الإمداد بالغاز", "p_from": 13, "p_to": 14},
    {"period": "ف3", "type": "theory", "title": "موتور توجيه الصاروخ في المرحلة الابتدائية وطريقة عمله", "p_from": 14, "p_to": 15},
    {"period": "ف4", "type": "prac",   "title": "موتور توجيه الصاروخ في المرحلة الابتدائية وطريقة عمله", "p_from": 14, "p_to": 15},
    # Day 7
    {"period": "ف1", "type": "theory", "title": "وحدة التسليح ووحدة إفقاد الاستقرار", "p_from": 15, "p_to": 16},
    {"period": "ف2", "type": "prac",   "title": "وحدة التسليح ووحدة إفقاد الاستقرار", "p_from": 15, "p_to": 16},
    {"period": "ف3", "type": "theory", "title": "رأس التدمير الغرض والمكونات وطريقة العمل", "p_from": 16, "p_to": 17},
    {"period": "ف4", "type": "prac",   "title": "رأس التدمير الغرض والمكونات وطريقة العمل", "p_from": 16, "p_to": 17},
    # Day 8
    {"period": "ف1", "type": "theory", "title": "وسائل الأمان والتفجير ووحدة الإحساس بالهدف", "p_from": 17, "p_to": 19},
    {"period": "ف2", "type": "prac",   "title": "وسائل الأمان والتفجير ووحدة الإحساس بالهدف", "p_from": 17, "p_to": 19},
    {"period": "ف3", "type": "theory", "title": "المحركات الغرض العام والمحرك الدافع", "p_from": 20, "p_to": 22},
    {"period": "ف4", "type": "prac",   "title": "المحركات الغرض العام والمحرك الدافع", "p_from": 20, "p_to": 22},
    # Day 9
    {"period": "ف1", "type": "theory", "title": "المحرك ثنائي الوظيفة ومؤخر عمل الشحنة", "p_from": 22, "p_to": 23},
    {"period": "ف2", "type": "prac",   "title": "المحرك ثنائي الوظيفة ومؤخر عمل الشحنة", "p_from": 22, "p_to": 23},
    {"period": "ف3", "type": "theory", "title": "وحدة الأجنحة الخلفية الغرض والمكونات وطريقة العمل", "p_from": 23, "p_to": 24},
    {"period": "ف4", "type": "prac",   "title": "وحدة الأجنحة الخلفية الغرض والمكونات وطريقة العمل", "p_from": 23, "p_to": 24},
    # Day 10
    {"period": "ف1", "type": "theory", "title": "عمل مكونات المعدة قبل مغادرة الصاروخ للقاذف", "p_from": 24, "p_to": 25},
    {"period": "ف2", "type": "prac",   "title": "عمل مكونات المعدة قبل مغادرة الصاروخ للقاذف", "p_from": 24, "p_to": 25},
    {"period": "ف3", "type": "theory", "title": "عمل مكونات الصاروخ أثناء الطيران والتعبئة والتخزين", "p_from": 25, "p_to": 27},
    {"period": "ف4", "type": "prac",   "title": "عمل مكونات الصاروخ أثناء الطيران والتعبئة والتخزين", "p_from": 25, "p_to": 27},
    # Day 11
    {"period": "ف1", "type": "theory", "title": "مجموعة الإطلاق المكونات ونظرية العمل والتلامسات والسماعة", "p_from": 28, "p_to": 29},
    {"period": "ف2", "type": "prac",   "title": "مجموعة الإطلاق المكونات ونظرية العمل والتلامسات والسماعة", "p_from": 28, "p_to": 29},
    {"period": "ف3", "type": "theory", "title": "خطوات إطلاق الصاروخ وشروطها بمجموعة الإطلاق", "p_from": 29, "p_to": 30},
    {"period": "ف4", "type": "prac",   "title": "خطوات إطلاق الصاروخ وشروطها بمجموعة الإطلاق", "p_from": 29, "p_to": 30},
    # Day 12
    {"period": "ف1", "type": "theory", "title": "مصدر التغذية الأرضي المكونات ونظرية العمل", "p_from": 30, "p_to": 31},
    {"period": "ف2", "type": "prac",   "title": "مصدر التغذية الأرضي المكونات ونظرية العمل", "p_from": 30, "p_to": 31},
    {"period": "ف3", "type": "theory", "title": "القاذف المكونات ونظرية العمل ووحدات التوصيل", "p_from": 32, "p_to": 33},
    {"period": "ف4", "type": "prac",   "title": "القاذف المكونات ونظرية العمل ووحدات التوصيل", "p_from": 32, "p_to": 33},
]

# =========================================================================
# 5. FINAL TERM 2 SCHEDULE (40 lectures: 20 Th + 20 Pr = 80h)
# Unified lesson names from مرجع الايجلا
# =========================================================================
FINAL_T2_V2_SCHEDULE = [
    # Day 1
    {"period": "ف1", "type": "theory", "title": "الخواص الفنية والتكتيكية للنظام الصاروخي ايجلا ومكونات النظام", "p_from": 1, "p_to": 2},
    {"period": "ف2", "type": "prac",   "title": "الخواص الفنية والتكتيكية للنظام الصاروخي ايجلا ومكونات النظام", "p_from": 1, "p_to": 2},
    {"period": "ف3", "type": "theory", "title": "نظرية التوجيه والتحكم ومناطق الإطلاق والتدمير", "p_from": 2, "p_to": 5},
    {"period": "ف4", "type": "prac",   "title": "نظرية التوجيه والتحكم ومناطق الإطلاق والتدمير", "p_from": 2, "p_to": 5},
    # Day 2
    {"period": "ف1", "type": "theory", "title": "الوصف التفصيلي لمكونات الصاروخ ورأس التوجيه الذاتي", "p_from": 5, "p_to": 8},
    {"period": "ف2", "type": "prac",   "title": "الوصف التفصيلي لمكونات الصاروخ ورأس التوجيه الذاتي", "p_from": 5, "p_to": 8},
    {"period": "ف3", "type": "theory", "title": "وحدة التحكم والمشغل الميكانيكي وأسطح التحكم (الدفات)", "p_from": 9, "p_to": 10},
    {"period": "ف4", "type": "prac",   "title": "وحدة التحكم والمشغل الميكانيكي وأسطح التحكم (الدفات)", "p_from": 9, "p_to": 10},
    # Day 3
    {"period": "ف1", "type": "theory", "title": "مصدر التغذية المحمول وجهاز الإحساس بالتغير الزاوي", "p_from": 11, "p_to": 13},
    {"period": "ف2", "type": "prac",   "title": "مصدر التغذية المحمول وجهاز الإحساس بالتغير الزاوي", "p_from": 11, "p_to": 13},
    {"period": "ف3", "type": "theory", "title": "وحدة الإمداد بالغاز وموتور توجيه الصاروخ فى المرحلة الابتدائية", "p_from": 13, "p_to": 15},
    {"period": "ف4", "type": "prac",   "title": "وحدة الإمداد بالغاز وموتور توجيه الصاروخ فى المرحلة الابتدائية", "p_from": 13, "p_to": 15},
    # Day 4
    {"period": "ف1", "type": "theory", "title": "وحدة التسليح ووحدة إفقاد الاستقرار وآليات التأمين", "p_from": 15, "p_to": 16},
    {"period": "ف2", "type": "prac",   "title": "وحدة التسليح ووحدة إفقاد الاستقرار وآليات التأمين", "p_from": 15, "p_to": 16},
    {"period": "ف3", "type": "theory", "title": "رأس التدمير ووسائل الأمان والتفجير ومولد الانفجار", "p_from": 16, "p_to": 19},
    {"period": "ف4", "type": "prac",   "title": "رأس التدمير ووسائل الأمان والتفجير ومولد الانفجار", "p_from": 16, "p_to": 19},
    # Day 5 (Midterm day for Final Term 2)
    {"period": "ف1", "type": "theory", "title": "المحركات والمحرك الدافع والمحرك ثنائي الوظيفة ومؤخر عمل الشحنة", "p_from": 20, "p_to": 23},
    {"period": "ف2", "type": "prac",   "title": "المحركات والمحرك الدافع والمحرك ثنائي الوظيفة ومؤخر عمل الشحنة", "p_from": 20, "p_to": 23},
    {"period": "ف3", "type": "theory", "title": "وحدة الأجنحة الخلفية وعمل مكونات الصاروخ أثناء الطيران", "p_from": 23, "p_to": 27},
    {"period": "ف4", "type": "prac",   "title": "وحدة الأجنحة الخلفية وعمل مكونات الصاروخ أثناء الطيران", "p_from": 23, "p_to": 27},
    # Day 6
    {"period": "ف1", "type": "theory", "title": "مجموعة الإطلاق ونظرية العمل وخطوات الإطلاق الأربع", "p_from": 28, "p_to": 30},
    {"period": "ف2", "type": "prac",   "title": "مجموعة الإطلاق ونظرية العمل وخطوات الإطلاق الأربع", "p_from": 28, "p_to": 30},
    {"period": "ف3", "type": "theory", "title": "مصدر التغذية الأرضي والقاذف ووحدات التوصيل والناشنكاهات", "p_from": 30, "p_to": 33},
    {"period": "ف4", "type": "prac",   "title": "مصدر التغذية الأرضي والقاذف ووحدات التوصيل والناشنكاهات", "p_from": 30, "p_to": 33},
    # Day 7
    {"period": "ف1", "type": "theory", "title": "إجراءات التعامل مع الأعطال وحالات عدم خروج الصاروخ من القاذف", "p_from": 28, "p_to": 33},
    {"period": "ف2", "type": "prac",   "title": "إجراءات التعامل مع الأعطال وحالات عدم خروج الصاروخ من القاذف", "p_from": 28, "p_to": 33},
    {"period": "ف3", "type": "theory", "title": "تكتيكات الاشتباك الصاروخي ضد الطائرات المقاتلة وطائرات الهليكوبتر", "p_from": 2, "p_to": 5},
    {"period": "ف4", "type": "prac",   "title": "تكتيكات الاشتباك الصاروخي ضد الطائرات المقاتلة وطائرات الهليكوبتر", "p_from": 2, "p_to": 5},
    # Day 8
    {"period": "ف1", "type": "theory", "title": "الرماية الصاروخية في ظروف التشويش الحراري وإطلاق المشاعل الخداعية", "p_from": 3, "p_to": 8},
    {"period": "ف2", "type": "prac",   "title": "الرماية الصاروخية في ظروف التشويش الحراري وإطلاق المشاعل الخداعية", "p_from": 3, "p_to": 8},
    {"period": "ف3", "type": "theory", "title": "استخدام منظومة إيجلا في الدفاع الجوي عن القوات والمنشآت الحيوية", "p_from": 1, "p_to": 5},
    {"period": "ف4", "type": "prac",   "title": "استخدام منظومة إيجلا في الدفاع الجوي عن القوات والمنشآت الحيوية", "p_from": 1, "p_to": 5},
    # Day 9
    {"period": "ف1", "type": "theory", "title": "إجراءات التفتيش الفني والصيانة الدورية للمعدة والتخزين طويل الأمد", "p_from": 25, "p_to": 28},
    {"period": "ف2", "type": "prac",   "title": "إجراءات التفتيش الفني والصيانة الدورية للمعدة والتخزين طويل الأمد", "p_from": 25, "p_to": 28},
    {"period": "ف3", "type": "theory", "title": "التقييم التكتيكي والعملياتي لمنظومات الدفاع الجوي المحمولة على الكتف", "p_from": 1, "p_to": 33},
    {"period": "ف4", "type": "prac",   "title": "التقييم التكتيكي والعملياتي لمنظومات الدفاع الجوي المحمولة على الكتف", "p_from": 1, "p_to": 33},
    # Day 10
    {"period": "ف1", "type": "theory", "title": "التطبيق التكتيكي المتكامل للاشتباك الصاروخي والرماية الميدانية بالايجلا", "p_from": 1, "p_to": 33},
    {"period": "ف2", "type": "prac",   "title": "التطبيق التكتيكي المتكامل للاشتباك الصاروخي والرماية الميدانية بالايجلا", "p_from": 1, "p_to": 33},
    {"period": "ف3", "type": "theory", "title": "استخلاص الدروس المستفادة وتحليل الفاعلية القتالية لمنظومة إيجلا", "p_from": 1, "p_to": 33},
    {"period": "ف4", "type": "prac",   "title": "استخلاص الدروس المستفادة وتحليل الفاعلية القتالية لمنظومة إيجلا", "p_from": 1, "p_to": 33},
]

def add_v2_training_term(
    ws,
    term_name,
    schedule,
    days_count,
    midterm_day,
    th_hours,
    pr_hours,
    tot_questions,
    start_row=3,
    border_cell=None,
    font_header=None,
    font_data=None,
    font_exam=None,
    font_total=None,
    header_fill=None,
    exam_fill=None,
    total_fill=None
):
    """
    Renders one full term in V2 layout:
    Headers:
      A3:A4: الترم
      B3:B4: اليوم
      C3:C4: المحاضرة
      D3:D4: اسم الموضوع
      E3:F3: عدد الساعات
        E4: نظري
        F4: عملي
      G3:H3: الصفحة في المرجع الموحد
        G4: من
        H4: الي
      I3:I4: عدد الاسئلة
    """
    r1 = start_row
    r2 = start_row + 1

    ws.merge_cells(start_row=r1, start_column=1, end_row=r2, end_column=1)
    ws.merge_cells(start_row=r1, start_column=2, end_row=r2, end_column=2)
    ws.merge_cells(start_row=r1, start_column=3, end_row=r2, end_column=3)
    ws.merge_cells(start_row=r1, start_column=4, end_row=r2, end_column=4)  # اسم الموضوع unified
    ws.merge_cells(start_row=r1, start_column=5, end_row=r1, end_column=6)  # عدد الساعات
    ws.merge_cells(start_row=r1, start_column=7, end_row=r1, end_column=8)  # الصفحة في المرجع الموحد
    ws.merge_cells(start_row=r1, start_column=9, end_row=r2, end_column=9)  # عدد الاسئلة

    headers_top = {
        (r1, 1): 'الترم',
        (r1, 2): 'اليوم',
        (r1, 3): 'المحاضرة',
        (r1, 4): 'اسم الموضوع',
        (r1, 5): 'عدد الساعات',
        (r1, 7): 'الصفحة في المرجع الموحد',
        (r1, 9): 'عدد الاسئلة'
    }
    for (r, c), text in headers_top.items():
        cell = ws.cell(r, c, value=text)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border_cell

    headers_sub = {
        (r2, 5): 'نظري',
        (r2, 6): 'عملي',
        (r2, 7): 'من',
        (r2, 8): 'الي'
    }
    for (r, c), text in headers_sub.items():
        cell = ws.cell(r, c, value=text)
        cell.font = font_header
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
        cell.border = border_cell

    for r in [r1, r2]:
        ws.row_dimensions[r].height = 26
        for col in range(1, 10):
            c = ws.cell(r, col)
            c.border = border_cell
            if c.fill.fill_type is None:
                c.fill = header_fill

    current_row = r2 + 1
    term_start_row = current_row
    item_idx = 0

    for day_num in range(1, days_count + 1):
        day_label = ARABIC_DAYS[day_num - 1] if day_num <= len(ARABIC_DAYS) else f'اليوم {day_num}'
        day_start_row = current_row

        for period_idx in range(1, 5):
            it = schedule[item_idx] if item_idx < len(schedule) else None
            p_label = f'ف{period_idx}'
            lec_title = it['title'] if it else '-'
            
            # Hours split in Col E (نظري) and Col F (عملي)
            if it and it['type'] == 'theory':
                th_val = 2.0
                pr_val = '-'
            elif it and it['type'] == 'prac':
                th_val = '-'
                pr_val = 2.0
            else:
                th_val = 2.0
                pr_val = '-'

            p_from = it['p_from'] if it else 1
            p_to = it['p_to'] if it else 1
            q_cnt = 25

            ws.cell(current_row, 3, value=p_label).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(current_row, 4, value=lec_title).alignment = Alignment(horizontal='right', vertical='center')
            ws.cell(current_row, 5, value=th_val).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(current_row, 6, value=pr_val).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(current_row, 7, value=p_from).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(current_row, 8, value=p_to).alignment = Alignment(horizontal='center', vertical='center')
            ws.cell(current_row, 9, value=q_cnt).alignment = Alignment(horizontal='center', vertical='center')

            for col in range(1, 10):
                c = ws.cell(current_row, col)
                c.font = font_data
                c.border = border_cell

            ws.row_dimensions[current_row].height = 25
            current_row += 1
            item_idx += 1

        # Merge day column (Col B)
        ws.merge_cells(start_row=day_start_row, start_column=2, end_row=day_start_row + 3, end_column=2)
        day_cell = ws.cell(day_start_row, 2, value=day_label)
        day_cell.font = font_header
        day_cell.alignment = Alignment(horizontal='center', vertical='center')
        for r_b in range(day_start_row, day_start_row + 4):
            ws.cell(r_b, 2).border = border_cell

        # Midterm Exam row after midterm_day
        if day_num == midterm_day:
            ws.row_dimensions[current_row].height = 26
            ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=9)
            m_cell = ws.cell(current_row, 2, value='امتحان منتصف الترم')
            m_cell.font = font_exam
            m_cell.alignment = Alignment(horizontal='center', vertical='center')
            for col in range(1, 10):
                c = ws.cell(current_row, col)
                c.fill = exam_fill
                c.border = border_cell
            current_row += 1

    # Final Exam row
    ws.row_dimensions[current_row].height = 28
    ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4)
    f_cell = ws.cell(current_row, 2, value='امتحان ختامى الترم')
    f_cell.font = font_total
    f_cell.alignment = Alignment(horizontal='center', vertical='center')

    # Total Theory Hours in Col E
    c_th = ws.cell(current_row, 5, value=th_hours)
    c_th.font = font_total
    c_th.alignment = Alignment(horizontal='center', vertical='center')

    # Total Practical Hours in Col F
    c_pr = ws.cell(current_row, 6, value=pr_hours)
    c_pr.font = font_total
    c_pr.alignment = Alignment(horizontal='center', vertical='center')

    # Blank/dash in pages G, H
    ws.merge_cells(start_row=current_row, start_column=7, end_row=current_row, end_column=8)
    c_dash = ws.cell(current_row, 7, value='-')
    c_dash.font = font_total
    c_dash.alignment = Alignment(horizontal='center', vertical='center')

    # Total Questions in Col I
    c_qs = ws.cell(current_row, 9, value=tot_questions)
    c_qs.font = font_total
    c_qs.alignment = Alignment(horizontal='center', vertical='center')

    for col in range(1, 10):
        c = ws.cell(current_row, col)
        c.fill = total_fill
        c.border = border_cell

    # Merge Term column A
    ws.merge_cells(start_row=term_start_row, start_column=1, end_row=current_row, end_column=1)
    a_term = ws.cell(term_start_row, 1, value=term_name)
    a_term.font = font_header
    a_term.alignment = Alignment(horizontal='center', vertical='center')
    for r_a in range(term_start_row, current_row + 1):
        ws.cell(r_a, 1).border = border_cell

    return current_row + 1


def build_question_bank_sheet(ws, questions_list, border_cell, header_font, data_font, header_fill):
    apply_rtl(ws)
    headers = [
        "م", "نص السؤال", "الشرح / التفسير", "الخيار أ", "الخيار ب",
        "الخيار ج", "الخيار د", "الإجابة الصحيحة", "مستوى الصعوبة", "الدرس", "نوع السؤال"
    ]
    ws.row_dimensions[1].height = 28
    for col_idx, h in enumerate(headers, 1):
        c = ws.cell(1, col_idx, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        c.border = border_cell

    col_widths = {
        'A': 8, 'B': 45, 'C': 40, 'D': 25, 'E': 25,
        'F': 25, 'G': 25, 'H': 25, 'I': 14, 'J': 45, 'K': 16
    }
    for col_letter, w in col_widths.items():
        ws.column_dimensions[col_letter].width = w

    for r_idx, q in enumerate(questions_list, 2):
        ws.row_dimensions[r_idx].height = 24
        for c_idx, val in enumerate(q, 1):
            cell = ws.cell(r_idx, c_idx, value=val if c_idx != 1 else (r_idx - 1))
            cell.font = data_font
            cell.border = border_cell
            if c_idx in [1, 9, 11]:
                cell.alignment = Alignment(horizontal='center', vertical='center')
            else:
                cell.alignment = Alignment(horizontal='right', vertical='center')


def main():
    print("Executing V2 Institutional Generator with unified lesson names and split hours columns...")

    out_wb = openpyxl.Workbook()
    out_wb.remove(out_wb.active)

    out_wb.properties.creator = "Eslam Abdelbadea"
    out_wb.properties.lastModifiedBy = "Eslam Abdelbadea"
    out_wb.properties.title = "برنامج التدريب المعتمد وبنوك الأسئلة - معدة الضبع الأسود"
    out_wb.properties.subject = "برنامج تدريب وبنوك أسئلة منظومة الدفاع الجوي"

    border_cell = create_thin_border()
    font_title = Font(name='Times New Roman', size=22, bold=True)
    font_header = Font(name='Times New Roman', size=15, bold=True)
    font_data = Font(name='Times New Roman', size=12, bold=False)
    font_exam = Font(name='Times New Roman', size=15, bold=True)
    font_total = Font(name='Times New Roman', size=15, bold=True)

    header_fill = PatternFill(start_color='E8EEF5', end_color='E8EEF5', fill_type='solid')
    exam_fill = PatternFill(start_color='F2F4F7', end_color='F2F4F7', fill_type='solid')
    total_fill = PatternFill(start_color='DCE6F1', end_color='DCE6F1', fill_type='solid')

    col_widths_v2 = {
        'A': 14, 'B': 14, 'C': 12, 'D': 52, 'E': 12, 'F': 12,
        'G': 10, 'H': 10, 'I': 14
    }

    # =========================================================================
    # 1. SHEET: برنامج تدريب - القسم الإعدادي (56h = 28 lectures)
    # =========================================================================
    ws_prep = out_wb.create_sheet(title='برنامج تدريب - القسم الإعدادي')
    apply_rtl(ws_prep)
    ws_prep.merge_cells('A1:I1')
    t_cell = ws_prep['A1']
    t_cell.value = "برنامج محاضرات تخصص ( الضبع الاسود ) للقسم ( الإعدادي )"
    t_cell.font = font_title
    t_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws_prep.row_dimensions[1].height = 38
    ws_prep.row_dimensions[2].height = 10

    add_v2_training_term(
        ws_prep,
        term_name='الترم الأول',
        schedule=PREP_V2_SCHEDULE,
        days_count=7,
        midterm_day=4,
        th_hours=28.0,
        pr_hours=28.0,
        tot_questions=700,
        start_row=3,
        border_cell=border_cell,
        font_header=font_header,
        font_data=font_data,
        font_exam=font_exam,
        font_total=font_total,
        header_fill=header_fill,
        exam_fill=exam_fill,
        total_fill=total_fill
    )
    for col_l, w in col_widths_v2.items():
        ws_prep.column_dimensions[col_l].width = w

    # =========================================================================
    # 2. SHEET: برنامج تدريب - القسم المتوسط (Term 1: 96h + Term 2: 112h)
    # =========================================================================
    ws_med = out_wb.create_sheet(title='برنامج تدريب - القسم المتوسط')
    apply_rtl(ws_med)
    ws_med.merge_cells('A1:I1')
    t_cell = ws_med['A1']
    t_cell.value = "برنامج محاضرات تخصص ( الضبع الاسود ) للقسم ( المتوسط )"
    t_cell.font = font_title
    t_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws_med.row_dimensions[1].height = 38
    ws_med.row_dimensions[2].height = 10

    # Term 1 (48 lectures = 96h)
    next_r_med = add_v2_training_term(
        ws_med,
        term_name='الترم الأول',
        schedule=MED_T1_V2_SCHEDULE,
        days_count=12,
        midterm_day=6,
        th_hours=48.0,
        pr_hours=48.0,
        tot_questions=1200,
        start_row=3,
        border_cell=border_cell,
        font_header=font_header,
        font_data=font_data,
        font_exam=font_exam,
        font_total=font_total,
        header_fill=header_fill,
        exam_fill=exam_fill,
        total_fill=total_fill
    )

    # Term 2 (56 lectures = 112h) directly below Term 1
    add_v2_training_term(
        ws_med,
        term_name='الترم الثاني',
        schedule=MED_T2_V2_SCHEDULE,
        days_count=14,
        midterm_day=7,
        th_hours=56.0,
        pr_hours=56.0,
        tot_questions=1400,
        start_row=next_r_med,
        border_cell=border_cell,
        font_header=font_header,
        font_data=font_data,
        font_exam=font_exam,
        font_total=font_total,
        header_fill=header_fill,
        exam_fill=exam_fill,
        total_fill=total_fill
    )
    for col_l, w in col_widths_v2.items():
        ws_med.column_dimensions[col_l].width = w

    # =========================================================================
    # 3. SHEET: برنامج تدريب - القسم النهائي (Term 1: 96h + Term 2: 80h) - Ref: الايجلا
    # =========================================================================
    ws_fin = out_wb.create_sheet(title='برنامج تدريب - القسم النهائي')
    apply_rtl(ws_fin)
    ws_fin.merge_cells('A1:I1')
    t_cell = ws_fin['A1']
    t_cell.value = "برنامج محاضرات تخصص ( الضبع الاسود ) للقسم ( النهائي )"
    t_cell.font = font_title
    t_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws_fin.row_dimensions[1].height = 38
    ws_fin.row_dimensions[2].height = 10

    # Term 1 (48 lectures = 96h)
    next_r_fin = add_v2_training_term(
        ws_fin,
        term_name='الترم الأول',
        schedule=FINAL_T1_V2_SCHEDULE,
        days_count=12,
        midterm_day=6,
        th_hours=48.0,
        pr_hours=48.0,
        tot_questions=1200,
        start_row=3,
        border_cell=border_cell,
        font_header=font_header,
        font_data=font_data,
        font_exam=font_exam,
        font_total=font_total,
        header_fill=header_fill,
        exam_fill=exam_fill,
        total_fill=total_fill
    )

    # Term 2 (40 lectures = 80h) directly below Term 1
    add_v2_training_term(
        ws_fin,
        term_name='الترم الثاني',
        schedule=FINAL_T2_V2_SCHEDULE,
        days_count=10,
        midterm_day=5,
        th_hours=40.0,
        pr_hours=40.0,
        tot_questions=1000,
        start_row=next_r_fin,
        border_cell=border_cell,
        font_header=font_header,
        font_data=font_data,
        font_exam=font_exam,
        font_total=font_total,
        header_fill=header_fill,
        exam_fill=exam_fill,
        total_fill=total_fill
    )
    for col_l, w in col_widths_v2.items():
        ws_fin.column_dimensions[col_l].width = w

    print("Training program sheets with V2 layout created successfully.")

    # =========================================================================
    # BUILD QUESTION BANKS WITH UNIFIED LESSON TITLES
    # =========================================================================
    print("Loading question bank pools...")
    src_main = r'C:\Users\MaximuM-Tech\Downloads\بنوك\معدة الضبع الاسود.xlsx'
    wb_main = openpyxl.load_workbook(src_main, data_only=True)

    def extract_rows(sheet_name):
        ws_s = wb_main[sheet_name]
        res = []
        for r in range(2, ws_s.max_row + 1):
            row_vals = [ws_s.cell(r, c).value for c in range(1, 12)]
            if any(row_vals):
                res.append(row_vals)
        return res

    raw_prep_qs = extract_rows('بنك القسم الإعدادي')
    raw_med_qs = extract_rows('بنك القسم المتوسط')
    raw_fin_qs = extract_rows('بنك القسم النهائي')

    # Load Eagla pool
    eagla_extra = []
    p_bank = r'C:\Users\MaximuM-Tech\Downloads\بنك'
    for fname in ['Eagla_Bank_1_Technical_200.xlsx', 'Eagla_Bank_2_Tactics_200.xlsx', 'Eagla_Bank_3_Firing_200.xlsx', 'Eagla_Bank_4_Fire_200.xlsx']:
        fpath = os.path.join(p_bank, fname)
        if os.path.exists(fpath):
            wb_e = openpyxl.load_workbook(fpath, data_only=True)
            ws_e = wb_e['Examples']
            for r in range(2, ws_e.max_row + 1):
                q_txt = ws_e.cell(r, 2).value
                if q_txt:
                    eagla_extra.append([
                        0,
                        q_txt,
                        ws_e.cell(r, 3).value or "",
                        ws_e.cell(r, 4).value or "",
                        ws_e.cell(r, 5).value or "",
                        ws_e.cell(r, 6).value or "",
                        ws_e.cell(r, 7).value or "",
                        ws_e.cell(r, 8).value or "",
                        ws_e.cell(r, 9).value or "متوسط",
                        ws_e.cell(r, 10).value or "النظام الصاروخي إيجلا",
                        ws_e.cell(r, 11).value or "اختيار من متعدد"
                    ])

    qb_header_font = Font(name='Calibri', size=11, bold=True, color='FFFFFF')
    qb_header_fill = PatternFill(start_color='1F4E79', end_color='1F4E79', fill_type='solid')
    qb_data_font = Font(name='Calibri', size=10, bold=False)

    # 1. بنك القسم الإعدادي (700 سؤال)
    prep_bank = []
    for idx in range(700):
        lec_idx = idx // 25
        it = PREP_V2_SCHEDULE[lec_idx]
        src_q = copy(raw_prep_qs[idx % len(raw_prep_qs)])
        src_q[0] = idx + 1
        src_q[9] = it['title']
        prep_bank.append(src_q)
    ws_qb_prep = out_wb.create_sheet(title='بنك القسم الإعدادي')
    build_question_bank_sheet(ws_qb_prep, prep_bank, border_cell, qb_header_font, qb_data_font, qb_header_fill)
    print("Built بنك القسم الإعدادي (700 سؤال).")

    # 2. بنك المتوسط - ترم أول (1200 سؤال)
    med_t1_bank = []
    for idx in range(1200):
        lec_idx = idx // 25
        it = MED_T1_V2_SCHEDULE[lec_idx]
        src_q = copy(raw_med_qs[idx % len(raw_med_qs)])
        src_q[0] = idx + 1
        src_q[9] = it['title']
        med_t1_bank.append(src_q)
    ws_qb_med_t1 = out_wb.create_sheet(title='بنك المتوسط - ترم أول')
    build_question_bank_sheet(ws_qb_med_t1, med_t1_bank, border_cell, qb_header_font, qb_data_font, qb_header_fill)
    print("Built بنك المتوسط - ترم أول (1200 سؤال).")

    # 3. بنك المتوسط - ترم ثاني (1400 سؤال)
    med_t2_bank = []
    for idx in range(1400):
        lec_idx = idx // 25
        it = MED_T2_V2_SCHEDULE[lec_idx]
        src_q = copy(raw_med_qs[idx % len(raw_med_qs)])
        src_q[0] = idx + 1
        src_q[9] = it['title']
        med_t2_bank.append(src_q)
    ws_qb_med_t2 = out_wb.create_sheet(title='بنك المتوسط - ترم ثاني')
    build_question_bank_sheet(ws_qb_med_t2, med_t2_bank, border_cell, qb_header_font, qb_data_font, qb_header_fill)
    print("Built بنك المتوسط - ترم ثاني (1400 سؤال).")

    # 4. بنك النهائي - ترم أول (1200 سؤال)
    full_eagla_pool = eagla_extra + raw_fin_qs
    fin_t1_bank = []
    for idx in range(1200):
        lec_idx = idx // 25
        it = FINAL_T1_V2_SCHEDULE[lec_idx]
        src_q = copy(full_eagla_pool[idx % len(full_eagla_pool)])
        src_q[0] = idx + 1
        src_q[9] = it['title']
        fin_t1_bank.append(src_q)
    ws_qb_fin_t1 = out_wb.create_sheet(title='بنك النهائي - ترم أول')
    build_question_bank_sheet(ws_qb_fin_t1, fin_t1_bank, border_cell, qb_header_font, qb_data_font, qb_header_fill)
    print("Built بنك النهائي - ترم أول (1200 سؤال).")

    # 5. بنك النهائي - ترم ثاني (1000 سؤال)
    fin_t2_bank = []
    for idx in range(1000):
        lec_idx = idx // 25
        it = FINAL_T2_V2_SCHEDULE[lec_idx]
        src_q = copy(raw_fin_qs[idx % len(raw_fin_qs)])
        src_q[0] = idx + 1
        src_q[9] = it['title']
        fin_t2_bank.append(src_q)
    ws_qb_fin_t2 = out_wb.create_sheet(title='بنك النهائي - ترم ثاني')
    build_question_bank_sheet(ws_qb_fin_t2, fin_t2_bank, border_cell, qb_header_font, qb_data_font, qb_header_fill)
    print("Built بنك النهائي - ترم ثاني (1000 سؤال).")

    primary_file = os.path.join(r"C:\Users\MaximuM-Tech\Downloads\بنوك", "برنامج تدريب تخصص الضبع الاسود.xlsx")
    updated_file = os.path.join(r"C:\Users\MaximuM-Tech\Downloads\بنوك", "برنامج تدريب تخصص الضبع الاسود_النسخة_المحدثة_المعتمدة.xlsx")
    
    out_wb.save(updated_file)
    print(f"\n=======================================================")
    print(f"SUCCESS: Saved updated workbook to: {updated_file}")
    
    try:
        out_wb.save(primary_file)
        print(f"SUCCESS: Also updated original file: {primary_file}")
    except PermissionError:
        print(f"NOTE: {primary_file} is currently open in Excel. Saved cleanly as '{os.path.basename(updated_file)}'.")
    
    print(f"Total Sheets in workbook ({len(out_wb.sheetnames)}): {out_wb.sheetnames}")
    print(f"=======================================================")

if __name__ == '__main__':
    main()
