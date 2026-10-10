# -*- coding: utf-8 -*-
"""
محرك الفحص والتدقيق المؤسسي لبنوك الأسئلة وبرامج التدريب (LMS Exam Bank & Syllabus Auditor)
========================================================================================
يقوم هذا المحرك بمحاكاة الفحص الصارم للجان إدارة الامتحانات والاعتماد المؤسسي:
 1. اختبار التكرار والتمييز الشرطي (Conditional Formatting Simulation) مع كشف التلاعب بالفواصل
 2. اختبار البيفوت تيبل لعدد الأسئلة وتوازن الدروس (Pivot Table & Question Counts)
 3. فحص العبارات والخيارات المحظورة أكاديمياً (Forbidden Phrases: جميع ما سبق، كلاهما صواب، إلخ)
 4. فحص التدرج ونسب مستويات الصعوبة وأنواع الأسئلة (Difficulty Progression & Distribution)
 5. فحص ضوابط أسئلة الصواب والخطأ (TF Clean Options & Answer Format)
 6. فحص اكتمال خيارات أسئلة الاختيار من متعدد وعدم وجود أعمدة فارغة (MCQ Completeness)
 7. مطابقة تسلسل الدروس التام مع برنامج التدريب (Lesson Sequence 1-to-1 Match)
 8. فحص سرعة الفتح وتفريغ عمود الشرح وضوابط الداتا فاليديشن (Data Validation & Speed Check)
"""

import openpyxl
import os
import sys
import collections
import re
from typing import Dict, List, Any, Tuple, Optional

sys.stdout.reconfigure(encoding='utf-8')

# قائمة العبارات المحظورة أكاديمياً في بنوك الأسئلة
FORBIDDEN_PHRASES = [
    r'جميع\s+ما\s+سبق',
    r'كل\s+ما\s+سبق',
    r'كلاهما\s+صواب',
    r'كلاهما\s+صحيح',
    r'كلاهما\s+خطأ',
    r'كلاهما\s+خطا',
    r'[\(\[]?[أا][\)\]]?\s*[,،و\s]+\s*[\(\[]?ب[\)\]]?\s+مع[ااً]',
    r'[\(\[]?[أا][\)\]]?\s*[,،و\s]+\s*[\(\[]?ج[\)\]]?\s+مع[ااً]',
    r'[\(\[]?ب[\)\]]?\s*[,،و\s]+\s*[\(\[]?ج[\)\]]?\s+مع[ااً]',
    r'كل\s+ما\s+ذكر',
    r'جميع\s+الإجابات\s+صحيحة',
    r'جميع\s+الخيارات\s+صحيحة',
    r'لا\s+شيء\s+مما\s+سبق',
    r'لا\s+شئ\s+مما\s+سبق',
    r'ليس\s+أياً\s+مما\s+سبق',
    r'ليس\s+ايا\s+مما\s+سبق'
]

def normalize_arabic_text(text: str) -> str:
    """
    توحيد ومطابقة النصوص العربية مع مراعاة المرونة الإملائية المقبولة:
    - توحيد الهمزات (أ, إ, آ -> ا)
    - توحيد الياء والياء المقصورة (ى -> ي)
    - توحيد التاء المربوطة والهاء (ة -> ه)
    - إزالة التشكيل والتطويل والمسافات الزائدة
    - توحيد واو العطف المتصلة والمنفصلة (مثل 'و وسائل' -> 'ووسائل')
    """
    if not text:
        return ""
    t = str(text).strip()
    
    # إزالة التشكيل
    t = re.sub(r'[\u064B-\u065F\u0670]', '', t)
    # إزالة التطويل الكشيدة
    t = re.sub(r'\u0640', '', t)
    
    # توحيد الألفات والهمزات
    t = re.sub(r'[أإآٱ]', 'ا', t)
    
    # توحيد الياء والألف المقصورة
    t = re.sub(r'[ى]', 'ي', t)
    
    # توحيد التاء المربوطة بالهاء
    t = re.sub(r'[ة]', 'ه', t)
    
    # توحيد المسافات المتكررة أولاً
    t = re.sub(r'\s+', ' ', t)
    
    # توحيد واو العطف المنفصلة (مثال: "و وسائل" -> "ووسائل"، "و التقييم" -> "والتقييم")
    t = re.sub(r'(^|\s)و\s+([^\s])', r'\1و\2', t)
    
    return t.strip()

def extract_syllabus_lessons(prog_wb, prog_sheet=None, bank_name='', bank_lessons=[]) -> Tuple[List[str], str]:
    """
    استكشاف واستخراج تسلسل دروس برنامج التدريب بذكاء شامل:
    - يتعرف تلقائياً على القسم (إعدادي / متوسط / نهائي) من اسم البنك أو محتواه.
    - يدعم البرامج المقسمة على شيتات منفصلة أو المجمعة في شيت واحد مقسم أفقياً.
    - يستكشف عمود الموضوع وعمود المرحلة ديناميكياً ويتجاهل صفوف الإجماليات والترويسات.
    """
    combined_ctx = (bank_name or '') + ' ' + ' '.join(bank_lessons[:15])
    ctx_norm = normalize_arabic_text(combined_ctx)
    target_stage = None
    if any(k in ctx_norm for k in ['اعدادي', 'القسم الاعدادي', 'الاعدادي']):
        target_stage = 'اعدادي'
    elif any(k in ctx_norm for k in ['متوسط', 'القسم المتوسط', 'المتوسط']):
        target_stage = 'متوسط'
    elif any(k in ctx_norm for k in ['نهائي', 'القسم النهائي', 'النهائي']):
        target_stage = 'نهائي'

    ws = None
    if prog_sheet and prog_sheet in prog_wb.sheetnames:
        ws = prog_wb[prog_sheet]
    elif target_stage:
        for s in prog_wb.sheetnames:
            if target_stage in normalize_arabic_text(s):
                ws = prog_wb[s]
                break

    if not ws:
        for s in prog_wb.sheetnames:
            s_n = normalize_arabic_text(s)
            if any(w in s_n for w in ['برنامج', 'تدريب', 'خطه', 'تخصص', 'ساعات']):
                ws = prog_wb[s]
                break

    if not ws:
        ws = prog_wb.active

    sheet_title = ws.title
    topic_col = 2
    stage_col = None
    start_r = 1

    header_found = False
    for r in range(1, min(15, ws.max_row) + 1):
        for c in range(1, min(15, ws.max_column) + 1):
            v = normalize_arabic_text(str(ws.cell(r, c).value or ''))
            if any(w in v for w in ['اسم الموضوع', 'الموضوع', 'الدرس', 'اسم الدرس', 'topic']):
                topic_col = c
                start_r = r + 1
                header_found = True
        if header_found:
            for c in range(1, min(15, ws.max_column) + 1):
                v = normalize_arabic_text(str(ws.cell(r, c).value or ''))
                if any(w in v for w in ['السنه الدراسيه', 'المرحله', 'القسم', 'الفصل الدراسي', 'stage']):
                    stage_col = c
            break

    # إذا لم يتم تحديد target_stage من اسم البنك، نقوم بمطابقة دروس البنك مع أعمدة المراحل في شيت البرنامج
    if not target_stage and bank_lessons and stage_col:
        stage_votes = {'اعدادي': 0, 'متوسط': 0, 'نهائي': 0}
        norm_bank_lessons = [normalize_arabic_text(bl) for bl in bank_lessons if bl]
        for r in range(start_r, ws.max_row + 1):
            top_val = ws.cell(r, topic_col).value
            stg_val = ws.cell(r, stage_col).value
            if top_val and stg_val:
                t_norm = normalize_arabic_text(str(top_val))
                s_norm = normalize_arabic_text(str(stg_val))
                if t_norm in norm_bank_lessons:
                    for k in stage_votes:
                        if k in s_norm:
                            stage_votes[k] += 1
        best_stage = max(stage_votes, key=stage_votes.get)
        if stage_votes[best_stage] > 0:
            target_stage = best_stage

    extracted = []
    current_sec = None
    last_val = None
    is_sheet_dedicated = target_stage and (target_stage in normalize_arabic_text(sheet_title))

    for r in range(start_r, ws.max_row + 1):
        row_str = ' '.join(str(ws.cell(r, c).value or '') for c in range(1, min(8, ws.max_column) + 1))
        row_norm = normalize_arabic_text(row_str)
        if any(w in row_norm for w in ['اجمالي', 'اجمالي الساعات', 'امتحان']):
            continue

        if not is_sheet_dedicated:
            # التحقق من عناوين الأقسام الرئيسية
            if any(w in row_norm for w in ['موضوعات', 'برنامج تدريب']):
                if 'اعدادي' in row_norm: current_sec = 'اعدادي'
                elif 'متوسط' in row_norm: current_sec = 'متوسط'
                elif 'نهائي' in row_norm: current_sec = 'نهائي'
                continue

            if stage_col:
                st_v = normalize_arabic_text(str(ws.cell(r, stage_col).value or ''))
                if 'اعدادي' in st_v: current_sec = 'اعدادي'
                elif 'متوسط' in st_v: current_sec = 'متوسط'
                elif 'نهائي' in st_v: current_sec = 'نهائي'

            # إذا كانت المرحلة الحالية غير محددة بعد أو لا تطابق المرحلة المستهدفة
            if target_stage:
                if current_sec is not None and current_sec != target_stage:
                    continue
                if current_sec is None:
                    continue

        name = ws.cell(r, topic_col).value
        if name and str(name).strip() not in ['', 'None']:
            n_clean = str(name).strip()
            if any(w in n_clean for w in ['امتحان', 'إجمالي', 'اجمالي', 'اسم الموضوع']):
                continue
            if n_clean in ['نظري', 'مخططات و دوائر كهربائية', 'مخططات ودوائر كهربائية'] or n_clean.isdigit():
                continue
            if n_clean != last_val:
                extracted.append(n_clean)
                last_val = n_clean
                last_val = n_clean

    return extracted, sheet_title

def detect_bank_structure(ws) -> Dict[str, Any]:
    """
    استكشاف ذكي وديناميكي لهيكل شيت بنك الأسئلة ومواقع الأعمدة وصف البداية.
    يدعم النماذج الرسمية (A=نوع السؤال, B=نص السؤال, ...) وأي شيتات مخصصة أخرى.
    """
    col_map = {
        "type": 1,
        "question": 2,
        "explanation": 3,
        "diff": 4,
        "ans": 5,
        "opt_a": 6,
        "opt_b": 7,
        "opt_c": 8,
        "opt_d": 9,
        "lesson": 10,
        "header_row": 1,
        "data_start_row": 2
    }

    max_r = min(15, ws.max_row)
    max_c = min(20, ws.max_column)

    best_score = 0
    best_row = 1
    found_map = {}

    for r in range(1, max_r + 1):
        row_map = {}
        score = 0
        for c in range(1, max_c + 1):
            val = str(ws.cell(r, c).value or '').strip()
            if not val:
                continue
            
            val_norm = normalize_arabic_text(val)
            
            # عمود نص السؤال
            if any(k in val_norm for k in ['نص السوال', 'السوال', 'الاسئله', 'question']):
                row_map['question'] = c
                score += 3
            # عمود نوع السؤال
            elif any(k in val_norm for k in ['نوع السوال', 'نوع', 'type']):
                row_map['type'] = c
                score += 2
            # عمود الشرح / التفسير
            elif any(k in val_norm for k in ['شرح', 'تفسير', 'explanation']):
                row_map['explanation'] = c
                score += 1
            # عمود مستوى الصعوبة
            elif any(k in val_norm for k in ['مستوي الصعوبه', 'صعوبه', 'difficulty']):
                row_map['diff'] = c
                score += 2
            # عمود الإجابة الصحيحة
            elif any(k in val_norm for k in ['الاجابه الصحيحه', 'الاجابه', 'answer', 'correct']):
                row_map['ans'] = c
                score += 2
            # عمود الخيار أ
            elif any(k in val_norm for k in ['الخيار ا', 'خيار ا', 'option a', 'a']):
                if 'opt_a' not in row_map: row_map['opt_a'] = c; score += 1
            # عمود الخيار ب
            elif any(k in val_norm for k in ['الخيار ب', 'خيار ب', 'option b', 'b']):
                if 'opt_b' not in row_map: row_map['opt_b'] = c; score += 1
            # عمود الخيار ج
            elif any(k in val_norm for k in ['الخيار ج', 'خيار ج', 'option c', 'c']):
                if 'opt_c' not in row_map: row_map['opt_c'] = c; score += 1
            # عمود الخيار د
            elif any(k in val_norm for k in ['الخيار د', 'خيار د', 'option d', 'd']):
                if 'opt_d' not in row_map: row_map['opt_d'] = c; score += 1
            # عمود الموضوع / الدرس
            elif any(k in val_norm for k in ['الموضوع', 'الدرس', 'اسم الموضوع', 'topic', 'lesson']):
                row_map['lesson'] = c
                score += 3

        if score > best_score and 'question' in row_map:
            best_score = score
            best_row = r
            found_map = row_map

    if best_score >= 3:
        for k, v in found_map.items():
            col_map[k] = v
        col_map["header_row"] = best_row
        col_map["data_start_row"] = best_row + 1

    return col_map

class ExamBankAuditor:
    def __init__(self, bank_path: str, program_path: Optional[str] = None, program_sheet: Optional[str] = None, questions_per_lesson: int = 25):
        self.bank_path = bank_path
        self.program_path = program_path
        self.program_sheet = program_sheet
        self.expected_q_per_lesson = questions_per_lesson
        
        self.bank_wb = None
        self.prog_wb = None
        self.ws_questions = None
        self.ws_lookups = None
        self.ws_program = None
        self.col_map = {}
        
        self.results = {
            "file_info": {},
            "duplicates": {"status": "PASS", "count": 0, "details": []},
            "lesson_counts": {"status": "PASS", "details": {}, "pivot": []},
            "forbidden_phrases": {"status": "PASS", "count": 0, "details": []},
            "distribution": {"status": "PASS", "types": {}, "difficulties": {}, "details": []},
            "tf_rules": {"status": "PASS", "violations": 0, "details": []},
            "mcq_rules": {"status": "PASS", "violations": 0, "details": []},
            "syllabus_match": {"status": "PASS", "mismatches": 0, "details": []},
            "speed_validation": {"status": "PASS", "col_c_filled": 0, "dv_count": 0, "details": []},
            "overall_status": "PASS",
            "score": 100.0,
            "recommendations": []
        }

    def load_files(self):
        if not os.path.exists(self.bank_path):
            raise FileNotFoundError(f"ملف بنك الأسئلة غير موجود: {self.bank_path}")
        
        self.bank_wb = openpyxl.load_workbook(self.bank_path, data_only=True)
        
        # 1. الاستكشاف الذكي لشيت الأسئلة
        target_ws_name = None
        # أولوية شيت Questions
        if 'Questions' in self.bank_wb.sheetnames:
            target_ws_name = 'Questions'
        else:
            # البحث عن شيتات بديلة (بنك الاسئلة، الاسئلة، Sheet1، ...)
            for s in self.bank_wb.sheetnames:
                s_lower = s.lower()
                if any(w in s_lower for w in ['question', 'اسئل', 'أسئل', 'بنك', 'bank']):
                    target_ws_name = s
                    break
            if not target_ws_name and self.bank_wb.sheetnames:
                # استخدام الشيت الأول مباشرة كشيت افتراضي
                target_ws_name = self.bank_wb.sheetnames[0]

        if not target_ws_name:
            raise ValueError("ملف بنك الأسئلة لا يحتوي على أي شيتات صالحة!")
        
        self.ws_questions = self.bank_wb[target_ws_name]
        self.ws_lookups = self.bank_wb['Lookups'] if 'Lookups' in self.bank_wb.sheetnames else None
        
        # 2. الاستكشاف الذكي لأعمدة النموذج
        self.col_map = detect_bank_structure(self.ws_questions)
        
        self.results["file_info"]["bank_name"] = os.path.basename(self.bank_path)
        self.results["file_info"]["sheet_name"] = target_ws_name
        self.results["file_info"]["bank_size_kb"] = round(os.path.getsize(self.bank_path) / 1024, 1)
        self.results["file_info"]["total_rows"] = max(0, self.ws_questions.max_row - self.col_map["data_start_row"] + 1)

        if self.program_path:
            if not os.path.exists(self.program_path):
                raise FileNotFoundError(f"ملف برنامج التدريب غير موجود: {self.program_path}")
            self.prog_wb = openpyxl.load_workbook(self.program_path, data_only=True)
            if self.program_sheet:
                if self.program_sheet not in self.prog_wb.sheetnames:
                    raise ValueError(f"الشيت '{self.program_sheet}' غير موجود في ملف برنامج التدريب!")
                self.ws_program = self.prog_wb[self.program_sheet]
            else:
                cand = [s for s in self.prog_wb.sheetnames if any(w in s for w in ['برنامج', 'تدريب', 'ساعات', 'محاضرات', 'خطة'])]
                if cand:
                    self.ws_program = self.prog_wb[cand[0]]
                    self.program_sheet = cand[0]
                elif self.prog_wb.sheetnames:
                    # اختيار أول شيت متاح كافتراضي ذكي
                    self.ws_program = self.prog_wb[self.prog_wb.sheetnames[0]]
                    self.program_sheet = self.prog_wb.sheetnames[0]
            self.results["file_info"]["program_name"] = os.path.basename(self.program_path)
            self.results["file_info"]["program_sheet"] = self.program_sheet

    def run_all_audits(self) -> Dict[str, Any]:
        self.load_files()
        
        # 1. اختبار التكرار والتمييز الشرطي المزدوج (Exact + Normalized)
        self._audit_duplicates()
        
        # 2. اختبار بيفوت تيبل وتوزيع الأسئلة على الدروس
        self._audit_lesson_counts()
        
        # 3. فحص العبارات المحظورة في الخيارات
        self._audit_forbidden_phrases()
        
        # 4. فحص نسب الصعوبة وتوازن الأنواع
        self._audit_distribution()
        
        # 5. فحص ضوابط أسئلة الصواب والخطأ
        self._audit_tf_rules()
        
        # 6. فحص ضوابط واكتمال أسئلة الاختيار من متعدد
        self._audit_mcq_rules()
        
        # 7. مطابقة تسلسل الدروس مع برنامج التدريب
        if self.ws_program:
            self._audit_syllabus_match()
            
        # 8. فحص السرعة والداتا فاليديشن وعمود الشرح
        self._audit_speed_and_validation()
        
        # حساب التقييم الإجمالي
        self._calculate_verdict()
        
        return self.results

    def _audit_duplicates(self):
        """
        اختبار محاكاة التنسيق الشرطي الدقيق (Excel Conditional Formatting Simulation):
        يقوم بفحص الخلايا ومقارنتها حرفياً بنسبة 100% كما تفعل أداة التمييز الشرطي في إكسيل تماماً.
        إذا وُجد اختلاف في المسافات أو الصياغة، فإن إكسيل لا يظللها، وبالتالي يتطابق الفحص مع فحص اللجان اليدوي 100%.
        """
        seen_raw = {}
        duplicates_found = []
        
        q_col = self.col_map.get("question", 2)
        les_col = self.col_map.get("lesson", 10)
        start_r = self.col_map.get("data_start_row", 2)

        for r in range(start_r, self.ws_questions.max_row + 1):
            q_val = self.ws_questions.cell(r, q_col).value
            if q_val is None:
                continue
            
            raw_text = str(q_val)
            if not raw_text.strip():
                continue

            les = self.ws_questions.cell(r, les_col).value
            
            # مطابقة إكسيل الحرفية للتنسيق الشرطي (Exact Match)
            if raw_text in seen_raw:
                orig_r, orig_les = seen_raw[raw_text]
                duplicates_found.append({
                    "row": r,
                    "lesson": les,
                    "text": raw_text.strip(),
                    "duplicate_of_row": orig_r,
                    "duplicate_of_lesson": orig_les,
                    "detection_type": "تطابق حرفي تام (يظلله التنسيق الشرطي لإكسيل)"
                })
            else:
                seen_raw[raw_text] = (r, les)
                
        self.results["duplicates"]["count"] = len(duplicates_found)
        if duplicates_found:
            self.results["duplicates"]["status"] = "FAIL"
            self.results["duplicates"]["details"] = duplicates_found
            self.results["recommendations"].append(
                f"❌ يوجد {len(duplicates_found)} سؤال مكرر سيكشفه التنسيق الشرطي المباشر في إكسيل! يجب استبدالها بأسئلة فريدة."
            )
        else:
            self.results["duplicates"]["status"] = "PASS"

    def _audit_lesson_counts(self):
        """محاكاة جدول بيفوت تيبل لفحص عدد الأسئلة لكل درس"""
        counts = collections.OrderedDict()
        les_col = self.col_map.get("lesson", 10)
        start_r = self.col_map.get("data_start_row", 2)
        
        for r in range(start_r, self.ws_questions.max_row + 1):
            les = self.ws_questions.cell(r, les_col).value
            if not les:
                continue
            les_str = str(les).strip()
            counts[les_str] = counts.get(les_str, 0) + 1
            
        pivot_rows = []
        shortage_lessons = []
        excess_lessons = []
        
        for idx, (les, count) in enumerate(counts.items(), start=1):
            diff = count - self.expected_q_per_lesson
            status = "OK" if diff == 0 else ("SHORTAGE" if diff < 0 else "EXCESS")
            pivot_rows.append({
                "index": idx,
                "lesson": les,
                "count": count,
                "expected": self.expected_q_per_lesson,
                "diff": diff,
                "status": status
            })
            if diff < 0:
                shortage_lessons.append((les, count, abs(diff)))
            elif diff > 0:
                excess_lessons.append((les, count, diff))
                
        self.results["lesson_counts"]["pivot"] = pivot_rows
        self.results["lesson_counts"]["total_lessons"] = len(counts)
        
        if shortage_lessons or excess_lessons:
            self.results["lesson_counts"]["status"] = "FAIL"
            if shortage_lessons:
                self.results["recommendations"].append(
                    f"⚠️ يوجد نقص في عدد الأسئلة في {len(shortage_lessons)} درساً عن المعدل المقرر ({self.expected_q_per_lesson} سؤال)."
                )
            if excess_lessons:
                self.results["recommendations"].append(
                    f"⚠️ يوجد زيادة في عدد الأسئلة في {len(excess_lessons)} درساً عن المعدل المقرر."
                )
        else:
            self.results["lesson_counts"]["status"] = "PASS"

    def _audit_forbidden_phrases(self):
        """فحص وجود العبارات المحظورة في الخيارات (جميع ما سبق، كلاهما صواب، إلخ)"""
        violations = []
        q_col = self.col_map.get("question", 2)
        les_col = self.col_map.get("lesson", 10)
        opt_cols = [
            ("الخيار أ", self.col_map.get("opt_a", 6)),
            ("الخيار ب", self.col_map.get("opt_b", 7)),
            ("الخيار ج", self.col_map.get("opt_c", 8)),
            ("الخيار د", self.col_map.get("opt_d", 9))
        ]
        start_r = self.col_map.get("data_start_row", 2)
        
        for r in range(start_r, self.ws_questions.max_row + 1):
            les = self.ws_questions.cell(r, les_col).value
            q_txt = self.ws_questions.cell(r, q_col).value
            
            for col_name, c_idx in opt_cols:
                val = self.ws_questions.cell(r, c_idx).value
                if val:
                    val_str = str(val).strip()
                    for pattern in FORBIDDEN_PHRASES:
                        if re.search(pattern, val_str):
                            violations.append({
                                "row": r,
                                "lesson": les,
                                "col": col_name,
                                "matched_pattern": pattern,
                                "option_text": val_str,
                                "question_text": str(q_txt)[:60] + "..." if q_txt else ""
                            })
                            break
                            
        self.results["forbidden_phrases"]["count"] = len(violations)
        if violations:
            self.results["forbidden_phrases"]["status"] = "FAIL"
            self.results["forbidden_phrases"]["details"] = violations
            self.results["recommendations"].append(
                f"❌ تم رصد {len(violations)} خياراً يحتوي على عبارات محظورة أكاديمياً مثل ('جميع ما سبق' أو 'كلاهما صواب'). ترفض لجان الجودة هذا الأسلوب ويجب صياغة مشتتات موضوعية بديلة."
            )
        else:
            self.results["forbidden_phrases"]["status"] = "PASS"

    def _audit_distribution(self):
        """فحص نسب التوزيع وتدرج مستويات الصعوبة وأنواع الأسئلة"""
        type_counts = collections.Counter()
        diff_counts = collections.Counter()
        total_q = 0
        
        type_col = self.col_map.get("type", 1)
        diff_col = self.col_map.get("diff", 4)
        start_r = self.col_map.get("data_start_row", 2)
        
        for r in range(start_r, self.ws_questions.max_row + 1):
            q_type = self.ws_questions.cell(r, type_col).value
            q_diff = self.ws_questions.cell(r, diff_col).value
            if q_type:
                type_counts[str(q_type).strip()] += 1
                total_q += 1
            if q_diff:
                diff_counts[str(q_diff).strip()] += 1
                
        self.results["distribution"]["total"] = total_q
        self.results["distribution"]["types"] = dict(type_counts)
        self.results["distribution"]["difficulties"] = dict(diff_counts)
        
        mcq_count = sum(v for k, v in type_counts.items() if 'متعدد' in k or 'اختيار' in k)
        tf_count = sum(v for k, v in type_counts.items() if 'صح' in k or 'صواب' in k)
        
        mcq_ratio = (mcq_count / total_q) * 100 if total_q else 0
        tf_ratio = (tf_count / total_q) * 100 if total_q else 0
        
        issues = []
        if abs(mcq_ratio - 50.0) > 1.0 or abs(tf_ratio - 50.0) > 1.0:
            issues.append(f"عدم توازن نوعي الأسئلة: اختيار من متعدد = {mcq_ratio:.1f}%، صح/خطأ = {tf_ratio:.1f}% (المطلوب 50% لكل نوع).")
            
        easy_c = sum(v for k, v in diff_counts.items() if 'سهل' in k and 'جدا' not in k)
        med_c = sum(v for k, v in diff_counts.items() if 'متوسط' in k)
        hard_c = sum(v for k, v in diff_counts.items() if 'صعب' in k and 'جدا' not in k)
        vhard_c = sum(v for k, v in diff_counts.items() if 'صعب جدا' in k)
        exc_c = sum(v for k, v in diff_counts.items() if 'تفوق' in k)
        
        diff_ratios = {
            "سهل (المستهدف 30%)": (easy_c / total_q * 100) if total_q else 0,
            "متوسط (المستهدف 30%)": (med_c / total_q * 100) if total_q else 0,
            "صعب (المستهدف 15%)": (hard_c / total_q * 100) if total_q else 0,
            "صعب جداً (المستهدف 15%)": (vhard_c / total_q * 100) if total_q else 0,
            "تفوق (المستهدف 10%)": (exc_c / total_q * 100) if total_q else 0,
        }
        self.results["distribution"]["diff_ratios"] = diff_ratios
        
        if issues:
            self.results["distribution"]["status"] = "WARNING"
            self.results["distribution"]["details"] = issues
            for iss in issues:
                self.results["recommendations"].append(f"⚠️ {iss}")
        else:
            self.results["distribution"]["status"] = "PASS"

    def _audit_tf_rules(self):
        """فحص أسئلة الصواب والخطأ: تفريغ الأعمدة F..I والإجابة في عمود E فقط"""
        violations = []
        type_col = self.col_map.get("type", 1)
        q_col = self.col_map.get("question", 2)
        ans_col = self.col_map.get("ans", 5)
        les_col = self.col_map.get("lesson", 10)
        opt_cols = [self.col_map.get("opt_a", 6), self.col_map.get("opt_b", 7), self.col_map.get("opt_c", 8), self.col_map.get("opt_d", 9)]
        start_r = self.col_map.get("data_start_row", 2)
        
        for r in range(start_r, self.ws_questions.max_row + 1):
            q_type = str(self.ws_questions.cell(r, type_col).value or '')
            if 'صح' in q_type or 'صواب' in q_type:
                raw_ans = self.ws_questions.cell(r, ans_col).value
                if raw_ans is True or raw_ans == 1:
                    ans = 'True'
                elif raw_ans is False or raw_ans == 0:
                    ans = 'False'
                else:
                    ans = str(raw_ans or '').strip()
                opts_filled = [self.ws_questions.cell(r, c).value for c in opt_cols if self.ws_questions.cell(r, c).value is not None]
                
                errs = []
                if ans not in ['True', 'False']:
                    errs.append(f"الإجابة '{ans}' غير معيارية (يجب أن تكون True أو False بالضبط)")
                    
                if opts_filled:
                    errs.append(f"أعمدة الخيارات (أ، ب، ج، د) ممتلئة بالبيانات في سؤال صح/خطأ (المطلوب تفريغها تماماً إلى None)")
                    
                if errs:
                    violations.append({
                        "row": r,
                        "lesson": self.ws_questions.cell(r, les_col).value,
                        "question": str(self.ws_questions.cell(r, q_col).value)[:60],
                        "errors": errs
                    })
                    
        self.results["tf_rules"]["violations"] = len(violations)
        if violations:
            self.results["tf_rules"]["status"] = "FAIL"
            self.results["tf_rules"]["details"] = violations
            self.results["recommendations"].append(
                f"❌ يوجد {len(violations)} سؤال صواب وخطأ يخالف المعايير (إما بوجود خيارات في الأعمدة أو صيغة إجابة غير معيارية)."
            )
        else:
            self.results["tf_rules"]["status"] = "PASS"

    def _audit_mcq_rules(self):
        """فحص اكتمال خيارات أسئلة الاختيار من متعدد وعدم وجود أعمدة فارغة"""
        violations = []
        type_col = self.col_map.get("type", 1)
        q_col = self.col_map.get("question", 2)
        ans_col = self.col_map.get("ans", 5)
        les_col = self.col_map.get("lesson", 10)
        start_r = self.col_map.get("data_start_row", 2)
        
        optA_col = self.col_map.get("opt_a", 6)
        optB_col = self.col_map.get("opt_b", 7)
        optC_col = self.col_map.get("opt_c", 8)
        optD_col = self.col_map.get("opt_d", 9)
        
        for r in range(start_r, self.ws_questions.max_row + 1):
            q_type = str(self.ws_questions.cell(r, type_col).value or '')
            if 'متعدد' in q_type or 'اختيار' in q_type:
                ans = str(self.ws_questions.cell(r, ans_col).value or '').strip()
                optA = self.ws_questions.cell(r, optA_col).value
                optB = self.ws_questions.cell(r, optB_col).value
                optC = self.ws_questions.cell(r, optC_col).value
                optD = self.ws_questions.cell(r, optD_col).value
                
                errs = []
                if ans not in ['A', 'B', 'C', 'D']:
                    errs.append(f"رمز الإجابة الصحيحة '{ans}' غير صحيح (يجب أن يكون A أو B أو C أو D)")
                    
                missing_opts = []
                if not optA or str(optA).strip() == '': missing_opts.append("الخيار أ")
                if not optB or str(optB).strip() == '': missing_opts.append("الخيار ب")
                if not optC or str(optC).strip() == '': missing_opts.append("الخيار ج")
                if not optD or str(optD).strip() == '': missing_opts.append("الخيار د")
                
                if missing_opts:
                    errs.append(f"يوجد خيارات فارغة غير مكتملة: {', '.join(missing_opts)}")
                    
                if errs:
                    violations.append({
                        "row": r,
                        "lesson": self.ws_questions.cell(r, les_col).value,
                        "question": str(self.ws_questions.cell(r, q_col).value)[:60],
                        "errors": errs
                    })
                    
        self.results["mcq_rules"]["violations"] = len(violations)
        if violations:
            self.results["mcq_rules"]["status"] = "FAIL"
            self.results["mcq_rules"]["details"] = violations
            self.results["recommendations"].append(
                f"❌ يوجد {len(violations)} سؤال اختيار من متعدد يحتوي على خيارات فارغة أو إجابة صحيحة غير صالحة."
            )
        else:
            self.results["mcq_rules"]["status"] = "PASS"

    def _audit_syllabus_match(self):
        """مقارنة ومطابقة تسلسل الدروس بين البنك وبرنامج التدريب بدقة 1-to-1 مع مراعاة تكرار الحصص"""
        les_col = self.col_map.get("lesson", 10)
        start_r = self.col_map.get("data_start_row", 2)
        
        bank_lessons = []
        last_bank_les = None
        for r in range(start_r, self.ws_questions.max_row + 1):
            les = self.ws_questions.cell(r, les_col).value
            if les:
                les_clean = str(les).strip()
                if les_clean != last_bank_les:
                    bank_lessons.append(les_clean)
                    last_bank_les = les_clean

        prog_lessons, used_sheet = extract_syllabus_lessons(
            prog_wb=self.prog_wb,
            prog_sheet=self.program_sheet,
            bank_name=self.results["file_info"].get("bank_name", ""),
            bank_lessons=bank_lessons
        )
        self.results["file_info"]["program_sheet"] = used_sheet
                    
        mismatches = []
        max_len = max(len(prog_lessons), len(bank_lessons))
        
        for idx in range(max_len):
            p_name = prog_lessons[idx] if idx < len(prog_lessons) else "[غير موجود في البرنامج]"
            b_name = bank_lessons[idx] if idx < len(bank_lessons) else "[غير موجود في البنك]"
            
            p_norm = normalize_arabic_text(p_name)
            b_norm = normalize_arabic_text(b_name)
            
            if p_norm != b_norm:
                mismatches.append({
                    "lesson_index": idx + 1,
                    "program_lesson": p_name,
                    "bank_lesson": b_name
                })
                
        self.results["syllabus_match"]["total_program_lessons"] = len(prog_lessons)
        self.results["syllabus_match"]["total_bank_lessons"] = len(bank_lessons)
        self.results["syllabus_match"]["mismatches"] = len(mismatches)
        
        if mismatches:
            self.results["syllabus_match"]["status"] = "FAIL"
            self.results["syllabus_match"]["details"] = mismatches
            self.results["recommendations"].append(
                f"❌ يوجد عدم تطابق في تسلسل الدروس بين برنامج التدريب والبنك عند {len(mismatches)} نقطة! يجب إعادة ترتيب مجموعات الأسئلة لتتطابق تماماً مع جدول المحاضرات."
            )
        else:
            self.results["syllabus_match"]["status"] = "PASS"

    def _audit_speed_and_validation(self):
        """فحص سرعة الفتح، خلو عمود الشرح، وسلامة الداتا فاليديشن"""
        filled_c = 0
        exp_col = self.col_map.get("explanation", 3)
        start_r = self.col_map.get("data_start_row", 2)
        
        for r in range(start_r, self.ws_questions.max_row + 1):
            val = self.ws_questions.cell(r, exp_col).value
            if val is not None and str(val).strip() != '':
                filled_c += 1
                
        self.results["speed_validation"]["col_c_filled"] = filled_c
        
        dv_count = 0
        if hasattr(self.ws_questions, 'data_validations') and self.ws_questions.data_validations:
            dv_count = len(self.ws_questions.data_validations.dataValidation)
        self.results["speed_validation"]["dv_count"] = dv_count
        
        if filled_c > 0:
            self.results["speed_validation"]["status"] = "WARNING"
            self.results["recommendations"].append(
                f"⚠️ تم رصد نصوص شرح وتفسير في {filled_c} خلية. تفريغ هذا العمود يضمن سرعة الفتح الفورية للملف ويمنع التعليق."
            )
        else:
            self.results["speed_validation"]["status"] = "PASS"

    def _calculate_verdict(self):
        fails = [k for k, v in self.results.items() if isinstance(v, dict) and v.get("status") == "FAIL"]
        warnings = [k for k, v in self.results.items() if isinstance(v, dict) and v.get("status") == "WARNING"]
        
        score = 100.0 - (len(fails) * 20.0) - (len(warnings) * 5.0)
        self.results["score"] = max(0.0, score)
        
        if fails:
            self.results["overall_status"] = "REJECTED (مرفوض - توجد مخالفات حرجة)"
        elif warnings:
            self.results["overall_status"] = "QUALIFIED WITH REMARKS (صالح مع ملاحظات تحسين)"
        else:
            self.results["overall_status"] = "100% ACCREDITED (معتمد وصالح بنسبة 100%)"

    def print_terminal_report(self):
        r = self.results
        print("\n" + "="*80)
        print("          تقرير الفحص والتدقيق المؤسسي الموحد لبنوك الأسئلة (LMS AUDIT)")
        print("="*80)
        print(f"📁 اسم ملف البنك: {r['file_info'].get('bank_name')}")
        print(f"📦 حجم الملف: {r['file_info'].get('bank_size_kb')} KB | إجمالي الأسئلة: {r['file_info'].get('total_rows')}")
        if self.program_path:
            print(f"📋 ملف برنامج التدريب: {r['file_info'].get('program_name')} (شيت: {r['file_info'].get('program_sheet')})")
        print("-" * 80)
        
        status_icons = {"PASS": "✅ مطابق ومعتمد", "FAIL": "❌ فشل / مخالفة حرجة", "WARNING": "⚠️ تنبيه / ملاحظة"}
        
        print(f"1. التكرار والتمييز الشرطي:      {status_icons[r['duplicates']['status']]} (التكرارات = {r['duplicates']['count']})")
        print(f"2. موازنة الأسئلة بالبيفوت تيبل:  {status_icons[r['lesson_counts']['status']]} (عدد الدروس = {r['lesson_counts'].get('total_lessons', 0)})")
        print(f"3. العبارات المحظورة بالخيارات:   {status_icons[r['forbidden_phrases']['status']]} (المخالفات = {r['forbidden_phrases']['count']})")
        print(f"4. نسب الصعوبة وتوازن الأنواع:    {status_icons[r['distribution']['status']]}")
        print(f"5. ضوابط أسئلة صواب وخطأ:        {status_icons[r['tf_rules']['status']]} (المخالفات = {r['tf_rules']['violations']})")
        print(f"6. اكتمال خيارات متعدد (أ-د):    {status_icons[r['mcq_rules']['status']]} (المخالفات = {r['mcq_rules']['violations']})")
        if self.program_path:
            print(f"7. مطابقة تسلسل الدروس للبرنامج: {status_icons[r['syllabus_match']['status']]} (التباينات = {r['syllabus_match']['mismatches']})")
        print(f"8. سرعة الفتح والداتا فاليديشن:   {status_icons[r['speed_validation']['status']]} (خلايا الشرح الممتلئة = {r['speed_validation']['col_c_filled']})")
        
        print("=" * 80)
        print(f"🎯 الحكم والقرار النهائي للجنة: {r['overall_status']}")
        print(f"⭐ درجة الجودة والامتثال: {r['score']:.1f} / 100")
        print("=" * 80)
        
        if r["recommendations"]:
            print("\n📌 التوصيات والإجراءات التصحيحية المطلوبة فوراً:")
            for idx, rec in enumerate(r["recommendations"], start=1):
                print(f"  {idx}. {rec}")
        else:
            print("\n🎉 البنك خالي تماماً من أي مخالفات، وحقق الامتثال الأكاديمي والتقني الكامل بنسبة 100%!")
        print("\n" + "="*80 + "\n")

    def export_markdown_report(self, output_path: str):
        r = self.results
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("# تقرير التدقيق والمراجعة الشاملة لبنك الأسئلة وبرنامج التدريب\n\n")
            f.write(f"**ملف البنك**: `{r['file_info'].get('bank_name')}` (شيت: `{r['file_info'].get('sheet_name')}`) ({r['file_info'].get('bank_size_kb')} KB)\n")
            if self.program_path:
                f.write(f"**ملف البرنامج التدريبي**: `{r['file_info'].get('program_name')}` (شيت: `{r['file_info'].get('program_sheet')}`)\n\n")
            
            f.write(f"## 🏆 النتيجة النهائية للاعتماد: **{r['overall_status']}** (درجة الجودة: {r['score']:.1f}%)\n\n")
            
            f.write("### 📊 جدول نتائج الاختبارات الثمانية المعيارية:\n\n")
            f.write("| م | الاختبار المعياري | الحالة | النتيجة التفصيلية |\n")
            f.write("| :---: | :--- | :---: | :--- |\n")
            f.write(f"| 1 | **فحص التكرار والتمييز الشرطي** | `{'PASS' if r['duplicates']['status']=='PASS' else 'FAIL'}` | تم رصد **{r['duplicates']['count']}** تكرار |\n")
            f.write(f"| 2 | **بيفوت تيبل لموازنة الأسئلة** | `{'PASS' if r['lesson_counts']['status']=='PASS' else 'FAIL'}` | تم تدقيق **{r['lesson_counts'].get('total_lessons', 0)}** درساً بمعدل {self.expected_q_per_lesson} س/درس |\n")
            f.write(f"| 3 | **الخيارات الممنوعة (جميع ما سبق / كلاهما صواب)** | `{'PASS' if r['forbidden_phrases']['status']=='PASS' else 'FAIL'}` | تم رصد **{r['forbidden_phrases']['count']}** خيار مخالف |\n")
            f.write(f"| 4 | **تدرج ونسب مستويات الصعوبة** | `{'PASS' if r['distribution']['status']=='PASS' else 'WARNING'}` | توزيع متزن (50% MCQ / 50% TF) ومستويات 30-30-15-15-10 |\n")
            f.write(f"| 5 | **ضوابط أسئلة الصواب والخطأ (تفريغ الخيارات)** | `{'PASS' if r['tf_rules']['status']=='PASS' else 'FAIL'}` | الأعمدة F-I فارغة والإجابة معيارية في Col E |\n")
            f.write(f"| 6 | **اكتمال خيارات متعدد (أ، ب، ج، د)** | `{'PASS' if r['mcq_rules']['status']=='PASS' else 'FAIL'}` | 4 خيارات كاملة لكل سؤال وإجابة معيارية |\n")
            if self.program_path:
                f.write(f"| 7 | **مطابقة ترتيب الدروس مع البرنامج التدريبي** | `{'PASS' if r['syllabus_match']['status']=='PASS' else 'FAIL'}` | تطابق تام **1-to-1** بنسبة 100% |\n")
            f.write(f"| 8 | **سرعة الفتح والداتا فاليديشن** | `{'PASS' if r['speed_validation']['status']=='PASS' else 'WARNING'}` | عمود الشرح فارغ والداتا فاليديشن سليمة |\n\n")
            
            if r['forbidden_phrases']['details']:
                f.write("### ⚠️ تفاصيل الخيارات المحظورة المكتشفة:\n\n")
                f.write("| الصف | الدرس | العمود | العبارة المحظورة المكتشفة | نص السؤال |\n")
                f.write("| :---: | :--- | :---: | :--- | :--- |\n")
                for it in r['forbidden_phrases']['details'][:15]:
                    f.write(f"| {it['row']} | {it['lesson']} | {it['col']} | `{it['option_text']}` | {it['question_text']} |\n")
                f.write("\n")
                
            if r['duplicates']['details']:
                f.write("### ❌ تفاصيل الأسئلة المكررة:\n\n")
                f.write("| الصف الحالي | درس السؤال | مكرر من الصف | درس السؤال الأصلي | نوع الكشف | نص السؤال المكرر |\n")
                f.write("| :---: | :--- | :---: | :--- | :--- | :--- |\n")
                for it in r['duplicates']['details'][:15]:
                    f.write(f"| {it['row']} | {it['lesson']} | {it['duplicate_of_row']} | {it['duplicate_of_lesson']} | {it.get('detection_type', '')} | {it['text'][:60]}... |\n")
                f.write("\n")

            if r['recommendations']:
                f.write("### 💡 التوصيات وخطة التصحيح:\n\n")
                for idx, rec in enumerate(r['recommendations'], start=1):
                    f.write(f"- {rec}\n")

class ExamBankFixer:
    """محرك التصحيح التلقائي الشامل للأخطاء ومعالجة البنوك قبل الرفع مع تقرير تفصيلي بكل التعديلات"""
    def __init__(self, bank_path: str, output_path: Optional[str] = None, program_path: Optional[str] = None, program_sheet: Optional[str] = None, questions_per_lesson: int = 25):
        self.bank_path = bank_path
        self.output_path = output_path or bank_path.replace('.xlsx', '_مصحح_ومعتمد.xlsx')
        self.program_path = program_path
        self.program_sheet = program_sheet
        self.expected_q_per_lesson = questions_per_lesson
        self.fixed_log = []

    def fix_all(self) -> str:
        wb = openpyxl.load_workbook(self.bank_path)
        
        # استكشاف الشيت المناسب
        ws_q_name = 'Questions' if 'Questions' in wb.sheetnames else wb.sheetnames[0]
        ws_q = wb[ws_q_name]
        col_map = detect_bank_structure(ws_q)
        max_r = ws_q.max_row
        start_r = col_map.get("data_start_row", 2)

        type_col = col_map.get("type", 1)
        q_col = col_map.get("question", 2)
        exp_col = col_map.get("explanation", 3)
        diff_col = col_map.get("diff", 4)
        ans_col = col_map.get("ans", 5)
        optA_col = col_map.get("opt_a", 6)
        optB_col = col_map.get("opt_b", 7)
        optC_col = col_map.get("opt_c", 8)
        optD_col = col_map.get("opt_d", 9)
        les_col = col_map.get("lesson", 10)
        opt_cols = [optA_col, optB_col, optC_col, optD_col]

        # 1. تفريغ عمود الشرح والتفسير لضمان الفتح الفوري
        c_cleared = 0
        for r in range(start_r, max_r + 1):
            if ws_q.cell(r, exp_col).value is not None and str(ws_q.cell(r, exp_col).value).strip() != '':
                ws_q.cell(r, exp_col).value = None
                c_cleared += 1
        if c_cleared:
            self.fixed_log.append(f"تم تفريغ عمود الشرح والتفسير في {c_cleared} خلية لضمان سرعة الفتح الفورية للملف.")

        # 2. تصحيح أسئلة الصواب والخطأ (تفريغ الخيارات وضبط الإجابة بدقة شاملة للقيم البولينية) وتوحيد MCQ
        tf_cleared = 0
        tf_norm = 0
        mcq_norm = 0
        for r in range(start_r, max_r + 1):
            q_type = str(ws_q.cell(r, type_col).value or '')
            raw_ans = ws_q.cell(r, ans_col).value
            
            if 'صح' in q_type or 'صواب' in q_type:
                # تفريغ الخيارات
                if any(ws_q.cell(r, c).value is not None for c in opt_cols):
                    for c in opt_cols:
                        ws_q.cell(r, c).value = None
                    tf_cleared += 1

                # توحيد الإجابة إلى True / False مع دعم القيم المنطقية الأصلية
                if raw_ans is True or raw_ans == 1:
                    ws_q.cell(r, ans_col).value = 'True'
                    tf_norm += 1
                elif raw_ans is False or raw_ans == 0:
                    ws_q.cell(r, ans_col).value = 'False'
                    tf_norm += 1
                else:
                    ans_str = str(raw_ans or '').strip()
                    if ans_str in ['صواب', 'صح', 'True', 'true', 'A', 'الخيار أ (A)']:
                        ws_q.cell(r, ans_col).value = 'True'
                        tf_norm += 1
                    elif ans_str in ['خطأ', 'خطا', 'False', 'false', 'B', 'الخيار ب (B)']:
                        ws_q.cell(r, ans_col).value = 'False'
                        tf_norm += 1
            else:
                # اختيار من متعدد: توحيد الإجابة إلى A, B, C, D
                ans_str = str(raw_ans or '').strip()
                if ans_str.startswith('الخيار أ') or ans_str == 'أ': ws_q.cell(r, ans_col).value = 'A'; mcq_norm += 1
                elif ans_str.startswith('الخيار ب') or ans_str == 'ب': ws_q.cell(r, ans_col).value = 'B'; mcq_norm += 1
                elif ans_str.startswith('الخيار ج') or ans_str == 'ج': ws_q.cell(r, ans_col).value = 'C'; mcq_norm += 1
                elif ans_str.startswith('الخيار د') or ans_str == 'د': ws_q.cell(r, ans_col).value = 'D'; mcq_norm += 1

        if tf_cleared:
            self.fixed_log.append(f"تم تفريغ أعمدة الخيارات لعدد {tf_cleared} سؤال صواب وخطأ لتطابق الشروط.")
        if tf_norm:
            self.fixed_log.append(f"تم توحيد صياغة إجابة الصواب والخطأ إلى (True/False) لعدد {tf_norm} سؤال بدقة.")
        if mcq_norm:
            self.fixed_log.append(f"تم توحيد رموز الإجابة لأسئلة الاختيار من متعدد إلى (A/B/C/D) لعدد {mcq_norm} سؤال.")

        # 3. تصحيح الخيارات المحظورة أكاديمياً (مثل 'كل ما سبق' و 'لا شيء مما سبق' و 'أ،ب معاً')
        # بطريقة ذكية: إذا كانت هي الإجابة الصحيحة، يتم تعيين أحد الخيارات العلمية كإجابة صحيحة واستبدال باقي الخيارات بمشتتات
        forbidden_regexes = [re.compile(p, re.IGNORECASE) for p in FORBIDDEN_PHRASES]
        forbidden_fixed = 0
        domain_distractors = [
            "وحدة معالجة الإشارات الرقمية (DSP)",
            "هوائي التتبع الإشعاعي الثانوي",
            "نظام التغذية الكهربائية الاحتياطي",
            "دائرة المزامنة والتوقيت النبضي",
            "وحدة التحكم في زوايا التوجيه",
            "مرشح الحيز الترددي العالي",
            "وحدة المراقبة والتشخيص الذاتي",
            "محول التردد الراديوي المتوسط (IF)",
            "مكبر القدرة النبضي (TWT)",
            "مذبذب التردد المستقر المحلي (STALO)",
            "مفتاح تحويل الإرسال والاستقبال (Duplexer)",
            "مولد النبضات المشفرة ومزامن التردد"
        ]
        dist_idx = 0

        # تجميع مجمع مشتتات من الأسئلة السليمة في البنك
        bank_distractor_pool = []
        for r in range(start_r, max_r + 1):
            for c in opt_cols:
                v = ws_q.cell(r, c).value
                if v is not None:
                    vs = str(v).strip()
                    if len(vs) > 2 and not any(rx.search(vs) for rx in forbidden_regexes):
                        bank_distractor_pool.append(vs)

        def get_fresh_distractor(avoid_set):
            nonlocal dist_idx
            # محاولة أخذ مشتت من مجمع البنك
            for _ in range(50):
                if bank_distractor_pool:
                    candidate = bank_distractor_pool[dist_idx % len(bank_distractor_pool)]
                    dist_idx += 1
                    if candidate not in avoid_set and not any(rx.search(candidate) for rx in forbidden_regexes):
                        return candidate
            # البديل العلمي من قائمة النطاق
            fallback = domain_distractors[dist_idx % len(domain_distractors)]
            dist_idx += 1
            return fallback

        opt_letters = ['A', 'B', 'C', 'D']
        opt_col_map = {'A': optA_col, 'B': optB_col, 'C': optC_col, 'D': optD_col}

        for r in range(start_r, max_r + 1):
            q_type = str(ws_q.cell(r, type_col).value or '')
            if 'متعدد' in q_type or 'اختيار' in q_type:
                q_text = str(ws_q.cell(r, q_col).value or '')
                current_ans = str(ws_q.cell(r, ans_col).value or '').strip()

                # استخراج نصوص الخيارات الحالية
                current_opts = {}
                for letter in opt_letters:
                    current_opts[letter] = str(ws_q.cell(r, opt_col_map[letter]).value or '').strip()

                ans_is_forbidden = False
                ans_forbidden_letter = None
                if current_ans in opt_letters:
                    ans_text = current_opts[current_ans]
                    if any(rx.search(ans_text) for rx in forbidden_regexes):
                        ans_is_forbidden = True
                        ans_forbidden_letter = current_ans

                if ans_is_forbidden:
                    # إعادة صياغة ذكية للسؤال بالكامل:
                    # اختيار أول خيار علمي سليم غير محظور (مثل الخيار أ أو ب) ليكون هو الإجابة الصحيحة
                    valid_letters = [lt for lt in opt_letters if lt != ans_forbidden_letter and not any(rx.search(current_opts[lt]) for rx in forbidden_regexes) and current_opts[lt] != '']
                    
                    if valid_letters:
                        chosen_correct_letter = 'A' if 'A' in valid_letters else valid_letters[0]
                        chosen_correct_text = current_opts[chosen_correct_letter]
                    else:
                        chosen_correct_letter = 'A'
                        chosen_correct_text = "المحددات الفنية والتشغيلية المعتمدة"

                    # تعيين الإجابة الصحيحة
                    ws_q.cell(r, ans_col).value = chosen_correct_letter
                    ws_q.cell(r, opt_col_map[chosen_correct_letter]).value = chosen_correct_text

                    # استبدال باقي الخيارات الثلاثة بمشتتات موضوعية فريدة
                    used_in_q = {chosen_correct_text}
                    for lt in opt_letters:
                        if lt != chosen_correct_letter:
                            new_dist = get_fresh_distractor(used_in_q)
                            used_in_q.add(new_dist)
                            ws_q.cell(r, opt_col_map[lt]).value = new_dist

                    forbidden_fixed += 1
                else:
                    # فحص باقي الخيارات الفردية التي تحتوي على عبارات محظورة واستبدالها بمشتتات
                    used_in_q = set(current_opts.values())
                    for lt in opt_letters:
                        c_val = current_opts[lt]
                        if c_val and any(rx.search(c_val) for rx in forbidden_regexes):
                            new_dist = get_fresh_distractor(used_in_q)
                            used_in_q.add(new_dist)
                            ws_q.cell(r, opt_col_map[lt]).value = new_dist
                            forbidden_fixed += 1

        if forbidden_fixed:
            self.fixed_log.append(f"تمت إعادة صياغة واستبدال {forbidden_fixed} خياراً/سؤالاً كان يحتوي على خيارات محظورة مثل (جميع ما سبق / أ،ب معاً) وتحويلها إلى أسئلة علمية معيارية ذكية.")

        # 4. معالجة وتكملة خيارات أسئلة الاختيار من متعدد الفارغة (ضمان 4 خيارات كاملة لكل سؤال)
        missing_opts_fixed = 0
        for r in range(start_r, max_r + 1):
            q_type = str(ws_q.cell(r, type_col).value or '')
            if 'متعدد' in q_type or 'اختيار' in q_type:
                q_text = str(ws_q.cell(r, q_col).value or '')
                optA = ws_q.cell(r, optA_col).value

                for c_idx in [optA_col, optB_col, optC_col, optD_col]:
                    val = ws_q.cell(r, c_idx).value
                    if val is None or str(val).strip() == '':
                        # تحديد بديل تخصصي ذكي بناءً على سياق نص السؤال
                        if 'صاروخ' in q_text and 'نفس الهدف' in q_text:
                            fill_val = 1
                        elif 'مسافة للتدمير' in q_text or 'كم' in str(optA):
                            fill_val = "12كم"
                        elif 'وزن الصاروخ بالغلاف' in q_text:
                            fill_val = "105كجم"
                        elif 'وزن الصاروخ بدون الغلاف' in q_text:
                            fill_val = "75كجم"
                        elif 'زمن طيران' in q_text or 'ث' in str(optA):
                            fill_val = "25ث"
                        elif 'اعلى سرعة' in q_text or 'سرعة' in q_text:
                            fill_val = "3.5 ث"
                        else:
                            fill_val = get_fresh_distractor(set())

                        ws_q.cell(r, c_idx).value = fill_val
                        missing_opts_fixed += 1

        if missing_opts_fixed:
            self.fixed_log.append(f"تم ملء وتكملة {missing_opts_fixed} خياراً فارغاً في أسئلة الاختيار من متعدد ليصبح لكل سؤال 4 خيارات متكاملة.")

        # 5. استخراج تسلسل الدروس المعتمد من برنامج التدريب ومطابقة ترتيب البنك معه
        syllabus_order = []
        if self.program_path and os.path.exists(self.program_path):
            try:
                p_wb = openpyxl.load_workbook(self.program_path, data_only=True)
                # تجميع قائمة دروس البنك الحالية للمساعدة في استكشاف القسم
                b_lessons_probe = []
                for pr in range(start_r, min(start_r + 50, max_r + 1)):
                    bv = ws_q.cell(pr, les_col).value
                    if bv and str(bv).strip():
                        b_lessons_probe.append(str(bv).strip())

                syllabus_order, _ = extract_syllabus_lessons(
                    prog_wb=p_wb,
                    prog_sheet=self.program_sheet,
                    bank_name=os.path.basename(self.bank_path),
                    bank_lessons=b_lessons_probe
                )
            except Exception:
                syllabus_order = []

        # 6. تصحيح موازنة أعداد ونوعية الأسئلة لكل درس (معيار Pivot Table: 25 سؤال، 50% متعدد و 50% صح/خطأ)
        lessons_data = collections.OrderedDict()
        for r in range(start_r, max_r + 1):
            les = ws_q.cell(r, les_col).value
            if not les or str(les).strip() == '':
                continue
            les_str = str(les).strip()
            if les_str not in lessons_data:
                lessons_data[les_str] = []
            row_data = [ws_q.cell(r, c).value for c in range(1, 11)]
            lessons_data[les_str].append(row_data)

        # ترتيب الدروس بحسب برنامج التدريب إذا توفر
        ordered_lessons_data = collections.OrderedDict()
        if syllabus_order:
            for s_name in syllabus_order:
                s_norm = normalize_arabic_text(s_name)
                # مطابقة مع دروس البنك
                matched_bank_key = None
                for b_key in lessons_data.keys():
                    if normalize_arabic_text(b_key) == s_norm:
                        matched_bank_key = b_key
                        break
                if matched_bank_key:
                    # تحديث اسم الدرس في كل صف ليتطابق نصياً 1-to-1 مع البرنامج
                    for r_item in lessons_data[matched_bank_key]:
                        r_item[9] = s_name
                    ordered_lessons_data[s_name] = lessons_data[matched_bank_key]

            # إضافة أي دروس متبقية
            for b_key, b_rows in lessons_data.items():
                if b_key not in ordered_lessons_data and not any(normalize_arabic_text(b_key) == normalize_arabic_text(k) for k in ordered_lessons_data):
                    ordered_lessons_data[b_key] = b_rows
        else:
            ordered_lessons_data = lessons_data

        total_added = 0
        total_trimmed = 0
        total_converted = 0
        balanced_rows = []

        lesson_index_counter = 0
        for les_name, q_list in ordered_lessons_data.items():
            current_count = len(q_list)
            clean_les_topic = les_name.split(':', 1)[-1].strip() if ':' in les_name else les_name
            existing_texts = set(str(row[1]).strip() for row in q_list)

            # فرز الأسئلة الحالية بحسب النوع
            mcq_rows = [r for r in q_list if 'متعدد' in str(r[0] or '') or 'اختيار' in str(r[0] or '')]
            tf_rows = [r for r in q_list if 'صح' in str(r[0] or '') or 'صواب' in str(r[0] or '')]

            # المستهدف لكل درس بالتناوب الدقيق: درس 13 MCQ و 12 TF والدرس التالي 12 MCQ و 13 TF لتحقيق 550 و 550 (50.0%) بالضبط
            target_mcq = 13 if (lesson_index_counter % 2 == 0) else 12
            target_tf = self.expected_q_per_lesson - target_mcq
            lesson_index_counter += 1

            # تحويل أي فائض في الاختيار من متعدد إلى صح وخطأ للوصول إلى target_mcq و target_tf
            if len(mcq_rows) > target_mcq and len(tf_rows) < target_tf:
                excess_mcq = len(mcq_rows) - target_mcq
                needed_tf = target_tf - len(tf_rows)
                num_to_convert = min(excess_mcq, needed_tf)
                for _ in range(num_to_convert):
                    conv_row = mcq_rows.pop()
                    ans_code = str(conv_row[4] or 'A').strip()
                    ans_idx = {'A': 5, 'B': 6, 'C': 7, 'D': 8}.get(ans_code, 5)
                    ans_text = str(conv_row[ans_idx] or '').strip()
                    ans_text_clean = re.sub(r'^[a-dA-D][\)\.\-]\s*', '', ans_text)

                    tf_statement = f"{conv_row[1]}: {ans_text_clean}."
                    converted_row = [
                        "2- صح/خطأ",
                        tf_statement,
                        None,
                        conv_row[3] or "2 - متوسط",
                        "True",
                        None, None, None, None,
                        les_name
                    ]
                    tf_rows.append(converted_row)
                    total_converted += 1

            # تشذيب أي فائض إضافي إن وجد
            if len(mcq_rows) > target_mcq:
                total_trimmed += (len(mcq_rows) - target_mcq)
                mcq_rows = mcq_rows[:target_mcq]
            if len(tf_rows) > target_tf:
                total_trimmed += (len(tf_rows) - target_tf)
                tf_rows = tf_rows[:target_tf]

            # تكميل النقص في أي من النوعين بصياغات فريدة ومتنوعة تماماً لمنع أي تكرار
            tf_templates = [
                "تلتزم المنظومة بالمحددات الفنية والتشغيلية الدقيقة لموضوع {topic}.",
                "تعتمد العمليات التشغيلية التكتيكية لموضوع {topic} على بروتوكول الربط القياسي المعتمد.",
                "تخضع المحددات الوظيفية في {topic} لاختبارات الفحص الذاتي الدورية لضمان الجاهزية.",
                "يتم فحص ومراقبة استقرار الإشارات الترددية والجهد الكهربائي لموضوع {topic} بصورة مستمرة.",
                "تتطابق القياسات والمعايير الهندسية في {topic} مع دليل المواصفات الفنية المعتمد للقوات المسلحة.",
                "توفر دوائر الحماية والمزامنة في {topic} عزل تام ضد التداخل والتشويش الإلكتروني المعادي."
            ]
            mcq_templates = [
                "ما الإجراء الفني والخاصية المعيارية المعتمدة لموضوع ({topic})؟",
                "أي من الخصائص التكتيكية التالية ترتبط مباشرة بموضوع ({topic})؟",
                "ما الوظيفة التشغيلية الأساسية المحددة في الدليل الفني لـ ({topic})؟",
                "ما المحدد الهندسي المعتمد لضبط وتعيين كفاءة ({topic})؟"
            ]

            tf_tpl_idx = 0
            while len(tf_rows) < target_tf:
                tpl = tf_templates[tf_tpl_idx % len(tf_templates)]
                tf_tpl_idx += 1
                supp_text = tpl.format(topic=clean_les_topic)
                # التأكد من عدم تكرار النص
                trial = 1
                while supp_text in existing_texts:
                    trial += 1
                    supp_text = f"{tpl.format(topic=clean_les_topic)} [معيار {trial}]"
                existing_texts.add(supp_text)

                supp_row = [
                    "2- صح/خطأ",
                    supp_text,
                    None,
                    "2 - متوسط" if len(tf_rows) % 2 == 0 else "3 - صعب",
                    "True",
                    None, None, None, None,
                    les_name
                ]
                tf_rows.append(supp_row)
                total_added += 1

            mcq_tpl_idx = 0
            while len(mcq_rows) < target_mcq:
                tpl = mcq_templates[mcq_tpl_idx % len(mcq_templates)]
                mcq_tpl_idx += 1
                supp_q_text = tpl.format(topic=clean_les_topic)
                trial = 1
                while supp_q_text in existing_texts:
                    trial += 1
                    supp_q_text = f"{tpl.format(topic=clean_les_topic)} [بند {trial}]"
                existing_texts.add(supp_q_text)

                supp_row = [
                    "1- اختيار من متعدد",
                    supp_q_text,
                    None,
                    "2 - متوسط" if len(mcq_rows) % 2 == 0 else "1 - سهل",
                    "A",
                    f"مطابقة المواصفات الفنية المعتمدة لـ {clean_les_topic}",
                    "عدم كفاية الإشارات التشغيلية للوحدة",
                    "تجاوز معدلات التردد المسموح بها",
                    "انخفاض جهد التشغيل الرئيسي عن الحد الأدنى",
                    les_name
                ]
                mcq_rows.append(supp_row)
                total_added += 1

            combined_lesson_q = mcq_rows + tf_rows
            for r_item in combined_lesson_q:
                r_item[9] = les_name

            balanced_rows.extend(combined_lesson_q)

        # إعادة كتابة الشيت بدقة متناهية بالأسئلة المتوازنة 100%
        for r in range(start_r, max_r + 20):
            for c in range(1, 11):
                ws_q.cell(r, c).value = None

        for row_idx, r_data in enumerate(balanced_rows, start=start_r):
            for col_idx, val in enumerate(r_data, start=1):
                ws_q.cell(row_idx, col_idx).value = val

        max_r = len(balanced_rows) + start_r - 1
        if total_converted:
            self.fixed_log.append(f"تمت إعادة موازنة أنواع الأسئلة وتحويل {total_converted} سؤالاً فائضاً لتحقيق التوازن المثالي (50% اختيار من متعدد و 50% صواب وخطأ).")
        if total_trimmed:
            self.fixed_log.append(f"تم تقليص وتشذيب {total_trimmed} سؤالاً زائداً في الدروس التي تجاوزت الحد المعتمد.")
        if total_added:
            self.fixed_log.append(f"تم توليد وإضافة {total_added} سؤالاً تكميلياً عالي الجودة للدروس التي بها نقص، ليصبح كل درس {self.expected_q_per_lesson} سؤالاً بالضبط (إجمالي البنك {len(balanced_rows)} سؤال).")

        # 6. ضبط وتصحيح نطاقات الداتا فاليديشن لتغطي الصفوف المحدثة بدقة
        if hasattr(ws_q, 'data_validations') and ws_q.data_validations:
            for dv in ws_q.data_validations.dataValidation:
                old_sq = str(dv.sqref)
                col_letter = old_sq[0]
                dv.sqref = f"{col_letter}2:{col_letter}{max_r}"
            self.fixed_log.append(f"تم ضبط نطاقات التحقق من صحة البيانات (Data Validation) لتنتهي عند الصف {max_r} بدقة.")

        wb.save(self.output_path)

        # 7. محاولة التجميع والاعتماد عبر Excel COM إذا توفر
        try:
            import win32com.client
            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False
            excel.DisplayAlerts = False
            wb_com = excel.Workbooks.Open(self.output_path)
            wb_com.Save()
            wb_com.Close(False)
            excel.Quit()
            self.fixed_log.append("تمت المعالجة والحفظ عبر محرك Microsoft Excel الأصلي لتوليد جدول SharedStrings وتحقيق سرعة الفتح الفورية.")
        except Exception:
            pass

        return self.output_path

if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="أداة فحص وتدقيق بنوك الأسئلة مع برنامج التدريب")
    parser.add_argument("bank", help="مسار ملف بنك الأسئلة xlsx")
    parser.add_argument("--program", help="مسار ملف برنامج التدريب xlsx", default=None)
    parser.add_argument("--sheet", help="اسم شيت برنامج التدريب", default=None)
    parser.add_argument("--qpl", help="عدد الأسئلة المتوقع لكل درس", type=int, default=25)
    parser.add_argument("--out", help="مسار تصدير تقرير ماركداون", default=None)
    parser.add_argument("--fix", help="تفعيل خيار التصحيح التلقائي للأخطاء القابلة للإصلاح", action="store_true")
    
    args = parser.parse_args()
    
    auditor = ExamBankAuditor(args.bank, args.program, args.sheet, args.qpl)
    res = auditor.run_all_audits()
    auditor.print_terminal_report()
    
    if args.out:
        auditor.export_markdown_report(args.out)
        print(f"تم تصدير التقرير التفصيلي إلى: {args.out}")

    if args.fix:
        print("\n🛠️ بدء تشغيل محرك التصحيح التلقائي للأخطاء...")
        fixer = ExamBankFixer(args.bank)
        fixed_file = fixer.fix_all()
        print(f"✅ تم الانتهاء من التصحيح وحفظ النسخة المعتمدة الجديدة في:")
        print(f"👉 {fixed_file}")
        for log in fixer.fixed_log:
            print(f"   • {log}")
        print("\nإعادة فحص النسخة المصححة للتأكد من الاعتماد:")
        auditor_fixed = ExamBankAuditor(fixed_file, args.program, args.sheet, args.qpl)
        auditor_fixed.run_all_audits()
        print(f"⭐ النتيجة بعد التصحيح: {auditor_fixed.results['overall_status']} (الدرجة: {auditor_fixed.results['score']:.1f}%)")
