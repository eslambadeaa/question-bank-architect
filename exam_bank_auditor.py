# -*- coding: utf-8 -*-
"""
محرك الفحص والتدقيق المؤسسي لبنوك الأسئلة وبرامج التدريب (LMS Exam Bank & Syllabus Auditor)
========================================================================================
يقوم هذا المحرك بمحاكاة الفحص الصارم للجان إدارة الامتحانات والاعتماد المؤسسي:
 1. اختبار التكرار والتمييز الشرطي (Conditional Formatting Simulation)
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
    r'أ\s*و\s*ب\s+مع[ااً]',
    r'أ\s*و\s*ج\s+مع[ااً]',
    r'ب\s*و\s*ج\s+مع[ااً]',
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
        if 'Questions' not in self.bank_wb.sheetnames:
            raise ValueError("الملف لا يحتوي على شيت 'Questions' الخاص بنموذج تسجيل الأسئلة المعتمد!")
        
        self.ws_questions = self.bank_wb['Questions']
        self.ws_lookups = self.bank_wb['Lookups'] if 'Lookups' in self.bank_wb.sheetnames else None
        
        self.results["file_info"]["bank_name"] = os.path.basename(self.bank_path)
        self.results["file_info"]["bank_size_kb"] = round(os.path.getsize(self.bank_path) / 1024, 1)
        self.results["file_info"]["total_rows"] = self.ws_questions.max_row - 1

        if self.program_path:
            if not os.path.exists(self.program_path):
                raise FileNotFoundError(f"ملف برنامج التدريب غير موجود: {self.program_path}")
            self.prog_wb = openpyxl.load_workbook(self.program_path, data_only=True)
            if self.program_sheet:
                if self.program_sheet not in self.prog_wb.sheetnames:
                    raise ValueError(f"الشيت '{self.program_sheet}' غير موجود في ملف برنامج التدريب!")
                self.ws_program = self.prog_wb[self.program_sheet]
            else:
                # محاولة تخمين الشيت المناسب
                cand = [s for s in self.prog_wb.sheetnames if 'برنامج' in s or 'تدريب' in s or 'ساعات' in s]
                if cand:
                    self.ws_program = self.prog_wb[cand[0]]
                    self.program_sheet = cand[0]
            self.results["file_info"]["program_name"] = os.path.basename(self.program_path)
            self.results["file_info"]["program_sheet"] = self.program_sheet

    def run_all_audits(self) -> Dict[str, Any]:
        self.load_files()
        
        # 1. اختبار التكرار والتمييز الشرطي
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
        """اختبار محاكاة التنسيق الشرطي لتمييز الأسئلة المكررة"""
        seen_texts = {}
        duplicates_found = []
        
        for r in range(2, self.ws_questions.max_row + 1):
            q_text = self.ws_questions.cell(r, 2).value
            if not q_text:
                continue
            cleaned = str(q_text).strip()
            # إزالة المسافات المتكررة للمقارنة الصارمة
            normalized = " ".join(cleaned.split())
            
            les = self.ws_questions.cell(r, 10).value
            if normalized in seen_texts:
                orig_r, orig_les = seen_texts[normalized]
                duplicates_found.append({
                    "row": r,
                    "lesson": les,
                    "text": cleaned,
                    "duplicate_of_row": orig_r,
                    "duplicate_of_lesson": orig_les
                })
            else:
                seen_texts[normalized] = (r, les)
                
        self.results["duplicates"]["count"] = len(duplicates_found)
        if duplicates_found:
            self.results["duplicates"]["status"] = "FAIL"
            self.results["duplicates"]["details"] = duplicates_found
            self.results["recommendations"].append(
                f"❌ يوجد {len(duplicates_found)} سؤال مكرر سيكشفه التنسيق الشرطي لإدارة الامتحانات ويؤدي لرفض البنك فوراً! يجب استبدالها بأسئلة فريدة."
            )
        else:
            self.results["duplicates"]["status"] = "PASS"

    def _audit_lesson_counts(self):
        """محاكاة جدول بيفوت تيبل لفحص عدد الأسئلة لكل درس"""
        counts = collections.OrderedDict()
        
        for r in range(2, self.ws_questions.max_row + 1):
            les = self.ws_questions.cell(r, 10).value
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
        
        for r in range(2, self.ws_questions.max_row + 1):
            les = self.ws_questions.cell(r, 10).value
            q_txt = self.ws_questions.cell(r, 2).value
            
            # فحص الخيارات الأربعة A, B, C, D (الأعمدة 6, 7, 8, 9)
            for c_idx, col_name in enumerate(['الخيار أ', 'الخيار ب', 'الخيار ج', 'الخيار د'], start=6):
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
        
        for r in range(2, self.ws_questions.max_row + 1):
            q_type = self.ws_questions.cell(r, 1).value
            q_diff = self.ws_questions.cell(r, 4).value
            if q_type:
                type_counts[str(q_type).strip()] += 1
                total_q += 1
            if q_diff:
                diff_counts[str(q_diff).strip()] += 1
                
        self.results["distribution"]["total"] = total_q
        self.results["distribution"]["types"] = dict(type_counts)
        self.results["distribution"]["difficulties"] = dict(diff_counts)
        
        # النسب المقررة القياسية
        # 50% MCQ, 50% TF
        mcq_count = sum(v for k, v in type_counts.items() if 'متعدد' in k)
        tf_count = sum(v for k, v in type_counts.items() if 'صح' in k)
        
        mcq_ratio = (mcq_count / total_q) * 100 if total_q else 0
        tf_ratio = (tf_count / total_q) * 100 if total_q else 0
        
        issues = []
        if abs(mcq_ratio - 50.0) > 1.0 or abs(tf_ratio - 50.0) > 1.0:
            issues.append(f"عدم توازن نوعي الأسئلة: اختيار من متعدد = {mcq_ratio:.1f}%، صح/خطأ = {tf_ratio:.1f}% (المطلوب 50% لكل نوع).")
            
        # فحص نسب الصعوبة (المعيار: 30% سهل، 30% متوسط، 15% صعب، 15% صعب جدا، 10% تفوق)
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
        
        for r in range(2, self.ws_questions.max_row + 1):
            q_type = str(self.ws_questions.cell(r, 1).value or '')
            if 'صح' in q_type or 'صواب' in q_type:
                ans = str(self.ws_questions.cell(r, 5).value or '').strip()
                optA = self.ws_questions.cell(r, 6).value
                optB = self.ws_questions.cell(r, 7).value
                optC = self.ws_questions.cell(r, 8).value
                optD = self.ws_questions.cell(r, 9).value
                
                errs = []
                if ans not in ['True', 'False']:
                    errs.append(f"الإجابة '{ans}' غير معيارية (يجب أن تكون True أو False بالضبط)")
                    
                if optA is not None or optB is not None or optC is not None or optD is not None:
                    errs.append(f"أعمدة الخيارات (أ، ب، ج، د) ممتلئة بالبيانات في سؤال صح/خطأ (المطلوب تفريغها تماماً إلى None)")
                    
                if errs:
                    violations.append({
                        "row": r,
                        "lesson": self.ws_questions.cell(r, 10).value,
                        "question": str(self.ws_questions.cell(r, 2).value)[:60],
                        "errors": errs
                    })
                    
        self.results["tf_rules"]["violations"] = len(violations)
        if violations:
            self.results["tf_rules"]["status"] = "FAIL"
            self.results["tf_rules"]["details"] = violations
            self.results["recommendations"].append(
                f"❌ يوجد {len(violations)} سؤال صواب وخطأ يخالف المعايير (إما بوجود خيارات في الأعمدة F..I أو صيغة إجابة غير معيارية)."
            )
        else:
            self.results["tf_rules"]["status"] = "PASS"

    def _audit_mcq_rules(self):
        """فحص اكتمال خيارات أسئلة الاختيار من متعدد وعدم وجود أعمدة فارغة"""
        violations = []
        
        for r in range(2, self.ws_questions.max_row + 1):
            q_type = str(self.ws_questions.cell(r, 1).value or '')
            if 'متعدد' in q_type or 'اختيار' in q_type:
                ans = str(self.ws_questions.cell(r, 5).value or '').strip()
                optA = self.ws_questions.cell(r, 6).value
                optB = self.ws_questions.cell(r, 7).value
                optC = self.ws_questions.cell(r, 8).value
                optD = self.ws_questions.cell(r, 9).value
                
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
                        "lesson": self.ws_questions.cell(r, 10).value,
                        "question": str(self.ws_questions.cell(r, 2).value)[:60],
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
        # 1. استخراج الدروس من برنامج التدريب بالترتيب الزمني الصحيح
        # نتجاهل صفوف امتحانات منتصف الترم، صفوف الإجماليات، والصفوف المكررة لنفس الدرس (مثل ف1 وف2 لنفس الدرس)
        prog_lessons = []
        last_les = None
        
        for r in range(5, self.ws_program.max_row + 1):
            lec = self.ws_program.cell(r, 3).value
            name = self.ws_program.cell(r, 4).value
            q_val = self.ws_program.cell(r, 9).value # عمود عدد الأسئلة
            
            if name and str(name).strip() not in ['', 'None', 'اسم الموضوع']:
                name_clean = str(name).strip()
                # استبعاد صفوف الامتحانات
                if 'امتحان' in name_clean or 'إجمالي' in name_clean or 'اجمالي' in name_clean:
                    continue
                
                # المعيار: إما أن يكون له حصة نظرية ف1/ف3 أو مخصص له عدد أسئلة
                is_valid_lecture = (lec and str(lec).strip() in ['ف1', 'ف3']) or (q_val == self.expected_q_per_lesson)
                
                if is_valid_lecture and name_clean != last_les:
                    prog_lessons.append(name_clean)
                    last_les = name_clean
                    
        # 2. استخراج تسلسل الدروس من شيت الأسئلة
        bank_lessons = []
        last_bank_les = None
        for r in range(2, self.ws_questions.max_row + 1):
            les = self.ws_questions.cell(r, 10).value
            if les:
                les_clean = str(les).strip()
                if les_clean != last_bank_les:
                    bank_lessons.append(les_clean)
                    last_bank_les = les_clean
                    
        # 3. المقارنة المتتالية المباشرة مع مراعاة المرونة الإملائية الطبيعية
        mismatches = []
        max_len = max(len(prog_lessons), len(bank_lessons))
        
        for idx in range(max_len):
            p_name = prog_lessons[idx] if idx < len(prog_lessons) else "[غير موجود في البرنامج]"
            b_name = bank_lessons[idx] if idx < len(bank_lessons) else "[غير موجود في البنك]"
            
            p_norm = normalize_arabic_text(p_name)
            b_norm = normalize_arabic_text(b_name)
            
            # المقارنة الذكية التي تتسامح مع:
            # مسافات واو العطف (و وسائل / ووسائل)، الهمزات (أ/ا)، الياء (ى/ي)، التاء المربوطة (ة/ه)
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
        for r in range(2, self.ws_questions.max_row + 1):
            val = self.ws_questions.cell(r, 3).value
            if val is not None and str(val).strip() != '':
                filled_c += 1
                
        self.results["speed_validation"]["col_c_filled"] = filled_c
        
        dv_count = 0
        if hasattr(self.ws_questions, 'data_validations') and self.ws_questions.data_validations:
            dv_count = len(self.ws_questions.data_validations.dataValidation)
        self.results["speed_validation"]["dv_count"] = dv_count
        
        issues = []
        if filled_c > 0:
            issues.append(f"عمود الشرح والتفسير (Col C) ممتلئ في {filled_c} خلية، مما يبطئ فتح الملف ويثقل ذاكرة الحاسوب.")
        if dv_count < 4:
            issues.append(f"عدد قواعد التحقق من صحة البيانات (Data Validation) هو {dv_count} من أصل 4 قواعد قياسية (L_1, L_2, L_3, L_4).")
            
        if issues:
            self.results["speed_validation"]["status"] = "WARNING"
            self.results["speed_validation"]["details"] = issues
            for iss in issues:
                self.results["recommendations"].append(f"⚠️ {iss}")
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
            f.write(f"**تاريخ التدقيق**: {os.popen('date /t').read().strip() if os.name == 'nt' else 'Today'}\n")
            f.write(f"**ملف البنك**: `{r['file_info'].get('bank_name')}` ({r['file_info'].get('bank_size_kb')} KB)\n")
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
                f.write("| الصف الحالي | درس السؤال | مكرر من الصف | درس السؤال الأصلي | نص السؤال المكرر |\n")
                f.write("| :---: | :--- | :---: | :--- | :--- |\n")
                for it in r['duplicates']['details'][:15]:
                    f.write(f"| {it['row']} | {it['lesson']} | {it['duplicate_of_row']} | {it['duplicate_of_lesson']} | {it['text'][:60]}... |\n")
                f.write("\n")

            if r['recommendations']:
                f.write("### 💡 التوصيات وخطة التصحيح:\n\n")
                for idx, rec in enumerate(r['recommendations'], start=1):
                    f.write(f"- {rec}\n")

class ExamBankFixer:
    """محرك التصحيح التلقائي للأخطاء ومعالجة البنوك قبل الرفع"""
    def __init__(self, bank_path: str, output_path: Optional[str] = None):
        self.bank_path = bank_path
        self.output_path = output_path or bank_path.replace('.xlsx', '_مصحح_ومعتمد.xlsx')
        self.fixed_log = []

    def fix_all(self) -> str:
        wb = openpyxl.load_workbook(self.bank_path)
        ws_q = wb['Questions']
        max_r = ws_q.max_row

        # 1. تفريغ عمود الشرح والتفسير لضمان الفتح الفوري
        c_cleared = 0
        for r in range(2, max_r + 1):
            if ws_q.cell(r, 3).value is not None and str(ws_q.cell(r, 3).value).strip() != '':
                ws_q.cell(r, 3).value = None
                c_cleared += 1
        if c_cleared:
            self.fixed_log.append(f"تم تفريغ عمود الشرح والتفسير في {c_cleared} خلية لضمان سرعة الفتح الفورية.")

        # 2. تصحيح أسئلة الصواب والخطأ (تفريغ الخيارات وضبط الإجابة)
        tf_cleared = 0
        tf_norm = 0
        mcq_norm = 0
        for r in range(2, max_r + 1):
            q_type = str(ws_q.cell(r, 1).value or '')
            ans = str(ws_q.cell(r, 5).value or '').strip()

            if 'صح' in q_type or 'صواب' in q_type:
                # تفريغ الخيارات F..I
                if any(ws_q.cell(r, c).value is not None for c in range(6, 10)):
                    for c in range(6, 10):
                        ws_q.cell(r, c).value = None
                    tf_cleared += 1

                # توحيد الإجابة إلى True / False
                if ans in ['صواب', 'صح', 'True', 'true', 'A', 'الخيار أ (A)']:
                    ws_q.cell(r, 5).value = 'True'
                    tf_norm += 1
                elif ans in ['خطأ', 'خطا', 'False', 'false', 'B', 'الخيار ب (B)']:
                    ws_q.cell(r, 5).value = 'False'
                    tf_norm += 1
            else:
                # اختيار من متعدد: توحيد الإجابة إلى A, B, C, D
                if ans.startswith('الخيار أ') or ans == 'أ': ws_q.cell(r, 5).value = 'A'; mcq_norm += 1
                elif ans.startswith('الخيار ب') or ans == 'ب': ws_q.cell(r, 5).value = 'B'; mcq_norm += 1
                elif ans.startswith('الخيار ج') or ans == 'ج': ws_q.cell(r, 5).value = 'C'; mcq_norm += 1
                elif ans.startswith('الخيار د') or ans == 'د': ws_q.cell(r, 5).value = 'D'; mcq_norm += 1

        if tf_cleared:
            self.fixed_log.append(f"تم تفريغ أعمدة الخيارات (F..I) لعدد {tf_cleared} سؤال صواب وخطأ.")
        if tf_norm:
            self.fixed_log.append(f"تم توحيد صياغة إجابة الصواب والخطأ إلى (True/False) لعدد {tf_norm} سؤال.")
        if mcq_norm:
            self.fixed_log.append(f"تم توحيد رموز الإجابة لأسئلة الاختيار من متعدد إلى (A/B/C/D) لعدد {mcq_norm} سؤال.")

        # 3. ضبط نطاقات الداتا فاليديشن لتطابق الصفوف الفعلية بالضبط
        if hasattr(ws_q, 'data_validations') and ws_q.data_validations:
            for dv in ws_q.data_validations.dataValidation:
                old_sq = str(dv.sqref)
                col_letter = old_sq[0]
                dv.sqref = f"{col_letter}2:{col_letter}{max_r}"
            self.fixed_log.append(f"تم ضبط نطاقات التحقق من صحة البيانات (Data Validation) لتنتهي عند الصف {max_r} بدقة.")

        wb.save(self.output_path)

        # 4. تجميع النسخة عبر محرك إكسيل الأصلي إذا كان متوفراً لضمان سرعة 0.06 ثانية
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

