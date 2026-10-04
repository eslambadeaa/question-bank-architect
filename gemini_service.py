"""
Gemini AI Engine & Prompt Integration Module
Enforces the Strict Single-Row Rule and (1 Lecture : 2 Hours : 25 Questions) Formula.
حقوق الملكية الفكرية وتطوير النظام: إسلام عبد البديع
"""

import json
import re
import os
import time
from typing import List, Dict, Any, Optional, Callable
from pydantic import BaseModel, Field

# Blacklisted terms strictly forbidden in questions and explanations
BLACKLISTED_WORDS = [
    "المعيار",
    "المعيار الأساسي",
    "نمط",
    "صيغة",
    "تقييم",
    "في هذا السياق"
]

DIFFICULTY_LEVELS = ["سهل", "متوسط", "صعب", "صعب جداً", "تفوق"]
QUESTION_TYPES = ["اختيار من متعدد", "صواب وخطأ"]


class QuestionItem(BaseModel):
    id: int = Field(description="الرقم التسلسلي للسؤال يبدأ من 1")
    question_text: str = Field(description="نص السؤال ويجب أن يذكر اسم المعدة صراحة")
    explanation: str = Field(default="", description="الشرح والتفسير الموجز المركز على توضيح الإجابة الصحيحة إن وجد")
    option_a: str = Field(description="الخيار (أ) - في أسئلة الصواب والخطأ يكون 'True'")
    option_b: str = Field(description="الخيار (ب) - في أسئلة الصواب والخطأ يكون 'False'")
    option_c: str = Field(default="", description="الخيار (ج) - فارغ في أسئلة الصواب والخطأ")
    option_d: str = Field(default="", description="الخيار (د) - فارغ في أسئلة الصواب والخطأ")
    correct_answer: str = Field(description="النص الدقيق للإجابة الصحيحة أو 'True'/'False' وليس الحرف")
    difficulty: str = Field(description="سهل / متوسط / صعب / صعب جداً / تفوق")
    lesson: str = Field(description="الموضوع الفني أو المكون الدقيق - ممنوع كتابة كلمة محاضرة")
    question_type: str = Field(description="اختيار من متعدد أو صواب وخطأ")


class QuestionBankResult(BaseModel):
    equipment_name: str
    section_name: str
    total_questions: int
    questions: List[QuestionItem]


class SyllabusItem(BaseModel):
    id: int = Field(description="م - الرقم التسلسلي للسطر يبدأ من 1")
    lesson_name: str = Field(description="عنوان الدرس أو المكون الفني الصريح من المرجع - ممنوع تماماً كتابة محاضرة أو الجزء أو رقم")
    teaching_hours: int = Field(default=2, description="عدد الساعات المخصصة = 2 دائماً لكل سطر")
    lecture_count: int = Field(default=1, description="عدد المحاضرات = 1 دائماً لكل سطر")
    question_count: int = Field(default=25, description="عدد الأسئلة المستهدفة = 25 دائماً لكل سطر")
    reference: str = Field(default="الدليل الفني المعتمد", description="المرجع المعتمد")


class SyllabusResult(BaseModel):
    equipment_name: str
    total_hours: int
    total_lectures: int
    syllabus_items: List[SyllabusItem]


def get_genai_client(api_key: Optional[str] = None):
    """
    Initializes and returns a Google GenAI Client.
    Checks explicit parameter, local saved config, environment variables, or Streamlit secrets.
    """
    from google import genai
    
    key = api_key or os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not key:
        config_path = os.path.join(os.path.dirname(__file__), ".api_key_config")
        if os.path.exists(config_path):
            try:
                with open(config_path, "r", encoding="utf-8") as f:
                    saved = f.read().strip()
                    if saved:
                        key = saved
            except Exception:
                pass
                
    if not key:
        try:
            import streamlit as st
            key = st.secrets.get("GEMINI_API_KEY") or st.secrets.get("GOOGLE_API_KEY")
        except Exception:
            pass
            
    if not key:
        raise ValueError("مفتاح Gemini API غير متوفر! يرجى إدخال المفتاح في الشريط الجانبي أو حفظه.")
        
    from google.genai import types
    return genai.Client(
        api_key=key,
        http_options=types.HttpOptions(timeout=300000)  # 5 minutes timeout to prevent WinError 10060 & ReadTimeout
    )


AVAILABLE_MODELS = [
    "gemini-3.5-flash",
    "gemini-3.5-flash-lite",
    "gemini-flash-lite-latest",
    "gemini-2.5-flash",
]


def generate_content_with_fallback(
    client,
    contents: Any,
    config: Optional[Dict[str, Any]] = None
) -> Any:
    """
    Calls generate_content across AVAILABLE_MODELS in priority order.
    Automatically handles rate limits (429), quota exhaustion, and service unavailability.
    """
    last_exc = None
    for model_name in AVAILABLE_MODELS:
        try:
            response = client.models.generate_content(
                model=model_name,
                contents=contents,
                config=config
            )
            return response
        except Exception as e:
            last_exc = e
            time.sleep(0.5)
            continue
            
    if last_exc:
        raise last_exc


def infer_equipment_name(client, text_sample: str) -> str:
    """Infers the technical equipment or system name from the source text."""
    prompt = f"""
    أنت خبير توثيق عسكري وهندسي رائد. قم بتحليل النص التالي المستخرج من دليل فني واستخرج بدقة الاسم الرسمي والمحدد للمعدة أو النظام الفني الذي يتناوله الدليل.
    
    شروط الاستخراج:
    1. اكتب اسم المعدة باللغة العربية بدقة مع ذكر طرازها أو رمزها إن وجد (مثال: "معدة الرادار ثلاثي الأبعاد TPS-70" أو "محطة تحلية المياه بالتناضح العكسي RO").
    2. أخرج الاسم فقط في سطر واحد بدون أي مقدمات أو شروحات إضافية.
    
    عينة النص:
    \"\"\"{text_sample[:4000]}\"\"\"
    """
    try:
        response = generate_content_with_fallback(
            client=client,
            contents=prompt,
        )
        inferred = response.text.strip().replace('"', '').replace("'", "")
        return inferred if inferred else "المعدة الفنية"
    except Exception:
        return "المعدة الفنية"


def clean_json_response(raw_text: str) -> str:
    """Removes markdown code fences and cleans response text to parse valid JSON."""
    text = raw_text.strip()
    if text.startswith("```"):
        text = re.sub(r"^```(?:json)?\s*", "", text, flags=re.MULTILINE)
        text = re.sub(r"\s*```$", "", text, flags=re.MULTILINE)
    return text.strip()


def sanitize_lesson_name(name: str, fallback_index: int = 1, equipment_name: str = "") -> str:
    """
    Strictly sanitizes lesson titles:
    - Forbids and eliminates: 'محاضرة', 'محاضرات', 'رقم X', 'الجزء 1/2', ranges like '1-2' or 'من 1 إلى 5'.
    - Returns purely the technical component or paragraph title.
    """
    cleaned = str(name).strip().strip('"\'')
    
    # Patterns to completely eliminate
    patterns_to_remove = [
        r"محاضرات\s*\d*[-–—:]*\d*",
        r"محاضرة\s*\d*[-–—:]*\d*",
        r"محاضرة\s*(الأولى|الثانية|الثالثة|الرابعة|الخامسة|السادسة|السابعة|الثامنة|التاسعة|العاشرة|الحادية عشرة|الثانية عشرة)?",
        r"الجزء\s*(الأول|الثاني|الثالث|الرابع|الخامس|\d+)?",
        r"رقم\s*\d+",
        r"^\s*[-–—:–/]\s*",
        r"\bمن\s+\d+\s+إلى\s+\d+\b",
        r"\b\d+[-–—]\d+\b"
    ]
    for pat in patterns_to_remove:
        cleaned = re.sub(pat, "", cleaned, flags=re.IGNORECASE)
        
    cleaned = re.sub(r"\s+", " ", cleaned).strip()
    cleaned = cleaned.strip(":-–— /.")
    
    if not cleaned or len(cleaned) < 3:
        cleaned = f"المكون الفني والتشغيلي {fallback_index}"
        if equipment_name:
            cleaned += f" لـ {equipment_name}"
            
    return cleaned


def validate_and_sanitize_questions(questions: List[Dict[str, Any]], equipment_name: str) -> List[Dict[str, Any]]:
    """
    Applies post-generation sanitation:
    - Removes blacklisted words.
    - Ensures equipment name is explicitly present in question text.
    - Verifies correct answers and option consistency.
    - Strictly sanitizes lesson titles to prevent forbidden words like 'محاضرة' or 'الجزء'.
    """
    sanitized = []
    
    for idx, q in enumerate(questions, start=1):
        q_text = str(q.get("question_text", "")).strip()
        explanation = str(q.get("explanation", "")).strip()
        lesson = str(q.get("lesson", "")).strip()
        q_type = str(q.get("question_type", "")).strip()
        difficulty = str(q.get("difficulty", "متوسط")).strip()
        
        # 1. Filter out blacklisted words
        for bad_word in BLACKLISTED_WORDS:
            q_text = re.sub(rf"\b{re.escape(bad_word)}\b", "", q_text)
            explanation = re.sub(rf"\b{re.escape(bad_word)}\b", "", explanation)
            lesson = re.sub(rf"\b{re.escape(bad_word)}\b", "", lesson)
            
        # 2. Enforce explicit equipment name in question text
        if equipment_name and equipment_name not in q_text:
            if not q_text.startswith("في ") and not q_text.startswith("وفقاً "):
                q_text = f"في معدة {equipment_name}، {q_text}"
            else:
                q_text = f"[{equipment_name}] {q_text}"
                
        # 3. Clean lesson column (strictly forbid 'محاضرة' or 'الجزء')
        lesson = sanitize_lesson_name(lesson, idx, equipment_name)
            
        # 4. Standardize options and question type (True/False and صواب أو خطأ)
        opt_a = str(q.get("option_a", "")).strip()
        opt_b = str(q.get("option_b", "")).strip()
        opt_c = str(q.get("option_c", "")).strip()
        opt_d = str(q.get("option_d", "")).strip()
        raw_correct = str(q.get("correct_answer", "")).strip()
        
        is_tf = (
            "صح" in q_type or "صواب" in q_type or "خطأ" in q_type or
            opt_a.lower() in ["صح", "صواب", "true"] or
            opt_b.lower() in ["خطأ", "false"] or
            (opt_c == "" and opt_d == "" and opt_a != "")
        )
        if is_tf:
            q_type = "صواب أو خطأ"
            opt_a = "صواب"
            opt_b = "خطأ"
            opt_c = ""
            opt_d = ""
            if raw_correct.lower() in ["أ", "a", "صواب", "صح", "true", "1", "الخيار أ (a)", "الخيار أ"]:
                correct_label = "الخيار أ (A)"
            else:
                correct_label = "الخيار ب (B)"
        else:
            q_type = "اختيار من متعدد"
            from excel_builder import get_correct_option_label
            correct_label = get_correct_option_label(raw_correct, opt_a, opt_b, opt_c, opt_d)
                
        if difficulty not in ["سهل", "متوسط", "صعب"]:
            difficulty = "متوسط"
            
        sanitized.append({
            "id": idx,
            "question_text": q_text,
            "explanation": explanation,
            "option_a": opt_a,
            "option_b": opt_b,
            "option_c": opt_c,
            "option_d": opt_d,
            "correct_answer": correct_label,
            "correct_label": correct_label,
            "difficulty": difficulty,
            "lesson": lesson,
            "question_type": q_type
        })
        
    return sanitized


def _generate_single_batch(
    client,
    source_chunk: str,
    equipment_name: str,
    section_name: str,
    batch_size: int = 25,
    batch_index: int = 1,
    total_batches: int = 1,
    target_lesson: Optional[str] = None,
    custom_instructions: str = "",
    allow_repetition: bool = False
) -> List[Dict[str, Any]]:
    # Exact breakdown matching user CSV standard:
    # 25 questions: 15 MCQ (7 Easy, 8 Medium) + 10 T/F (4 Hard, 4 Very Hard, 2 Superior)
    if batch_size == 25:
        mcq_count = 15
        tf_count = 10
    else:
        mcq_count = round(batch_size * 0.6)
        tf_count = batch_size - mcq_count
    
    lesson_clause = f"عنوان الدرس / المكون الفني المستهدف: [{target_lesson}]. يجب أن يكون حقل (lesson) في جميع أسئلة هذه الدفعة هو: \"{target_lesson}\" بالضبط وبدون أي تحريف." if target_lesson else "حقل الدرس (lesson): اكتب الموضوع الفني الدقيق أو المكون التكنولوجي المأخوذ من الدليل. يُمنع منعاً باتاً ومطلقاً استخدام كلمات مثل ('محاضرة'، 'محاضرة 1'، 'الجزء 1/2')."
    
    rep_clause = (
        "7. خيار تكرار وتثبيت الأسئلة (مُفعّل لهذا القسم): يُسمح بإعادة تكرار وصياغة أسئلة مفاهيمية أو تأكيدية مشتركة لترسيخ المفاهيم والمعارف التشغيلية الفنية الأساسية بين الدروس في هذا القسم."
        if allow_repetition
        else "7. منع تكرار الأسئلة نهائياً (صارم): يُحظر تماماً تكرار أي سؤال أو فكرة بأي صيغة، ويجب أن يكون كل سؤال فريداً ومميزاً بنسبة 100% دون أي تطابق."
    )

    prompt = f"""
    أنت كبير مهندسي نظم التعليم والتقييم الفني (Senior LMS Assessment Architect) وخبير رائد في بناء بنوك الأسئلة للمعدات التكنولوجية والعسكرية.
    
    المهمة:
    قم ببناء دفعة أسئلة معيارية متطابقة 100% مع النموذج القياسي لبنوك الأسئلة المعتمدة لقسم: [{section_name}].
    المعدة المستهدفة: [{equipment_name}]
    {lesson_clause}
    الدفعة الحالية: رقم {batch_index} من إجمالي {total_batches} دفعات.
    عدد الأسئلة المطلوب في هذه الدفعة: {batch_size} سؤالاً دقيقاً ومميزاً ومستنبطاً بالكامل من الدليل الفني.
    
    الهيكلة الدقيقة للأسئلة ومستويات الصعوبة وفق النموذج المعتمد:
    1. الأسئلة من 1 إلى {mcq_count} (عدد {mcq_count} سؤالاً):
       - نوع السؤال (question_type): "اختيار من متعدد".
       - أربعة خيارات واضحة وغير مكررة (الخيار أ، الخيار ب، الخيار ج، الخيار د).
       - الإجابة الصحيحة (correct_answer): اسم تسمية الخيار فقط حصراً: "الخيار أ (A)" أو "الخيار ب (B)" أو "الخيار ج (C)" أو "الخيار د (D)".
       - توزيع الصعوبة: الأسئلة من 1 إلى 7 بمستوى "سهل"، والأسئلة من 8 إلى 15 بمستوى "متوسط".
       
    2. الأسئلة من {mcq_count + 1} إلى {batch_size} (عدد {tf_count} أسئلة):
       - نوع السؤال (question_type): "صواب أو خطأ".
       - الخيار أ: "صواب"
       - الخيار ب: "خطأ"
       - الخيار ج: "" (سلسلة فارغة تماماً)
       - الخيار د: "" (سلسلة فارغة تماماً)
       - الإجابة الصحيحة (correct_answer): تكون إما "الخيار أ (A)" أو "الخيار ب (B)".
       - توزيع الصعوبة: مستوى "صعب".
    
    القواعد الصارمة والواجبات الإلزامية:
    1. التأصيل الفني الكامل (100% Grounded): جميع الأسئلة والإجابات والخيارات مستقاة حصراً من النص الفني المرفق دون أي اختلاق.
    2. ذكر اسم المعدة صراحة في نص كل سؤال بدون استثناء (مثال: "في معدة {equipment_name}، ما هو...").
    3. حقل الإجابة الصحيحة (correct_answer): يجب أن يحتوي حصراً وبالتطابق التام على واحد من النصوص التالية فقط: "الخيار أ (A)" أو "الخيار ب (B)" أو "الخيار ج (C)" أو "الخيار د (D)".
    4. الشرح والتفسير (explanation): موجز ومركز جداً يوضح باقتضاب سبب صحة الإجابة، أو اتركه فارغاً.
    5. حقل الدرس (lesson): اكتب عنوان المكون الفني المعتمد.
    6. قائمة الكلمات المحظورة (Blacklisted Words): يُحظر تماماً كتابة أو استخدام أي من الكلمات التالية في أي سؤال أو شرح:
       ["المعيار"، "المعيار الأساسي"، "نمط"، "صيغة"، "تقييم"، "في هذا السياق"].
    {rep_clause}
    
    تعليمات إضافية مخصصة من المستخدم:
    {custom_instructions if custom_instructions else "التزم بالدقة الفنية الشديدة للمصطلحات والمقاييس والأرقام الفنية المذكورة وتغطية كافة جوانب التشغيل والأعطال والأمان."}
    
    النص المرجعي المستخرج من الدليل الفني:
    \"\"\"{source_chunk[:350000]}\"\"\"
    """

    try:
        response = generate_content_with_fallback(
            client=client,
            contents=prompt,
            config={
                "response_mime_type": "application/json",
                "response_schema": QuestionBankResult,
                "temperature": 0.2,
            }
        )
        raw_json = clean_json_response(response.text)
        data = json.loads(raw_json)
        return data.get("questions", [])
    except Exception:
        json_prompt = prompt + "\n\nأعد الإخراج ككائن JSON صالح به المفاتيح: equipment_name, section_name, total_questions, questions."
        response = generate_content_with_fallback(
            client=client,
            contents=json_prompt,
            config={
                "response_mime_type": "application/json",
                "temperature": 0.2,
            }
        )
        raw_json = clean_json_response(response.text)
        data = json.loads(raw_json)
        return data.get("questions", [])


def generate_question_bank(
    client,
    source_text: str,
    equipment_name: str,
    section_name: str = "العام",
    num_questions: int = 25,
    custom_instructions: str = "",
    status_callback: Optional[Callable[[str], None]] = None,
    syllabus_items: Optional[List[Dict[str, Any]]] = None,
    allow_repetition: bool = False
) -> Dict[str, Any]:
    """
    Generates an LMS-compliant Question Bank supporting up to 2500 questions
    via intelligent batching and rate-limit resilient processing.
    Directly aligns questions lesson-by-lesson when syllabus_items are provided.
    """
    all_raw_questions = []
    text_length = len(source_text)
    chunk_size = min(text_length, 350000)
    
    # Mode A: Generating questions lesson-by-lesson matching syllabus items
    if syllabus_items and len(syllabus_items) > 0:
        total_lessons = len(syllabus_items)
        if status_callback:
            status_callback(f"جارٍ بناء بنك أسئلة [{section_name}] لكل درس على حدة ({total_lessons} درساً × 25 سؤالاً)...")
            
        for l_idx, s_item in enumerate(syllabus_items, start=1):
            lesson_title = s_item.get("lesson_name", f"درس {l_idx}")
            if status_callback:
                status_callback(
                    f"جارٍ بناء أسئلة درس ({l_idx}/{total_lessons}): [{lesson_title}] "
                    f"- 25 سؤالاً (15 اختيار من متعدد + 10 صواب وخطأ True/False)..."
                )
            
            if text_length > chunk_size and total_lessons > 1:
                stride = max(1, (text_length - chunk_size) // max(1, total_lessons - 1))
                start_pos = (l_idx - 1) * stride
                end_pos = min(text_length, start_pos + chunk_size)
                source_chunk = source_text[start_pos:end_pos]
            else:
                source_chunk = source_text[:chunk_size]
                
            batch_questions = []
            for attempt in range(3):
                try:
                    batch_questions = _generate_single_batch(
                        client=client,
                        source_chunk=source_chunk,
                        equipment_name=equipment_name,
                        section_name=section_name,
                        batch_size=25,
                        batch_index=l_idx,
                        total_batches=total_lessons,
                        target_lesson=lesson_title,
                        custom_instructions=custom_instructions,
                        allow_repetition=allow_repetition
                    )
                    if batch_questions:
                        break
                except Exception as e:
                    if attempt < 2:
                        time.sleep(2 * (attempt + 1))
                    else:
                        if status_callback:
                            status_callback(f"⚠️ تعذر استخراج أسئلة درس [{lesson_title}]: {str(e)}")
                        
            for q in batch_questions:
                q["lesson"] = lesson_title
                
            all_raw_questions.extend(batch_questions)
            if total_lessons > 2 and l_idx < total_lessons:
                time.sleep(0.4)
                
    # Mode B: Standard batching (when syllabus_items not provided)
    else:
        BATCH_SIZE = 25
        if num_questions <= BATCH_SIZE:
            batch_sizes = [num_questions]
        else:
            full_batches = num_questions // BATCH_SIZE
            remainder = num_questions % BATCH_SIZE
            batch_sizes = [BATCH_SIZE] * full_batches
            if remainder > 0:
                batch_sizes.append(remainder)
                
        total_batches = len(batch_sizes)
        for b_idx, b_size in enumerate(batch_sizes, start=1):
            if status_callback:
                status_callback(
                    f"جارٍ إعداد بنك أسئلة [{section_name}]: دفعة {b_idx} من {total_batches} "
                    f"(تم إنجاز {len(all_raw_questions)} من أصل {num_questions} سؤالاً)..."
                )
            
            if text_length > chunk_size and total_batches > 1:
                stride = max(1, (text_length - chunk_size) // max(1, total_batches - 1))
                start_pos = (b_idx - 1) * stride
                end_pos = min(text_length, start_pos + chunk_size)
                source_chunk = source_text[start_pos:end_pos]
            else:
                source_chunk = source_text[:chunk_size]
                
            batch_questions = []
            for attempt in range(3):
                try:
                    batch_questions = _generate_single_batch(
                        client=client,
                        source_chunk=source_chunk,
                        equipment_name=equipment_name,
                        section_name=section_name,
                        batch_size=b_size,
                        batch_index=b_idx,
                        total_batches=total_batches,
                        custom_instructions=custom_instructions,
                        allow_repetition=allow_repetition
                    )
                    if batch_questions:
                        break
                except Exception as e:
                    if attempt < 2:
                        time.sleep(2 * (attempt + 1))
                    else:
                        if status_callback:
                            status_callback(f"تنبيه: تعذر إكمال الدفعة {b_idx} وسيتم استكمال المعالجة: {str(e)}")
                                
            all_raw_questions.extend(batch_questions)
            if total_batches > 2 and b_idx < total_batches:
                time.sleep(0.4)

    if not all_raw_questions:
        err_msg = f"تعذر توليد أسئلة قسم [{section_name}]. يرجى التحقق من توفر حصة API أو الاتصال بالشبكة."
        if status_callback:
            status_callback(f"❌ {err_msg}")
        raise RuntimeError(err_msg)

    sanitized_questions = validate_and_sanitize_questions(all_raw_questions, equipment_name)
    
    return {
        "equipment_name": equipment_name,
        "section_name": section_name,
        "total_questions": len(sanitized_questions),
        "questions": sanitized_questions
    }


def _generate_single_syllabus_chunk(
    client,
    source_chunk: str,
    equipment_name: str,
    num_lectures: int,
    reference_name: str = "الدليل الفني المعتمد",
    section_name: str = "",
    custom_instructions: str = ""
) -> List[Dict[str, Any]]:
    """Generates a small chunk of syllabus items (<=28) rapidly without risk of timeout."""
    sec_clause = f"لقسم: [{section_name}]." if section_name else ""
    prompt = f"""
    أنت مدير تدريب فني عسكري وأكاديمي (Technical Curriculum Director).
    
    المهمة:
    قم بتحليل الدليل الفني التالي لمعدة: [{equipment_name}] {sec_clause}
    واستخرج بدقة قائمة موضوعات ومكونات فنية بعدد {num_lectures} موضوعاً/درساً فردياً مستقلاً بالضبط.
    
    القواعد الإلزامية الصارمة:
    1. عدد الدروس المطلوب بالضبط: استخرج {num_lectures} سطراً فردياً مستقلاً (يمثل كل سطر درساً فنياً محدداً).
    2. منع الدمج والنطاقات نهائياً:
       - يُمنع منعاً باتاً دمج أكثر من موضوع أو درس في سطر واحد.
       - يُمنع كتابة نطاقات مثل ("محاضرات 1-2" أو "1 إلى 5").
       - كل عنصر يمثل درساً فردياً واحداً ومستقلاً بنسبة 100%.
    3. صياغة خانة "عنوان الدرس / المكون الفني" (lesson_name):
       - اكتب اسم الفقرة أو المكون الفني الصريح المأخوذ من المرجع (مثال: "الفكرة العامة والأغراض وأوضاع الإطلاق"، "منظومة التوجيه والتحكم بالهوائي").
       - يُحظر تماماً كتابة كلمة "محاضرة"، "محاضرات"، "رقم X"، أو "الجزء 1/2" داخل عنوان الدرس.
    4. المرجع المعتمد (reference): اكتب اسم المرجع: "{reference_name}".
    
    5. التغطية الشاملة لكامل المرجع (Comprehensive Full Coverage):
       يجب أن تغطي الدروس كامل المحتوى المرجعي المرفق من أول صفحة إلى آخر صفحة بتوزيع متوازن ومتسلسل منطقياً يشمل كافة الأبواب والفصول والمكونات الفنية لمنظومة [{equipment_name}]، ويُمنع منعاً باتاً الاقتصار على المقدمة أو الفصول الأولى فقط.
    
    تعليمات إضافية:
    {custom_instructions}
    
    النص المرجعي المستخرج:
    \"\"\"{source_chunk[:400000]}\"\"\"
    """

    for attempt in range(3):
        try:
            response = generate_content_with_fallback(
                client=client,
                contents=prompt,
                config={
                    "response_mime_type": "application/json",
                    "response_schema": SyllabusResult,
                    "temperature": 0.2,
                }
            )
            raw_json = clean_json_response(response.text)
            data = json.loads(raw_json)
            raw_items = data.get("syllabus_items", [])
            if raw_items:
                return raw_items
        except Exception:
            if attempt < 2:
                time.sleep(2 * (attempt + 1))
            else:
                try:
                    response = generate_content_with_fallback(
                        client=client,
                        contents=prompt + "\n\nأعد الإخراج ككائن JSON صالح يحتوي على syllabus_items.",
                        config={
                            "response_mime_type": "application/json",
                            "temperature": 0.2,
                        }
                    )
                    raw_json = clean_json_response(response.text)
                    data = json.loads(raw_json)
                    raw_items = data.get("syllabus_items", [])
                    if raw_items:
                        return raw_items
                except Exception:
                    pass
    return []


def generate_syllabus(
    client,
    source_text: str,
    equipment_name: str,
    total_lectures: int = 10,
    total_hours: Optional[int] = None,
    target_questions: Optional[int] = None,
    reference_name: str = "الدليل الفني المعتمد",
    custom_instructions: str = "",
    status_callback: Optional[Callable[[str], None]] = None,
    sections_breakdown: Optional[List[Dict[str, Any]]] = None,
    *args,
    **kwargs
) -> Dict[str, Any]:
    """
    Generates a structured Syllabus & Index adhering strictly to the:
    Strict Single-Row Rule and (1 Lecture : 2 Hours : 25 Questions) Formula.
    Processes in safe chunks covering the full depth of the reference manual.
    """
    all_raw_items = []
    
    # Mode A: If sections_breakdown is provided (Scenario 1)
    if sections_breakdown and len(sections_breakdown) > 0:
        for sec_info in sections_breakdown:
            sec_name = sec_info.get("name", "قسم")
            sec_l = sec_info.get("lectures", 28)
            sec_text = sec_info.get("text", "") or source_text
            sec_ref = sec_info.get("reference", reference_name) or reference_name
            sec_eq = sec_info.get("equipment_name", equipment_name) or equipment_name
            
            CHUNK_MAX = 28
            s_chunks = [CHUNK_MAX] * (sec_l // CHUNK_MAX)
            rem = sec_l % CHUNK_MAX
            if rem > 0:
                s_chunks.append(rem)
                
            sec_len = len(sec_text)
            chunk_chars = min(sec_len, 400000)
            
            for c_idx, c_size in enumerate(s_chunks, start=1):
                if status_callback:
                    status_callback(
                        f"جارٍ إعداد فهرس [{sec_name}] ({len(all_raw_items)} من أصل {total_lectures} درساً)..."
                    )
                if sec_len > chunk_chars and len(s_chunks) > 1:
                    stride = max(1, (sec_len - chunk_chars) // max(1, len(s_chunks) - 1))
                    start_p = (c_idx - 1) * stride
                    end_p = min(sec_len, start_p + chunk_chars)
                    chunk_source = sec_text[start_p:end_p]
                else:
                    chunk_source = sec_text[:chunk_chars]

                chunk_items = _generate_single_syllabus_chunk(
                    client=client,
                    source_chunk=chunk_source,
                    equipment_name=sec_eq,
                    num_lectures=c_size,
                    reference_name=sec_ref,
                    section_name=sec_name,
                    custom_instructions=custom_instructions
                )
                all_raw_items.extend(chunk_items)
                if len(s_chunks) > 1:
                    time.sleep(0.3)
                    
    # Mode B: Standard chunking by max 28 lectures per API call
    else:
        CHUNK_MAX = 28
        if total_lectures <= CHUNK_MAX:
            chunks = [total_lectures]
        else:
            chunks = [CHUNK_MAX] * (total_lectures // CHUNK_MAX)
            rem = total_lectures % CHUNK_MAX
            if rem > 0:
                chunks.append(rem)
                
        text_length = len(source_text)
        chunk_chars = min(text_length, 400000)
        
        for c_idx, c_size in enumerate(chunks, start=1):
            if status_callback:
                status_callback(
                    f"جارٍ إعداد فهرس الدروس: دفعة {c_idx} من {len(chunks)} "
                    f"({len(all_raw_items)} من أصل {total_lectures} درساً)..."
                )
            if text_length > chunk_chars and len(chunks) > 1:
                stride = max(1, (text_length - chunk_chars) // max(1, len(chunks) - 1))
                start_p = (c_idx - 1) * stride
                end_p = min(text_length, start_p + chunk_chars)
                chunk_text = source_text[start_p:end_p]
            else:
                chunk_text = source_text[:chunk_chars]
                
            chunk_items = _generate_single_syllabus_chunk(
                client=client,
                source_chunk=chunk_text,
                equipment_name=equipment_name,
                num_lectures=c_size,
                reference_name=reference_name,
                custom_instructions=custom_instructions
            )
            all_raw_items.extend(chunk_items)
            if len(chunks) > 1:
                time.sleep(0.3)

    # Post-process: Guarantee 100% strict adherence to the (1 : 2 : 25) single-row rule
    clean_items = []
    for idx in range(1, total_lectures + 1):
        if idx - 1 < len(all_raw_items):
            it = all_raw_items[idx - 1]
            raw_title = it.get("lesson_name", "")
            raw_ref = str(it.get("reference", reference_name)).strip() or reference_name
        else:
            raw_title = f"المكون الفني والتطبيقي {idx}"
            raw_ref = reference_name
            
        sanitized_title = sanitize_lesson_name(raw_title, idx, equipment_name)
        
        clean_items.append({
            "id": idx,
            "lesson_name": sanitized_title,
            "teaching_hours": 2,      # STRICTLY 2 hours per row
            "lecture_count": 1,       # STRICTLY 1 lecture per row
            "question_count": 25,     # STRICTLY 25 questions per row
            "reference": raw_ref
        })

    return {
        "equipment_name": equipment_name,
        "total_hours": len(clean_items) * 2,
        "total_lectures": len(clean_items),
        "total_questions": len(clean_items) * 25,
        "syllabus_items": clean_items
    }
