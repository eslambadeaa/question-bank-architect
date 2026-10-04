"""
منظومة بنوك الأسئلة والمناهج التفاعلية من الأدلة الفنية
تطبيق القاعدة الصارمة لمنع دمج المحاضرات وتطبيق معادلة (1 محاضرة : 2 ساعة : 25 سؤال)
تطوير: Eslam Abdelbadea
"""

import os
import io
import time
from datetime import datetime
from typing import List, Dict, Any, Optional
import pandas as pd
import streamlit as st

# Internal modules with live reload to avoid stale cache
import importlib
import file_parsers
import gemini_service
import excel_builder

importlib.reload(file_parsers)
importlib.reload(gemini_service)
importlib.reload(excel_builder)

import zipfile
from file_parsers import extract_source_document, get_pdf_page_count, get_docx_approx_page_count
from gemini_service import (
    get_genai_client,
    infer_equipment_name,
    generate_question_bank,
    generate_syllabus
)
from excel_builder import build_workbook, build_specialty_program_workbook

# -----------------------------------------------------------------------------
# Streamlit Page Configuration
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="منظومة بنوك الأسئلة الفنية | Eslam Abdelbadea",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# Custom RTL & Modern Arabic Styling (Cairo Font, Navy Palette, Cards)
# -----------------------------------------------------------------------------
CUSTOM_CSS = """
<style>
@import url('https://fonts.googleapis.com/css2?family=Cairo:wght@400;600;700;800;900&family=Tajawal:wght@400;500;700&display=swap');

html, body, [class*="css"], .stMarkdown, .stText, div, p, span, h1, h2, h3, h4, h5, h6 {
    font-family: 'Cairo', 'Tajawal', sans-serif !important;
    direction: rtl;
    text-align: right;
}

/* Sidebar Styling */
[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0e1e38 0%, #162a4d 100%) !important;
    color: #ffffff !important;
}

[data-testid="stSidebar"] * {
    color: #f1f5f9 !important;
}

[data-testid="stSidebar"] .stMarkdown h1, 
[data-testid="stSidebar"] .stMarkdown h2, 
[data-testid="stSidebar"] .stMarkdown h3 {
    color: #60a5fa !important;
}

/* Brand Card */
.brand-card {
    background: linear-gradient(135deg, rgba(30, 58, 138, 0.7) 0%, rgba(15, 23, 42, 0.9) 100%);
    border: 1px solid rgba(96, 165, 250, 0.4);
    border-radius: 12px;
    padding: 16px;
    text-align: center;
    box-shadow: 0 4px 15px rgba(0, 0, 0, 0.2);
    margin-bottom: 20px;
}

.brand-badge {
    display: inline-block;
    background: #3b82f6;
    color: #ffffff;
    font-size: 11px;
    font-weight: 700;
    padding: 3px 10px;
    border-radius: 20px;
    margin-bottom: 8px;
    letter-spacing: 0.5px;
}

.brand-title {
    font-size: 15px;
    font-weight: 800;
    color: #ffffff;
    margin: 4px 0;
}

.brand-subtitle {
    font-size: 12px;
    color: #93c5fd;
    margin: 0;
}

/* App Header Card */
.app-header {
    background: linear-gradient(135deg, #1e3a8a 0%, #0f172a 100%);
    border-radius: 16px;
    padding: 24px 30px;
    color: white;
    box-shadow: 0 8px 25px rgba(30, 58, 138, 0.25);
    margin-bottom: 24px;
    border-right: 6px solid #3b82f6;
}

.app-header h1 {
    font-size: 26px;
    font-weight: 900;
    margin: 0 0 8px 0;
    color: #ffffff !important;
}

.app-header p {
    font-size: 14px;
    color: #cbd5e1;
    margin: 0;
}

/* Strict Rule Banner */
.rule-banner {
    background: #eff6ff;
    border: 2px solid #3b82f6;
    border-radius: 10px;
    padding: 14px 18px;
    margin-bottom: 20px;
    color: #1e3a8a;
    font-weight: 700;
}

.stDataFrame {
    direction: rtl !important;
}

.stButton > button {
    border-radius: 8px;
    font-weight: 700;
    transition: all 0.2s ease-in-out;
}
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Persistent API Key Management
# -----------------------------------------------------------------------------
CONFIG_KEY_FILE = os.path.join(os.path.dirname(__file__), ".api_key_config")

def get_saved_key() -> str:
    if os.path.exists(CONFIG_KEY_FILE):
        try:
            with open(CONFIG_KEY_FILE, "r", encoding="utf-8") as f:
                k = f.read().strip()
                if k:
                    return k
        except Exception:
            pass
    env_k = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    return env_k.strip() if env_k else ""

def persist_key(k: str):
    try:
        with open(CONFIG_KEY_FILE, "w", encoding="utf-8") as f:
            f.write(k.strip())
        os.environ["GEMINI_API_KEY"] = k.strip()
    except Exception:
        pass

def remove_persisted_key():
    if os.path.exists(CONFIG_KEY_FILE):
        try:
            os.remove(CONFIG_KEY_FILE)
        except Exception:
            pass
    if "GEMINI_API_KEY" in os.environ:
        del os.environ["GEMINI_API_KEY"]

# -----------------------------------------------------------------------------
# Initialize Session State
# -----------------------------------------------------------------------------
if "inferred_equipment" not in st.session_state:
    st.session_state.inferred_equipment = ""
if "extracted_sources" not in st.session_state:
    st.session_state.extracted_sources = {}
if "generation_results" not in st.session_state:
    st.session_state.generation_results = None
if "excel_buffer" not in st.session_state:
    st.session_state.excel_buffer = None
if "current_api_key" not in st.session_state:
    st.session_state.current_api_key = get_saved_key()

# -----------------------------------------------------------------------------
# Sidebar: Branding & Configuration
# -----------------------------------------------------------------------------
with st.sidebar:
    st.markdown(
        f"""
        <div class="brand-card">
            <div style="font-size: 13px; font-weight: 700; color: #93c5fd; margin-bottom: 3px;">تطوير المنظومة</div>
            <div style="font-size: 16px; font-weight: 800; color: #ffffff;">Eslam Abdelbadea</div>
            <div style="font-size: 11px; color: #cbd5e1; margin-top: 4px;">جميع حقوق التطوير محفوظة ©</div>
        </div>
        """,
        unsafe_allow_html=True
    )
    
    st.header("⚙️ إعدادات النظام")
    
    # API Key Input & Persistence Actions
    api_key = st.text_input(
        "مفتاح Google Gemini API:",
        value=st.session_state.current_api_key,
        type="password",
        help="أدخل مفتاح Gemini API الخاص بك لتشغيل النموذج."
    )
    
    col_save_k, col_clear_k = st.columns([1, 1])
    with col_save_k:
        if st.button("💾 حفظ المفتاح", use_container_width=True, help="حفظ المفتاح محلياً لعدم كتابته كل مرة"):
            if api_key.strip():
                persist_key(api_key.strip())
                st.session_state.current_api_key = api_key.strip()
                st.toast("✅ تم حفظ مفتاح الـ API بنجاح في النظام!", icon="💾")
                st.success("تم الحفظ محلياً!")
            else:
                st.warning("يرجى كتابة المفتاح أولاً لحفظه.")
                
    with col_clear_k:
        if st.button("🗑️ حذف المحفوظ", use_container_width=True, help="حذف المفتاح المحفوظ محلياً"):
            remove_persisted_key()
            st.session_state.current_api_key = ""
            st.toast("🗑️ تم حذف المفتاح المحفوظ بنجاح.", icon="🗑️")
            st.rerun()

    if os.path.exists(CONFIG_KEY_FILE):
        st.caption("🔒 تم التعرف على مفتاح محفوظ تلقائياً وجاهز للاستخدام.")
    
    model_name = st.selectbox(
        "النموذج الذكي المعتمد:",
        options=["gemini-3.5-flash", "gemini-3.5-flash-lite", "gemini-2.5-flash"],
        index=0,
        help="نموذج gemini-3.5-flash يوفر أقصى سرعة ودقة متطورة لبناء بنوك الأسئلة والمناهج."
    )
    
    st.divider()
    
    st.markdown(
        r"""
        ### 📐 المعيار القياسي المعتمد:
        - ⚖️ **المعادلة الفردية (1 : 2 : 25)**:
          * 1 محاضرة فردية مستقلة
          * 2 ساعتان تدريسيتان زوجيتان
          * 25 سؤالاً مستهدفاً لكل سطر
        - 📊 **التوزيع القياسي للأقسام**:
          * 📘 **القسم الإعدادي**: 28 م (56 س | 700 سؤال)
          * 📙 **القسم المتوسط**: 56 م (112 س | 1,400 سؤال)
          * 📕 **القسم النهائي**: 40 م (80 س | 1,000 سؤال)
          * 🎯 **الإجمالي**: 124 م (248 س | 3,100 سؤال)
        - 🚫 **منع النطاقات والدمج نهائياً**: كل سطر يمثل درساً فردياً مستقلاً 100%.
        - 🏷️ **صياغة عنوان الدرس**: اسم المكون الفني الصريح دون كلمات (محاضرة، رقم، جزء).
        """
    )
    st.caption("الإصدار 2.5 - إنتاجي معتمد")

# -----------------------------------------------------------------------------
# Header Banner
# -----------------------------------------------------------------------------
st.markdown(
    f"""
    <div class="app-header">
        <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
            <div>
                <h1>منظومة بنوك الأسئلة والجدول الزمني للأدلة الفنية</h1>
                <p>توليد الجداول الزمنية المعيارية (1 : 2 : 25) وبنوك الأسئلة المتوافقة مع أنظمة إدارة التعلم LMS</p>
            </div>
            <div style="text-align: left; background: rgba(255,255,255,0.1); padding: 8px 16px; border-radius: 8px;">
                <span style="font-size: 12px; color: #93c5fd;">المطور والمشرف العام:</span><br>
                <b style="font-size: 15px; color: #fef08a;">إسلام عبد البديع</b>
            </div>
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

# Rule Explanation Notice Banner
st.markdown(
    """
    <div class="rule-banner">
        ⚡ <b>المعيار القياسي المعتمد لتوزيع الساعات والمحاضرات وبنوك الأسئلة (Standard Rule):</b><br>
        • <b>المعادلة الفردية الصارمة (1 : 2 : 25):</b> كل سطر = 1 محاضرة فردية : 2 ساعتان تدريسيتان زوجيتان صحيحتان : 25 سؤالاً دقيقاً (يُمنع الدمج أو النطاقات نهائياً).<br>
        • <b>معيار توزيع الأقسام القياسي:</b><br>
        &nbsp;&nbsp;▫️ <b>القسم الإعدادي:</b> 28 محاضرة فردية = 56 ساعة تدريسية = 700 سؤال.<br>
        &nbsp;&nbsp;▫️ <b>القسم المتوسط:</b> 56 محاضرة فردية = 112 ساعة تدريسية = 1,400 سؤال.<br>
        &nbsp;&nbsp;▫️ <b>القسم النهائي:</b> 40 محاضرة فردية = 80 ساعة تدريسية = 1,000 سؤال.<br>
        &nbsp;&nbsp;🎯 <b>إجمالي المنظومة الكاملة:</b> 124 محاضرة فردية = 248 ساعة تدريسية = 3,100 سؤال.
    </div>
    """,
    unsafe_allow_html=True
)

# -----------------------------------------------------------------------------
# -----------------------------------------------------------------------------
# Section 1: Workflow Scenario Selection
# -----------------------------------------------------------------------------
st.subheader("🎯 1. اختيار سيناريو العمل")

scenario_choice = st.radio(
    "حدد وضع التوليد المطلوب:",
    options=[
        "السيناريو 4: منظومة التدريب المتكاملة (برامج التدريب المعتمدة + بنوك الأسئلة الخمسة + برنامج تدريب تخصص)",
        "السيناريو 1: توليد شامل للأقسام الثلاثة مع فهرس الدروس القياسي (1 : 2 : 25)",
        "السيناريو 2: توليد مخصص لقسم أو دورة محددة فقط",
        "السيناريو 3: توليد الجدول الزمني وفهرس الدروس فقط"
    ],
    index=0
)

# -----------------------------------------------------------------------------
# Section 2: Upload Documents
# -----------------------------------------------------------------------------
st.subheader("📁 2. رفع الأدلة والمراجع الفنية")

uploaded_files = st.file_uploader(
    "قم برفع ملفات المراجع الفنية (PDF أو DOCX):",
    type=["pdf", "docx"],
    accept_multiple_files=True,
    help="يمكنك رفع مراجع متعددة وتخصيص كل مرجع لأي قسم وتحديد نطاق صفحاته بشكل منفصل."
)

uploaded_docs_data = {}
if uploaded_files:
    st.write("##### 📄 مراجعة وتأكيد عدد صفحات المراجع المرفوعة:")
    st.caption("يقوم النظام بالتعرف التلقائي الذكي على عدد الصفحات من بيانات الملف، ويمكنك تعديل أو تأكيد عدد الصفحات الفعلي لكل مرجع لضمان شمول المنهج كاملاً وبدون أي اجتزاء:")
    
    for f in uploaded_files:
        f_bytes = f.getvalue()
        f_name = f.name
        detected_p = 1
        if f_name.lower().endswith(".pdf"):
            try:
                detected_p = get_pdf_page_count(f_bytes)
            except Exception:
                detected_p = 1
        elif f_name.lower().endswith((".docx", ".doc")):
            try:
                detected_p = get_docx_approx_page_count(f_bytes)
            except Exception:
                detected_p = 1
                
        session_key = f"file_pages_override_{f_name}"
        if session_key in st.session_state and st.session_state[session_key] < detected_p:
            st.session_state[session_key] = detected_p
            
        with st.container(border=True):
            col_doc_info, col_doc_pages = st.columns([2, 1])
            with col_doc_info:
                file_size_kb = round(len(f_bytes) / 1024, 1)
                st.markdown(f"📖 **{f_name}** &nbsp; `({file_size_kb} KB)`")
                st.caption(f"🔍 تم الكشف التلقائي المحدث: **{detected_p} صفحة**.")
            with col_doc_pages:
                user_pages = st.number_input(
                    f"إجمالي عدد الصفحات المعتمد لـ ({f_name}):",
                    min_value=1,
                    max_value=10000,
                    value=detected_p,
                    key=session_key,
                    help="عدّل هذا الرقم إذا كان المرجع الفعلي يحتوي على صفحات أكثر أو إذا أردت تضمين نطاق أوسع."
                )
        uploaded_docs_data[f_name] = {
            "bytes": f_bytes,
            "total_pages": user_pages
        }
    
    st.success(f"📑 تم تجهيز وتأكيد **{len(uploaded_docs_data)}** مستند(ات) مرجعي(ة). يمكنك الآن تخصيص المراجع ونطاق الصفحات وخيارات التكرار لكل قسم في الخطوة التالية.")

# -----------------------------------------------------------------------------
# Section 3: Equipment & Section Customization (Source, Pages, Repetition)
# -----------------------------------------------------------------------------
st.subheader("⚙️ 3. تخصيص الأقسام والمراجع وخيارات التكرار واسم المعدة لكل قسم")


def render_section_equipment_selector(
    sec_key: str,
    sec_label: str,
    sec_files_map: Dict[str, str],
    default_name: str = "رادار المراقبة الجوية والملاحة"
) -> str:
    """
    Renders equipment name input for a specific section with its own dedicated auto-infer button:
    - Auto-infer button analyzes the specific files selected for this section.
    - User can manually edit or write any custom equipment name for this section.
    """
    inferred_key = f"eq_inferred_{sec_key}"
    input_key = f"eq_input_{sec_key}"
    
    if inferred_key not in st.session_state:
        st.session_state[inferred_key] = ""
        
    st.markdown(f"🏷️ **اسم المعدة أو المنظومة الفنية المعتمد لـ [{sec_label}]:**")
    col_infer_btn, col_eq_field = st.columns([1, 2])
    
    with col_infer_btn:
        st.markdown("<div style='height: 4px;'></div>", unsafe_allow_html=True)
        has_files = len(sec_files_map) > 0
        infer_clicked = st.button(
            f"🔍 استنتاج معدة {sec_label}",
            key=f"btn_infer_{sec_key}",
            disabled=not has_files,
            help=f"استنتاج اسم المعدة الفنية تلقائياً بناءً على المراجع المحددة لـ {sec_label} فقط."
        )
        if infer_clicked:
            if not api_key:
                st.error("يرجى إدخال مفتاح Gemini API أولاً في الشريط الجانبي!")
            else:
                try:
                    with st.spinner(f"جارٍ تحليل مراجع {sec_label} لاستنتاج اسم المعدة..."):
                        client = get_genai_client(api_key)
                        first_f_name = list(sec_files_map.keys())[0]
                        first_f_bytes = uploaded_docs_data[first_f_name]["bytes"]
                        sample_txt, _ = extract_source_document(first_f_name, first_f_bytes, "All")
                        inferred_res = infer_equipment_name(client, sample_txt[:20000])
                        st.session_state[inferred_key] = inferred_res
                        st.session_state[input_key] = inferred_res
                        st.toast(f"✅ تم استنتاج اسم المعدة لـ {sec_label}: {inferred_res}", icon="🔍")
                        st.rerun()
                except Exception as e:
                    st.error(f"خطأ أثناء استنتاج اسم المعدة: {str(e)}")
                    
    with col_eq_field:
        current_eq_val = st.session_state.get(inferred_key) or default_name
        sec_eq_val = st.text_input(
            f"اسم المعدة لـ [{sec_label}] (المعتمد في نص الأسئلة والفهرس):",
            value=current_eq_val,
            key=input_key,
            help=f"سيتم ذكر هذا الاسم صراحة في جميع أسئلة وفهرس {sec_label}."
        )
        
    return sec_eq_val.strip() if sec_eq_val.strip() else default_name


def render_section_source_selector(sec_key: str, sec_label: str):
    """
    Renders flexible manual and page selection for a section:
    - تخصيص مرجع كامل
    - أو تحديد عدد معين من الصفحات الأولى
    - أو اختيار صفحات محددة من مراجع متعددة
    """
    st.markdown(f"###### 📚 المراجع والصفحات المعتمدة لـ ({sec_label}):")
    st.caption("يمكنك تخصيص مرجع كامل، أو تحديد عدد معين من الصفحات الأولى، أو اختيار صفحات محددة من مراجع متعددة:")
    
    sec_files_config = {}
    if not uploaded_docs_data:
        st.warning("يرجى رفع ملفات المراجع في الخطوة 2 أعلاه للتمكن من تخصيصها.")
        return sec_files_config

    for f_idx, (f_name, f_info) in enumerate(uploaded_docs_data.items()):
        total_p = f_info["total_pages"]
        with st.container(border=True):
            col_inc, col_detail = st.columns([1, 2])
            with col_inc:
                inc = st.checkbox(
                    f"📄 {f_name}",
                    value=True,
                    key=f"inc_{sec_key}_{f_name}",
                    help=f"إجمالي الصفحات المعتمدة: {total_p}"
                )
                st.caption(f"إجمالي الصفحات المعتمدة: **{total_p} صفحة**")
            
            if inc:
                with col_detail:
                    p_mode = st.radio(
                        "خيارات الصفحات المستخرجة من هذا المرجع:",
                        options=["كامل المرجع (جميع الصفحات)", "تحديد عدد معين من الصفحات الأولى", "تحديد صفحات أو نطاقات مخصصة"],
                        key=f"pmode_{sec_key}_{f_name}",
                        horizontal=True
                    )
                    
                    if p_mode == "كامل المرجع (جميع الصفحات)":
                        final_range = "All"
                        st.caption(f"✅ سيتم استخراج كامل محتوى المرجع بنسبة 100% (إجمالي {total_p} صفحة).")
                    elif p_mode == "تحديد عدد معين من الصفحات الأولى":
                        c_p = st.number_input(
                            f"عدد الصفحات الأولى من {f_name} (من 1 إلى N):",
                            min_value=1,
                            max_value=max(total_p, 5000),
                            value=min(total_p, 30),
                            key=f"cnt_{sec_key}_{f_name}"
                        )
                        final_range = f"1-{c_p}"
                    else:
                        cust_r = st.text_input(
                            f"أدخل أرقام أو نطاقات الصفحات لـ ({f_name}):",
                            value=f"1-{min(total_p, 25)}",
                            key=f"cust_{sec_key}_{f_name}",
                            help=f"أمثلة: 1-{min(total_p, 15)}, 20-{min(total_p, 35)} أو 5, 8, 12-20"
                        )
                        final_range = cust_r.strip() if cust_r.strip() else "All"
                        
                    sec_files_config[f_name] = final_range

    return sec_files_config


# Structure configuration based on chosen scenario
sections_setup = []

if "السيناريو 4" in scenario_choice:
    st.info("📋 **منظومة التدريب المتكاملة (السيناريو 4):** توليد برامج التدريب للأقسام الثلاثة بهيكل الفترات اليومية (الإعدادي فترتان، المتوسط والنهائي 4 فترات)، وبنوك الأسئلة الخمسة المعتمدة بنظام تسمية الخيارات وبيانات التحقق (Data Validation)، ومصنف برنامج تدريب تخصص المعتمد.")
    
    col_scope1, col_scope2 = st.columns([1, 1])
    with col_scope1:
        s4_gen_mode = st.radio(
            "نطاق التوليد المطلوب للمنظومة:",
            options=[
                "🚀 التوليد المتكامل الشامل (برامج التدريب + بنوك الأسئلة الخمسة 2750 سؤالاً + برنامج تدريب تخصص)",
                "⚡ التوليد السريع لبرامج التدريب وبرنامج تدريب تخصص فقط (استخراج الموضوعات والساعات بدون توليد بنوك الأسئلة)"
            ],
            index=0,
            help="اختر التوليد السريع لاعتماد وتنسيق خطة التدريب وبرنامج تدريب تخصص في ثوانٍ للاختبار السريع."
        )

    st.write("##### 🎛️ تخصيص الأقسام الثلاثة والمراجع الفنية والمعدات:")
    tab_prep, tab_med, tab_fin = st.tabs(["📘 القسم الإعدادي (56 س)", "📙 القسم المتوسط (208 س)", "📕 القسم النهائي (176 س)"])
    
    with tab_prep:
        st.caption("🔹 **القسم الإعدادي:** 14 موضوعاً = 56 ساعة (28 نظري + 28 عملي) | فترتان في اليوم (ف1 نظري، ف2 عملي) | امتحان منتصف الترم بعد اليوم السابع | بنك أسئلة: 350 سؤالاً.")
        col_lp, col_rp = st.columns([1, 1])
        with col_lp:
            lec_prep = st.number_input("عدد موضوعات القسم الإعدادي:", min_value=1, max_value=100, value=14, key="l_prep_s4")
            st.caption(f"⏱️ الساعات: **{lec_prep * 4} ساعة** (28 س نظري + 28 س عملي) | 📝 بنك الأسئلة: **{lec_prep * 25} سؤالاً**")
        with col_rp:
            allow_rep_prep = st.checkbox("🔄 تفعيل خيار تكرار الأسئلة في الإعدادي", value=False, key="rep_prep_s4")
        files_prep = render_section_source_selector("prep_s4", "القسم الإعدادي")
        eq_prep = render_section_equipment_selector("prep_s4", "القسم الإعدادي", files_prep, "معدة التخصص الفني")
        sections_setup.append({
            "name": "القسم الإعدادي",
            "lectures": lec_prep,
            "questions_count": lec_prep * 25,
            "allow_repetition": allow_rep_prep,
            "files": files_prep,
            "equipment_name": eq_prep,
            "terms": [
                {
                    "term_name": "الترم الأول",
                    "lectures": lec_prep,
                    "periods_per_day": 2,
                    "midterm_day": max(1, lec_prep // 2)
                }
            ]
        })
        
    with tab_med:
        st.caption("🔹 **القسم المتوسط:** 52 موضوعاً = 208 ساعات (104 نظري + 104 عملي) | 4 فترات في اليوم: الترم الأول (24 موضوعاً = 96 س - 600 سؤال)، الترم الثاني (28 موضوعاً = 112 س - 700 سؤال).")
        col_lm1, col_lm2 = st.columns([1, 1])
        with col_lm1:
            lec_med_t1 = st.number_input("موضوعات الترم الأول (المتوسط):", min_value=1, max_value=100, value=24, key="l_med_t1")
            lec_med_t2 = st.number_input("موضوعات الترم الثاني (المتوسط):", min_value=1, max_value=100, value=28, key="l_med_t2")
            total_med_lec = lec_med_t1 + lec_med_t2
            st.caption(f"⏱️ إجمالي ساعات المتوسط: **{total_med_lec * 4} ساعة** | 📝 بنك الأسئلة: **{total_med_lec * 25} سؤالاً** (بنكان)")
        with col_lm2:
            allow_rep_med = st.checkbox("🔄 تفعيل خيار تكرار الأسئلة في المتوسط", value=False, key="rep_med_s4")
        files_med = render_section_source_selector("med_s4", "القسم المتوسط")
        eq_med = render_section_equipment_selector("med_s4", "القسم المتوسط", files_med, "معدة التخصص الفني")
        sections_setup.append({
            "name": "القسم المتوسط",
            "lectures": total_med_lec,
            "questions_count": total_med_lec * 25,
            "allow_repetition": allow_rep_med,
            "files": files_med,
            "equipment_name": eq_med,
            "terms": [
                {
                    "term_name": "الترم الأول",
                    "lectures": lec_med_t1,
                    "periods_per_day": 4,
                    "midterm_day": 6
                },
                {
                    "term_name": "الترم الثاني",
                    "lectures": lec_med_t2,
                    "periods_per_day": 4,
                    "midterm_day": 7
                }
            ]
        })
        
    with tab_fin:
        st.caption("🔹 **القسم النهائي:** 44 موضوعاً = 176 ساعة (88 نظري + 88 عملي) | 4 فترات في اليوم: الترم الأول (24 موضوعاً = 96 س - 600 سؤال)، الترم الثاني (20 موضوعاً = 80 س - 500 سؤال).")
        col_lf1, col_lf2 = st.columns([1, 1])
        with col_lf1:
            lec_fin_t1 = st.number_input("موضوعات الترم الأول (النهائي):", min_value=1, max_value=100, value=24, key="l_fin_t1")
            lec_fin_t2 = st.number_input("موضوعات الترم الثاني (النهائي):", min_value=1, max_value=100, value=20, key="l_fin_t2")
            total_fin_lec = lec_fin_t1 + lec_fin_t2
            st.caption(f"⏱️ إجمالي ساعات النهائي: **{total_fin_lec * 4} ساعة** | 📝 بنك الأسئلة: **{total_fin_lec * 25} سؤالاً** (بنكان)")
        with col_lf2:
            allow_rep_fin = st.checkbox("🔄 تفعيل خيار تكرار الأسئلة في النهائي", value=False, key="rep_fin_s4")
        files_fin = render_section_source_selector("fin_s4", "القسم النهائي")
        eq_fin = render_section_equipment_selector("fin_s4", "القسم النهائي", files_fin, "معدة التخصص الفني")
        sections_setup.append({
            "name": "القسم النهائي",
            "lectures": total_fin_lec,
            "questions_count": total_fin_lec * 25,
            "allow_repetition": allow_rep_fin,
            "files": files_fin,
            "equipment_name": eq_fin,
            "terms": [
                {
                    "term_name": "الترم الأول",
                    "lectures": lec_fin_t1,
                    "periods_per_day": 4,
                    "midterm_day": 6
                },
                {
                    "term_name": "الترم الثاني",
                    "lectures": lec_fin_t2,
                    "periods_per_day": 4,
                    "midterm_day": 5
                }
            ]
        })

elif "السيناريو 1" in scenario_choice:
    st.write("##### 🎛️ تخصيص الأقسام الثلاثة والمراجع وخيارات تكرار الأسئلة (فهرس قياسي 1 : 2 : 25):")
    tab_prep, tab_med, tab_fin = st.tabs(["📘 القسم الإعدادي", "📙 القسم المتوسط", "📕 القسم النهائي"])
    
    with tab_prep:
        col_lp, col_rp = st.columns([1, 1])
        with col_lp:
            lec_prep = st.number_input("عدد المحاضرات (السطور الفردية):", min_value=1, max_value=300, value=28, key="l_prep_s1")
            hours_prep = lec_prep * 2
            q_count_prep = lec_prep * 25
            st.caption(f"⏱️ الساعات المخصصة: **{hours_prep} ساعة** (معيار قياسي) | 📝 بنك الأسئلة: **{q_count_prep} سؤالاً** (قاعدة 1 : 2 : 25)")
        with col_rp:
            allow_rep_prep = st.checkbox("🔄 تفعيل خيار تكرار الأسئلة في القسم الإعدادي", value=False, key="rep_prep_s1")
        files_prep = render_section_source_selector("prep_s1", "القسم الإعدادي")
        eq_prep = render_section_equipment_selector("prep_s1", "القسم الإعدادي", files_prep, "رادار المراقبة الجوية والملاحة")
        sections_setup.append({
            "name": "القسم الإعدادي",
            "lectures": lec_prep,
            "questions_count": q_count_prep,
            "allow_repetition": allow_rep_prep,
            "files": files_prep,
            "equipment_name": eq_prep
        })
        
    with tab_med:
        col_lm, col_rm = st.columns([1, 1])
        with col_lm:
            lec_med = st.number_input("عدد المحاضرات (السطور الفردية):", min_value=1, max_value=300, value=56, key="l_med_s1")
            hours_med = lec_med * 2
            q_count_med = lec_med * 25
            st.caption(f"⏱️ الساعات المخصصة: **{hours_med} ساعة** (معيار قياسي) | 📝 بنك الأسئلة: **{q_count_med} سؤالاً** (قاعدة 1 : 2 : 25)")
        with col_rm:
            allow_rep_med = st.checkbox("🔄 تفعيل خيار تكرار الأسئلة في القسم المتوسط", value=False, key="rep_med_s1")
        files_med = render_section_source_selector("med_s1", "القسم المتوسط")
        eq_med = render_section_equipment_selector("med_s1", "القسم المتوسط", files_med, "رادار المراقبة الجوية والملاحة")
        sections_setup.append({
            "name": "القسم المتوسط",
            "lectures": lec_med,
            "questions_count": q_count_med,
            "allow_repetition": allow_rep_med,
            "files": files_med,
            "equipment_name": eq_med
        })
        
    with tab_fin:
        col_lf, col_rf = st.columns([1, 1])
        with col_lf:
            lec_fin = st.number_input("عدد المحاضرات (السطور الفردية):", min_value=1, max_value=300, value=40, key="l_fin_s1")
            hours_fin = lec_fin * 2
            q_count_fin = lec_fin * 25
            st.caption(f"⏱️ الساعات المخصصة: **{hours_fin} ساعة** (معيار قياسي) | 📝 بنك الأسئلة: **{q_count_fin} سؤالاً** (قاعدة 1 : 2 : 25)")
        with col_rf:
            allow_rep_fin = st.checkbox("🔄 تفعيل خيار تكرار الأسئلة في القسم النهائي", value=False, key="rep_fin_s1")
        files_fin = render_section_source_selector("fin_s1", "القسم النهائي")
        eq_fin = render_section_equipment_selector("fin_s1", "القسم النهائي", files_fin, "رادار المراقبة الجوية والملاحة")
        sections_setup.append({
            "name": "القسم النهائي",
            "lectures": lec_fin,
            "questions_count": q_count_fin,
            "allow_repetition": allow_rep_fin,
            "files": files_fin,
            "equipment_name": eq_fin
        })

elif "السيناريو 2" in scenario_choice:
    st.write("##### 🎛️ محددات القسم المستهدف والمراجع وخيار التكرار:")
    col_s2_name, col_s2_l = st.columns([2, 1])
    with col_s2_name:
        custom_sec_name = st.text_input("اسم القسم المستهدف:", value="القسم التخصصي")
    with col_s2_l:
        custom_sec_l = st.number_input("عدد المحاضرات (السطور الفردية):", min_value=1, max_value=300, value=28)
        custom_sec_h = custom_sec_l * 2
        custom_sec_q = custom_sec_l * 25
        st.caption(f"⏱️ الساعات: **{custom_sec_h} ساعة** | 📝 الأسئلة: **{custom_sec_q} سؤالاً** (قاعدة 1 : 2 : 25)")
        
    custom_sec_rep = st.checkbox(
        f"🔄 تفعيل خيار تكرار الأسئلة في [{custom_sec_name}]",
        value=False,
        key="rep_s2",
        help="يسمح بتكرار وصياغة أسئلة مفاهيمية لترسيخ المفاهيم الأساسية في هذا القسم."
    )
    files_s2 = render_section_source_selector("s2", custom_sec_name)
    eq_s2 = render_section_equipment_selector("s2", custom_sec_name, files_s2, "المنظومة التخصصية")
    sections_setup.append({
        "name": custom_sec_name,
        "lectures": custom_sec_l,
        "questions_count": custom_sec_q,
        "allow_repetition": custom_sec_rep,
        "files": files_s2,
        "equipment_name": eq_s2
    })

elif "السيناريو 3" in scenario_choice:
    st.write("##### 🎛️ محددات الجدول الزمني وفهرس الدروس:")
    s3_lectures = st.number_input("إجمالي عدد المحاضرات (عدد السطور الفردية في الفهرس):", min_value=1, max_value=300, value=124)
    s3_hours = s3_lectures * 2
    s3_target_q = s3_lectures * 25
    st.caption(f"⏱️ إجمالي الساعات: **{s3_hours} ساعة** (معيار قياسي) | 📝 إجمالي الأسئلة المخططة: **{s3_target_q} سؤالاً** (بمعدل 25 سؤالاً لكل سطر)")
    
    files_s3 = render_section_source_selector("s3", "الجدول الزمني وفهرس الدروس")
    eq_s3 = render_section_equipment_selector("s3", "الجدول الزمني", files_s3, "المنظومة الفنية المعتمدة")
    sections_setup.append({
        "name": "الجدول الزمني وفهرس الدروس",
        "lectures": s3_lectures,
        "questions_count": s3_target_q,
        "allow_repetition": False,
        "files": files_s3,
        "equipment_name": eq_s3
    })

# Custom Injected Instructions
custom_prompt_rules = st.text_area(
    "💡 توجيهات وتوصيات إضافية لمحرك الذكاء الاصطناعي (اختياري):",
    placeholder="مثال: التركيز على استخراج مسميات المكونات الدقيقة للمعدة كعناوين للدروس، والتركيز على إجراءات الفحص الدوري والمواصفات الرقمية...",
    height=80
)

# -----------------------------------------------------------------------------
# Section 4: Execution Engine
# -----------------------------------------------------------------------------
st.divider()

col_btn, col_info = st.columns([1, 2])
with col_btn:
    generate_clicked = st.button(
        "🚀 بدء توليد الفهرس الفردي وبنك الأسئلة",
        type="primary",
        use_container_width=True,
        disabled=len(uploaded_docs_data) == 0
    )
with col_info:
    if len(uploaded_docs_data) == 0:
        st.info("قم برفع ملف مرجعي واحد على الأقل للمتابعة.")
    else:
        st.caption(f"تم تجهيز {len(uploaded_docs_data)} مستند(ات) للمعالجة وتوزيعها على الأقسام المحددة.")

if generate_clicked:
    if not api_key:
        st.error("مفتاح Gemini API مطلوب! يرجى إدخاله أو حفظه في الشريط الجانبي.")
        st.stop()
        
    missing_eq = [s["name"] for s in sections_setup if not s.get("equipment_name", "").strip()]
    if missing_eq:
        st.error(f"يرجى تحديد أو استنتاج اسم المعدة للأقسام التالية: {', '.join(missing_eq)}")
        st.stop()
        
    primary_equipment_name = sections_setup[0]["equipment_name"] if sections_setup else "المعدة الفنية"

    client = get_genai_client(api_key)
    progress_bar = st.progress(0)
    status_text = st.empty()

    try:
        # Step 1: Text extraction per section
        status_text.text("1/4: جارٍ استخراج وتصفية النصوص من الصفحات والمراجع المحددة لكل قسم...")
        progress_bar.progress(15)

        all_extracted_sections_text = {}
        all_refs_by_section = {}
        combined_all_corpus_list = []

        for sec in sections_setup:
            sec_name = sec["name"]
            sec_texts = []
            sec_refs = []
            for f_name, page_range in sec["files"].items():
                if f_name in uploaded_docs_data:
                    f_bytes = uploaded_docs_data[f_name]["bytes"]
                    txt, meta = extract_source_document(f_name, f_bytes, page_range)
                    if txt.strip():
                        sec_texts.append(f"=== مرجع: {f_name} (صفحات: {page_range}) ===\n{txt}")
                        sec_refs.append(f_name)
            
            all_extracted_sections_text[sec_name] = "\n\n".join(sec_texts)
            all_refs_by_section[sec_name] = sec_refs
            if sec_texts:
                combined_all_corpus_list.extend(sec_texts)

        # Fallback if any section had no files selected: extract all uploaded files
        if not combined_all_corpus_list:
            for f_name, f_info in uploaded_docs_data.items():
                txt, _ = extract_source_document(f_name, f_info["bytes"], "All")
                combined_all_corpus_list.append(txt)

        combined_corpus = "\n\n".join(combined_all_corpus_list)
        if not combined_corpus.strip():
            st.error("لم يتم العثور على نصوص قابلة للاستخراج في الصفحات المحددة. يرجى التحقق من أرقام الصفحات ومحتوى الملفات المرفوعة.")
            st.stop()

        total_extracted_words = sum(len(txt.split()) for txt in combined_all_corpus_list)
        total_extracted_chars = len(combined_corpus)
        status_text.text(f"2/4: تم استخراج {total_extracted_words:,} كلمة ({total_extracted_chars:,} حرفاً) من كامل المرجع بنجاح. جارٍ استدعاء محرك Gemini لبناء الفهرس الفردي...")
        progress_bar.progress(35)

        generated_syllabus_data = None
        generated_qb_list = []

        # ----------------------------------------------------
        # Scenario Routing & Execution
        # ----------------------------------------------------
        if "السيناريو 4" in scenario_choice:
            is_fast_mode = "التوليد السريع" in s4_gen_mode
            all_file_names = list(uploaded_docs_data.keys())
            
            # Extract topics for the 5 terms (Prep 14, Med T1 24, Med T2 28, Fin T1 24, Fin T2 20)
            sec_breakdown_arg = [
                {
                    "name": "القسم الإعدادي",
                    "lectures": 14,
                    "text": all_extracted_sections_text.get("القسم الإعدادي", "") or combined_corpus,
                    "reference": " ،".join(sections_setup[0]["files"].keys()) if sections_setup[0]["files"] else "الدليل الفني المعتمد",
                    "equipment_name": sections_setup[0].get("equipment_name", primary_equipment_name)
                },
                {
                    "name": "القسم المتوسط - ترم أول",
                    "lectures": 24,
                    "text": all_extracted_sections_text.get("القسم المتوسط", "") or combined_corpus,
                    "reference": " ،".join(sections_setup[1]["files"].keys()) if sections_setup[1]["files"] else "الدليل الفني المعتمد",
                    "equipment_name": sections_setup[1].get("equipment_name", primary_equipment_name)
                },
                {
                    "name": "القسم المتوسط - ترم ثاني",
                    "lectures": 28,
                    "text": all_extracted_sections_text.get("القسم المتوسط", "") or combined_corpus,
                    "reference": " ،".join(sections_setup[1]["files"].keys()) if sections_setup[1]["files"] else "الدليل الفني المعتمد",
                    "equipment_name": sections_setup[1].get("equipment_name", primary_equipment_name)
                },
                {
                    "name": "القسم النهائي - ترم أول",
                    "lectures": 24,
                    "text": all_extracted_sections_text.get("القسم النهائي", "") or combined_corpus,
                    "reference": " ،".join(sections_setup[2]["files"].keys()) if sections_setup[2]["files"] else "الدليل الفني المعتمد",
                    "equipment_name": sections_setup[2].get("equipment_name", primary_equipment_name)
                },
                {
                    "name": "القسم النهائي - ترم ثاني",
                    "lectures": 20,
                    "text": all_extracted_sections_text.get("القسم النهائي", "") or combined_corpus,
                    "reference": " ،".join(sections_setup[2]["files"].keys()) if sections_setup[2]["files"] else "الدليل الفني المعتمد",
                    "equipment_name": sections_setup[2].get("equipment_name", primary_equipment_name)
                }
            ]
            
            total_syl_lec = 14 + 24 + 28 + 24 + 20 # 110
            status_text.text(f"3/4: جارٍ استخراج وتنسيق موضوعات التدريب للأقسام الثلاثة ({total_syl_lec} موضوعاً معتمداً)...")
            generated_syllabus_data = generate_syllabus(
                client=client,
                source_text=combined_corpus,
                equipment_name=primary_equipment_name,
                total_lectures=total_syl_lec,
                reference_name=" ،".join(all_file_names) if all_file_names else "الدليل الفني المعتمد",
                custom_instructions=custom_prompt_rules,
                sections_breakdown=sec_breakdown_arg
            )
            progress_bar.progress(50)
            
            all_syl_items = generated_syllabus_data.get("syllabus_items", [])
            prep_items = all_syl_items[0:14]
            med_t1_items = all_syl_items[14:38]
            med_t2_items = all_syl_items[38:66]
            fin_t1_items = all_syl_items[66:90]
            fin_t2_items = all_syl_items[90:110]
            
            bank_segments = [
                {"name": "بنك القسم الإعدادي", "items": prep_items, "sec_key": "القسم الإعدادي", "sec_idx": 0},
                {"name": "بنك المتوسط - ترم أول", "items": med_t1_items, "sec_key": "القسم المتوسط", "sec_idx": 1},
                {"name": "بنك المتوسط - ترم ثاني", "items": med_t2_items, "sec_key": "القسم المتوسط", "sec_idx": 1},
                {"name": "بنك النهائي - ترم أول", "items": fin_t1_items, "sec_key": "القسم النهائي", "sec_idx": 2},
                {"name": "بنك النهائي - ترم ثاني", "items": fin_t2_items, "sec_key": "القسم النهائي", "sec_idx": 2},
            ]
            
            if not is_fast_mode:
                for b_idx, b_info in enumerate(bank_segments):
                    b_name = b_info["name"]
                    b_items = b_info["items"]
                    b_sec = sections_setup[b_info["sec_idx"]]
                    b_eq = b_sec.get("equipment_name", primary_equipment_name)
                    b_rep = b_sec.get("allow_repetition", False)
                    b_corpus = all_extracted_sections_text.get(b_info["sec_key"], "") or combined_corpus
                    
                    status_text.text(f"جارٍ توليد أسئلة [{b_name}] ({len(b_items)} درساً × 25 = {len(b_items)*25} سؤالاً)...")
                    qb_res = generate_question_bank(
                        client=client,
                        source_text=b_corpus,
                        equipment_name=b_eq,
                        section_name=b_name,
                        num_questions=len(b_items) * 25,
                        custom_instructions=custom_prompt_rules,
                        syllabus_items=b_items,
                        allow_repetition=b_rep,
                        status_callback=lambda msg: status_text.text(msg)
                    )
                    qb_res["equipment_name"] = b_eq
                    generated_qb_list.append(qb_res)
                    progress_bar.progress(int(50 + (b_idx + 1) * (40 / len(bank_segments))))
            else:
                progress_bar.progress(90)
                status_text.text("⚡ التوليد السريع: تم استخراج وهيكلة الموضوعات بنجاح، جارٍ تجهيز المصنفات المعتمدة...")

            # Build Training Programs configuration
            s4_sections_setup = [
                {
                    "name": "القسم الإعدادي",
                    "equipment_name": sections_setup[0].get("equipment_name", primary_equipment_name),
                    "lectures": 14,
                    "terms": [
                        {
                            "term_name": "الترم الأول",
                            "items": prep_items,
                            "periods_per_day": 2,
                            "days_count": 14,
                            "midterm_day": 7
                        }
                    ]
                },
                {
                    "name": "القسم المتوسط",
                    "equipment_name": sections_setup[1].get("equipment_name", primary_equipment_name),
                    "lectures": 52,
                    "terms": [
                        {
                            "term_name": "الترم الأول",
                            "items": med_t1_items,
                            "periods_per_day": 4,
                            "days_count": 12,
                            "midterm_day": 6
                        },
                        {
                            "term_name": "الترم الثاني",
                            "items": med_t2_items,
                            "periods_per_day": 4,
                            "days_count": 14,
                            "midterm_day": 7
                        }
                    ]
                },
                {
                    "name": "القسم النهائي",
                    "equipment_name": sections_setup[2].get("equipment_name", primary_equipment_name),
                    "lectures": 44,
                    "terms": [
                        {
                            "term_name": "الترم الأول",
                            "items": fin_t1_items,
                            "periods_per_day": 4,
                            "days_count": 12,
                            "midterm_day": 6
                        },
                        {
                            "term_name": "الترم الثاني",
                            "items": fin_t2_items,
                            "periods_per_day": 4,
                            "days_count": 10,
                            "midterm_day": 5
                        }
                    ]
                }
            ]
            sections_setup = s4_sections_setup
            
            # Specialty sections data for 'برنامج تدريب تخصص'
            specialty_sections_data = [
                {
                    "name": "القسم الإعدادي",
                    "equipment_name": sections_setup[0].get("equipment_name", primary_equipment_name),
                    "banner_title": f"موضوعات برنامج تدريب ( القسم الإعدادي ) - تخصص {sections_setup[0].get('equipment_name', primary_equipment_name)}",
                    "class_label": "الإعدادي",
                    "total_label": "إجمالي القسم الإعدادي",
                    "topics": [
                        {"term": "الترم الأول", "topic": it.get("lesson_name", ""), "th": 2, "pr": 2}
                        for it in prep_items
                    ]
                },
                {
                    "name": "القسم المتوسط",
                    "equipment_name": sections_setup[1].get("equipment_name", primary_equipment_name),
                    "banner_title": f"موضوعات برنامج تدريب ( القسم المتوسط ) - تخصص {sections_setup[1].get('equipment_name', primary_equipment_name)}",
                    "class_label": "المتوسط",
                    "total_label": "إجمالي القسم المتوسط",
                    "topics": [
                        {"term": "الترم الأول", "topic": it.get("lesson_name", ""), "th": 2, "pr": 2}
                        for it in med_t1_items
                    ] + [
                        {"term": "الترم الثاني", "topic": it.get("lesson_name", ""), "th": 2, "pr": 2}
                        for it in med_t2_items
                    ]
                },
                {
                    "name": "القسم النهائي",
                    "equipment_name": sections_setup[2].get("equipment_name", primary_equipment_name),
                    "banner_title": f"موضوعات برنامج تدريب ( القسم النهائي ) - تخصص {sections_setup[2].get('equipment_name', primary_equipment_name)}",
                    "class_label": "النهائي",
                    "total_label": "إجمالي القسم النهائي",
                    "topics": [
                        {"term": "الترم الأول", "topic": it.get("lesson_name", ""), "th": 2, "pr": 2}
                        for it in fin_t1_items
                    ] + [
                        {"term": "الترم الثاني", "topic": it.get("lesson_name", ""), "th": 2, "pr": 2}
                        for it in fin_t2_items
                    ]
                }
            ]

        elif "السيناريو 1" in scenario_choice:
            total_syl_lec = sum(s["lectures"] for s in sections_setup)
            all_file_names = list(uploaded_docs_data.keys())
            sec_breakdown_arg = [
                {
                    "name": sec["name"],
                    "lectures": sec["lectures"],
                    "text": all_extracted_sections_text.get(sec["name"], "") or combined_corpus,
                    "reference": " ،".join(sec["files"].keys()) if sec["files"] else (" ،".join(all_file_names) if all_file_names else "الدليل الفني المعتمد"),
                    "equipment_name": sec.get("equipment_name", primary_equipment_name)
                }
                for sec in sections_setup
            ]
            
            # Generate Single-Row Syllabus
            status_text.text(f"3/4: جارٍ بناء الجدول الزمني وفهرس الدروس للأقسام الثلاثة ({total_syl_lec} درساً فردياً - 1:2:25)...")
            generated_syllabus_data = generate_syllabus(
                client=client,
                source_text=combined_corpus,
                equipment_name=primary_equipment_name,
                total_lectures=total_syl_lec,
                reference_name=" ،".join(all_file_names) if all_file_names else "الدليل الفني المعتمد",
                custom_instructions=custom_prompt_rules,
                sections_breakdown=sec_breakdown_arg
            )
            progress_bar.progress(50)

            all_syl_items = generated_syllabus_data.get("syllabus_items", [])
            cur_lec_offset = 0
            step_fraction = 40 / len(sections_setup)

            for idx, sec in enumerate(sections_setup):
                sec_name = sec["name"]
                sec_l_count = sec["lectures"]
                sec_q_count = sec["questions_count"]
                sec_allow_rep = sec["allow_repetition"]
                sec_eq = sec.get("equipment_name", primary_equipment_name)
                sec_items = all_syl_items[cur_lec_offset : cur_lec_offset + sec_l_count]
                cur_lec_offset += sec_l_count
                
                rep_status_msg = "مع تفعيل خيار التكرار" if sec_allow_rep else "بدون تكرار"
                status_text.text(f"جارٍ توليد بنك أسئلة [{sec_name} - {sec_eq}] ({len(sec_items)} درساً × 25 سؤالاً = {sec_q_count} سؤالاً - {rep_status_msg})...")
                
                sec_corpus = all_extracted_sections_text.get(sec_name, "")
                if not sec_corpus.strip():
                    sec_corpus = combined_corpus
                    
                qb_res = generate_question_bank(
                    client=client,
                    source_text=sec_corpus,
                    equipment_name=sec_eq,
                    section_name=sec_name,
                    num_questions=sec_q_count,
                    custom_instructions=custom_prompt_rules,
                    syllabus_items=sec_items,
                    allow_repetition=sec_allow_rep,
                    status_callback=lambda msg: status_text.text(msg)
                )
                qb_res["equipment_name"] = sec_eq
                generated_qb_list.append(qb_res)
                progress_bar.progress(int(50 + (idx + 1) * step_fraction))

        elif "السيناريو 2" in scenario_choice:
            sec = sections_setup[0]
            sec_name = sec["name"]
            sec_l = sec["lectures"]
            sec_q = sec["questions_count"]
            sec_rep = sec["allow_repetition"]
            sec_eq = sec.get("equipment_name", "المنظومة التخصصية")
            
            rep_msg = "مع تفعيل خيار التكرار" if sec_rep else "بدون تكرار"
            status_text.text(f"3/4: جارٍ بناء الجدول الزمني وبنك أسئلة [{sec_name} - {sec_eq}] ({sec_l} درساً فردياً - {rep_msg})...")
            
            sec_refs_list = sec["files"].keys()
            ref_label = " ،".join(sec_refs_list) if sec_refs_list else "الدليل الفني المعتمد"
            
            generated_syllabus_data = generate_syllabus(
                client=client,
                source_text=combined_corpus,
                equipment_name=sec_eq,
                total_lectures=sec_l,
                reference_name=ref_label,
                custom_instructions=custom_prompt_rules,
                status_callback=lambda msg: status_text.text(msg)
            )
            progress_bar.progress(65)

            sec_corpus = all_extracted_sections_text.get(sec_name, "")
            if not sec_corpus.strip():
                sec_corpus = combined_corpus
                
            qb_res = generate_question_bank(
                client=client,
                source_text=sec_corpus,
                equipment_name=sec_eq,
                section_name=sec_name,
                num_questions=sec_q,
                custom_instructions=custom_prompt_rules,
                syllabus_items=generated_syllabus_data.get("syllabus_items", []),
                allow_repetition=sec_rep,
                status_callback=lambda msg: status_text.text(msg)
            )
            qb_res["equipment_name"] = sec_eq
            generated_qb_list.append(qb_res)
            progress_bar.progress(90)

        elif "السيناريو 3" in scenario_choice:
            sec = sections_setup[0]
            sec_l = sec["lectures"]
            sec_eq = sec.get("equipment_name", "المنظومة الفنية المعتمدة")
            status_text.text(f"3/4: جارٍ إعداد الجدول الزمني وفهرس الدروس لـ [{sec_eq}] ({sec_l} درساً فردياً)...")
            
            sec_refs_list = sec["files"].keys()
            ref_label = " ،".join(sec_refs_list) if sec_refs_list else "الدليل الفني المعتمد"
            
            generated_syllabus_data = generate_syllabus(
                client=client,
                source_text=combined_corpus,
                equipment_name=sec_eq,
                total_lectures=sec_l,
                reference_name=ref_label,
                custom_instructions=custom_prompt_rules,
                status_callback=lambda msg: status_text.text(msg)
            )
            progress_bar.progress(90)

        # Step 4: Build Workbook
        is_scenario_4 = "السيناريو 4" in scenario_choice
        if "السيناريو 3" not in scenario_choice and not (is_scenario_4 and locals().get("is_fast_mode", False)):
            total_qs = sum(len(qb.get("questions", [])) for qb in generated_qb_list)
            if total_qs == 0:
                raise RuntimeError("لم يتم توليد أي أسئلة في بنك الأسئلة! يرجى التحقق من مفتاح الـ API ومعدل الاستهلاك اليومي للنماذج.")

        status_text.text("4/4: جارٍ إنشاء مصنف Excel وتطبيق التنسيق اليميني والقواعد المعتمدة...")
        excel_buf = build_workbook(
            syllabus_data=generated_syllabus_data,
            question_banks=generated_qb_list,
            equipment_name=primary_equipment_name,
            training_program_mode=is_scenario_4,
            sections_setup=sections_setup
        )

        specialty_buf = None
        zip_buf = None
        clean_eq_fname = "".join([c if c.isalnum() else "_" for c in primary_equipment_name])

        if is_scenario_4 and locals().get("specialty_sections_data"):
            specialty_buf = build_specialty_program_workbook(primary_equipment_name, specialty_sections_data)
            zip_buf = io.BytesIO()
            with zipfile.ZipFile(zip_buf, "w", zipfile.ZIP_DEFLATED) as zf:
                zf.writestr(f"المصنف_الشامل_المعتمد_{clean_eq_fname}.xlsx", excel_buf.getvalue())
                zf.writestr(f"برنامج_تدريب_تخصص_{clean_eq_fname}.xlsx", specialty_buf.getvalue())
            zip_buf.seek(0)

        progress_bar.progress(100)
        status_text.text("✅ اكتملت المعالجة وتطبيق النموذج المعتمد بنجاح تام!")

        st.session_state.generation_results = {
            "syllabus": generated_syllabus_data,
            "question_banks": generated_qb_list,
            "equipment_name": primary_equipment_name,
            "specialty_data": locals().get("specialty_sections_data"),
            "is_scenario_4": is_scenario_4,
            "timestamp": datetime.now().strftime("%Y-%m-%d %H:%M")
        }
        st.session_state.excel_buffer = excel_buf
        st.session_state.specialty_buffer = specialty_buf
        st.session_state.zip_buffer = zip_buf

    except Exception as exc:
        st.error(f"❌ حدث خطأ أثناء المعالجة: {str(exc)}")
        st.exception(exc)

# -----------------------------------------------------------------------------
# Section 5: Display Results, Previews & Download
# -----------------------------------------------------------------------------
if st.session_state.generation_results and st.session_state.excel_buffer:
    res = st.session_state.generation_results
    st.divider()
    
    st.subheader("📑 استعراض وتحميل منظومة التدريب وبنوك الأسئلة المعتمدة")
    
    file_timestamp = datetime.now().strftime("%Y%m%d_%H%M")
    clean_eq_fname = "".join([c if c.isalnum() else "_" for c in res["equipment_name"]])

    col_dl1, col_dl2, col_dl3 = st.columns(3)
    with col_dl1:
        st.download_button(
            label="📥 تحميل المصنف الشامل الكامل (.xlsx)",
            data=st.session_state.excel_buffer.getvalue(),
            file_name=f"المصنف_الشامل_المعتمد_{clean_eq_fname}_{file_timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True,
            help="يحتوي على برامج التدريب للأقسام الثلاثة + بنوك الأسئلة الخمسة + برنامج تدريب تخصص."
        )
    with col_dl2:
        if st.session_state.get("specialty_buffer"):
            st.download_button(
                label="📥 تحميل ملف برنامج تدريب تخصص (.xlsx)",
                data=st.session_state.specialty_buffer.getvalue(),
                file_name=f"برنامج_تدريب_تخصص_{clean_eq_fname}_{file_timestamp}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="secondary",
                use_container_width=True,
                help="ملف برنامج تدريب تخصص المستقل المطابق تماماً لنموذج الاعتماد الرسمي."
            )
    with col_dl3:
        if st.session_state.get("zip_buffer"):
            st.download_button(
                label="📦 تحميل الحزمة الكاملة (.zip)",
                data=st.session_state.zip_buffer.getvalue(),
                file_name=f"حزمة_التدريب_وبنوك_الأسئلة_{clean_eq_fname}_{file_timestamp}.zip",
                mime="application/zip",
                type="secondary",
                use_container_width=True,
                help="تحميل كلا الملفين داخل أرشيف مضغوط واحد."
            )

    st.info("✨ تم توثيق وتأمين كافة الملفات والمصنفات وفق النماذج المعتمدة رسمياً ومطابقة 100% لمعايير المنظومة.")

    # Tabs for viewing
    tab_titles = []
    if res.get("specialty_data"):
        tab_titles.append("📋 جدول برنامج تدريب تخصص")
    if res["syllabus"]:
        tab_titles.append("📅 فهرس الدروس والموضوعات")
    for qb in res["question_banks"]:
        tab_titles.append(f"📝 {qb['section_name']}")

    tabs = st.tabs(tab_titles)
    tab_idx = 0

    # 1. Preview Specialty Training Program
    if res.get("specialty_data") and tab_idx < len(tabs):
        with tabs[tab_idx]:
            spec_rows = []
            for sec_d in res["specialty_data"]:
                sec_lbl = sec_d.get("name", "")
                c_lbl = sec_d.get("class_label", "")
                t_list = sec_d.get("topics", [])
                for s_idx, t in enumerate(t_list, start=1):
                    spec_rows.append({
                        "م": s_idx,
                        "القسم": sec_lbl,
                        "اسم الموضوع": t.get("topic", ""),
                        "السنة الدراسية": c_lbl,
                        "ساعات نظري": t.get("th", 2),
                        "ساعات عملي": t.get("pr", 2),
                        "الترم": t.get("term", "الترم الأول"),
                        "يصلح للامتحان المنظومة": "يصلح لعدد ساعات النظرى فقط"
                    })
            df_spec = pd.DataFrame(spec_rows)
            st.write(f"**إجمالي موضوعات التخصص المعتمدة:** {len(df_spec)} موضوعاً تدريبياً")
            st.dataframe(df_spec, use_container_width=True, hide_index=True)
        tab_idx += 1

    # 1. Preview Syllabus
    if res["syllabus"] and tab_idx < len(tabs):
        with tabs[tab_idx]:
            s_data = res["syllabus"]
            items = s_data.get("syllabus_items", [])
            df_syl = pd.DataFrame(items)
            
            c_m1, c_m2, c_m3 = st.columns(3)
            c_m1.metric("إجمالي المحاضرات (السطور)", f"{s_data.get('total_lectures')} محاضرة فردية")
            c_m2.metric("إجمالي الساعات المخصصة", f"{s_data.get('total_hours')} ساعة (معدل 2 س/درس)")
            c_m3.metric("إجمالي الأسئلة المستهدفة", f"{s_data.get('total_questions')} سؤالاً (معدل 25 س/درس)")
            
            col_rename_syl = {
                "id": "م",
                "lesson_name": "عنوان الدرس / المكون الفني",
                "teaching_hours": "عدد الساعات المخصصة",
                "lecture_count": "عدد المحاضرات",
                "question_count": "عدد الأسئلة المستهدفة",
                "reference": "المرجع المعتمد"
            }
            df_syl = df_syl.rename(columns=col_rename_syl)
            st.dataframe(df_syl, use_container_width=True, hide_index=True)
            
        tab_idx += 1

    # 2. Preview Question Banks
    for qb in res["question_banks"]:
        if tab_idx < len(tabs):
            with tabs[tab_idx]:
                q_list = qb.get("questions", [])
                st.write(f"**القسم:** {qb.get('section_name')} | **عدد الأسئلة المتولدة:** {len(q_list)} سؤالاً")
                
                df_q = pd.DataFrame(q_list)
                col_rename_q = {
                    "id": "م",
                    "question_text": "نص السؤال",
                    "explanation": "الشرح والتفسير",
                    "option_a": "الخيار أ",
                    "option_b": "الخيار ب",
                    "option_c": "الخيار ج",
                    "option_d": "الخيار د",
                    "correct_answer": "الإجابة الصحيحة",
                    "difficulty": "مستوى الصعوبة",
                    "lesson": "الدرس الفني المعتمد",
                    "question_type": "نوع السؤال"
                }
                df_q = df_q.rename(columns=col_rename_q)
                st.dataframe(df_q, use_container_width=True, hide_index=True)
            tab_idx += 1

# -----------------------------------------------------------------------------
# Footer
# -----------------------------------------------------------------------------
st.markdown("---")
st.markdown(
    """
    <div style="text-align: center; color: #64748b; font-size: 13px; padding: 10px 0;">
        <span>تطوير: <b>Eslam Abdelbadea</b> | جميع حقوق التطوير محفوظة ©</span>
    </div>
    """,
    unsafe_allow_html=True
)

# -----------------------------------------------------------------------------
# Vercel Serverless WSGI Entrypoint Compatibility
# -----------------------------------------------------------------------------
def handler(environ, start_response):
    status = '200 OK'
    response_headers = [('Content-Type', 'text/html; charset=utf-8')]
    start_response(status, response_headers)
    html_file = os.path.join(os.path.dirname(__file__), 'index.html')
    if os.path.exists(html_file):
        with open(html_file, 'rb') as f:
            return [f.read()]
    return [b"<h1>Question Bank Architect</h1><p>Developed by Eslam Abdelbadea</p>"]

app = handler
application = handler

