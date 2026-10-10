# -*- coding: utf-8 -*-
"""
واجهة التدقيق والمراجعة والاعتماد المؤسسي لبنوك الأسئلة وبرامج التدريب
Streamlit Tab & Component for Exam Bank Auditor
تطوير: Eslam Abdelbadea
"""

import os
import io
import tempfile
import pandas as pd
import streamlit as st
import openpyxl
from exam_bank_auditor import ExamBankAuditor, ExamBankFixer

def render_audit_tab():
    st.markdown(
        """
        <div style="background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%); padding: 22px 26px; border-radius: 14px; border-right: 6px solid #10b981; color: white; margin-bottom: 24px; box-shadow: 0 4px 20px rgba(0,0,0,0.15);">
            <div style="display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap;">
                <div>
                    <h2 style="margin: 0; color: #f8fafc; font-size: 22px; font-weight: 800;">🛡️ منظومة التدقيق والمراجعة والاعتماد المؤسسي</h2>
                    <p style="margin: 6px 0 0 0; color: #94a3b8; font-size: 14px;">محاكاة كاملة لفحص لجان الامتحانات: كشف التكرار (Conditional Formatting)، بيفوت تيبل، العبارات المحظورة، ومطابقة ترتيب الدروس</p>
                </div>
                <div style="background: rgba(16, 185, 129, 0.15); border: 1px solid #10b981; padding: 6px 14px; border-radius: 8px;">
                    <span style="color: #6ee7b7; font-weight: 700; font-size: 13px;">8 اختبارات معيارية صارمة</span>
                </div>
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    col_upload_bank, col_upload_prog = st.columns(2)

    with col_upload_bank:
        st.subheader("📑 1. ملف بنك الأسئلة")
        uploaded_bank = st.file_uploader(
            "اختر ملف بنك الأسئلة (Excel .xlsx):",
            type=["xlsx"],
            key="audit_bank_file",
            help="ملف بنك الأسئلة المتضمن شيت Questions وشيت Lookups"
        )
        if uploaded_bank:
            st.success(f"✅ تم تحميل: **{uploaded_bank.name}** ({uploaded_bank.size / 1024:.1f} KB)")

    with col_upload_prog:
        st.subheader("📋 2. ملف برنامج التدريب (اختياري للمطابقة)")
        uploaded_prog = st.file_uploader(
            "اختر ملف برنامج التدريب المعتمد (Excel .xlsx):",
            type=["xlsx"],
            key="audit_prog_file",
            help="اختياري: لمطابقة التسلسل والترتيب الفعلي للدروس مع شيت برنامج التدريب"
        )
        prog_sheet_name = None
        if uploaded_prog:
            st.success(f"✅ تم تحميل: **{uploaded_prog.name}** ({uploaded_prog.size / 1024:.1f} KB)")
            try:
                wb_temp = openpyxl.load_workbook(io.BytesIO(uploaded_prog.getvalue()), read_only=True)
                available_sheets = wb_temp.sheetnames
                prog_sheet_name = st.selectbox(
                    "حدد شيت البرنامج التدريبي المراد المطابقة معه:",
                    options=available_sheets,
                    index=0,
                    help="اختر الشيت الذي يحتوي على جدول الدروس والمحاضرات (مثل: برنامج تدريب - القسم المتوسط)"
                )
            except Exception as e:
                st.warning(f"تعذر قراءة أسماء الشيتات: {e}")

    col_qpl, col_opt = st.columns([1, 2])
    with col_qpl:
        expected_q = st.number_input("عدد الأسئلة المستهدف لكل درس:", min_value=5, max_value=100, value=25, step=1)
    with col_opt:
        st.markdown("<div style='height: 28px;'></div>", unsafe_allow_html=True)
        auto_fix_enabled = st.checkbox("تفعيل خاصية التصحيح التلقائي للأخطاء الفنية (تفريغ خيارات صواب وخطأ، تفريغ الشرح، ضبط الإجابات) وتوليد نسخة مصححة معتمدة", value=True)

    st.markdown("<div style='height: 10px;'></div>", unsafe_allow_html=True)

    if st.button("🚀 بدء التدقيق والفحص الشامل وإصدار تقرير الاعتماد", type="primary", use_container_width=True):
        if not uploaded_bank:
            st.error("⚠️ يرجى رفع ملف بنك الأسئلة أولاً لبدء الفحص.")
            return

        with st.spinner("⏳ جاري تنفيذ الاختبارات الثمانية المعيارية وفحص مصفوفة الأسئلة..."):
            with tempfile.TemporaryDirectory() as tmpdir:
                bank_tmp_path = os.path.join(tmpdir, uploaded_bank.name)
                with open(bank_tmp_path, "wb") as f:
                    f.write(uploaded_bank.getvalue())

                prog_tmp_path = None
                if uploaded_prog:
                    prog_tmp_path = os.path.join(tmpdir, uploaded_prog.name)
                    with open(prog_tmp_path, "wb") as f:
                        f.write(uploaded_prog.getvalue())

                try:
                    auditor = ExamBankAuditor(
                        bank_path=bank_tmp_path,
                        program_path=prog_tmp_path,
                        program_sheet=prog_sheet_name,
                        questions_per_lesson=expected_q
                    )
                    results = auditor.run_all_audits()

                    # عرض النتيجة الرئيسية
                    score = results["score"]
                    status = results["overall_status"]

                    score_color = "#10b981" if score >= 90 else ("#f59e0b" if score >= 75 else "#ef4444")
                    bg_color = "rgba(16, 185, 129, 0.1)" if score >= 90 else ("rgba(245, 158, 11, 0.1)" if score >= 90 else "rgba(239, 68, 68, 0.1)")

                    st.markdown(
                        f"""
                        <div style="background: {bg_color}; border: 2px solid {score_color}; border-radius: 12px; padding: 20px; text-align: center; margin: 20px 0;">
                            <span style="font-size: 16px; font-weight: 700; color: #64748b;">النتيجة المؤسسية النهائية للاعتماد:</span><br>
                            <span style="font-size: 32px; font-weight: 900; color: {score_color};">{status}</span><br>
                            <span style="font-size: 18px; font-weight: 800; color: #1e293b;">مؤشر جودة البنك ومطابقته: <b>{score:.1f}%</b></span>
                        </div>
                        """,
                        unsafe_allow_html=True
                    )

                    # جدول ملخص الاختبارات
                    st.write("### 📊 لوحة نتائج الاختبارات الثمانية:")
                    
                    tests_summary = [
                        {
                            "الاختبار": "1. التكرار والتمييز الشرطي (Conditional Formatting)",
                            "الحالة": "✅ سليم ومطابق" if results["duplicates"]["status"] == "PASS" else "❌ مرفوض (يوجد تكرار)",
                            "التفاصيل": f"تم رصد {results['duplicates']['count']} تكرار لفظي"
                        },
                        {
                            "الاختبار": "2. بيفوت تيبل لموازنة الأسئلة لكل درس",
                            "الحالة": "✅ سليم ومطابق" if results["lesson_counts"]["status"] == "PASS" else "❌ خلل في النصاب",
                            "التفاصيل": f"تم فحص {results['lesson_counts'].get('total_lessons', 0)} درساً بمعدل {expected_q} س/درس"
                        },
                        {
                            "الاختبار": "3. العبارات المحظورة (جميع ما سبق / كلاهما صواب)",
                            "الحالة": "✅ سليم ومطابق" if results["forbidden_phrases"]["status"] == "PASS" else "❌ خيارات محظورة",
                            "التفاصيل": f"تم رصد {results['forbidden_phrases']['count']} خيار مخالف أكاديمياً"
                        },
                        {
                            "الاختبار": "4. نسب وتدرج مستويات الصعوبة وأنواع الأسئلة",
                            "الحالة": "✅ متوازن" if results["distribution"]["status"] == "PASS" else "⚠️ تنبيه توازن",
                            "التفاصيل": f"نسب الصعوبة وتوزيع 50% متعدد و50% صواب وخطأ"
                        },
                        {
                            "الاختبار": "5. ضوابط أسئلة صواب وخطأ (تفريغ الخيارات F..I)",
                            "الحالة": "✅ سليم ومطابق" if results["tf_rules"]["status"] == "PASS" else "❌ أعمدة خيارات ممتلئة",
                            "التفاصيل": "الأعمدة F-I فارغة تماماً والإجابة فقط في العمود E"
                        },
                        {
                            "الاختبار": "6. اكتمال خيارات اختيار من متعدد (أ، ب، ج، د)",
                            "الحالة": "✅ سليم ومطابق" if results["mcq_rules"]["status"] == "PASS" else "❌ خيارات ناقصة",
                            "التفاصيل": "4 خيارات مكتملة لكل سؤال وإجابة محددة كـ (A, B, C, D)"
                        },
                        {
                            "الاختبار": "7. مطابقة تسلسل وترتيب الدروس مع البرنامج التدريبي",
                            "الحالة": ("✅ تطابق تام 100%" if results["syllabus_match"]["status"] == "PASS" else "❌ عدم تطابق في الترتيب") if uploaded_prog else "⚪ تم التخطي (لم يُرفع برنامج)",
                            "التفاصيل": f"مطابقة تسلسل الدروس 1-to-1 مع شيت التدريب" if uploaded_prog else "لم يتم تحديد ملف برنامج تدريبي"
                        },
                        {
                            "الاختبار": "8. سرعة الفتح، الشرح، وصلاحية الـ Data Validation",
                            "الحالة": "✅ فائق السرعة" if results["speed_validation"]["status"] == "PASS" else "⚠️ تنبيه سرعة",
                            "التفاصيل": f"تفريغ عمود الشرح وصحة نطاقات القوائم المنسدلة"
                        }
                    ]
                    st.dataframe(pd.DataFrame(tests_summary), use_container_width=True, hide_index=True)

                    # عرض التفاصيل إن وجدت مخالفات
                    if results["duplicates"]["count"] > 0:
                        with st.expander(f"❌ تفاصيل الأسئلة المكررة ({results['duplicates']['count']} سؤال)", expanded=True):
                            df_dups = pd.DataFrame(results["duplicates"]["details"])
                            st.dataframe(df_dups, use_container_width=True)

                    if results["forbidden_phrases"]["count"] > 0:
                        with st.expander(f"⚠️ تفاصيل العبارات المحظورة في الخيارات ({results['forbidden_phrases']['count']} خيار)", expanded=True):
                            df_forb = pd.DataFrame(results["forbidden_phrases"]["details"])
                            st.dataframe(df_forb, use_container_width=True)

                    if results["lesson_counts"]["discrepancies"]:
                        with st.expander("⚠️ تفاصيل تفاوت عدد الأسئلة في الدروس عن النصاب المطلوب", expanded=True):
                            df_disc = pd.DataFrame(results["lesson_counts"]["discrepancies"])
                            st.dataframe(df_disc, use_container_width=True)

                    if uploaded_prog and results["syllabus_match"]["mismatches"]:
                        with st.expander("❌ تفاصيل اختلاف ترتيب الدروس بين البنك والبرنامج", expanded=True):
                            df_miss = pd.DataFrame(results["syllabus_match"]["mismatches"])
                            st.dataframe(df_miss, use_container_width=True)

                    # التوصيات
                    if results["recommendations"]:
                        st.write("### 💡 التوصيات والإجراءات التصحيحية المقترحة:")
                        for rec in results["recommendations"]:
                            st.info(rec)

                    # تصدير تقرير التدقيق
                    md_report_path = os.path.join(tmpdir, "Audit_Report.md")
                    auditor.export_markdown_report(md_report_path)
                    with open(md_report_path, "r", encoding="utf-8") as f:
                        md_content = f.read()

                    col_rep_btn, col_fix_btn = st.columns(2)
                    with col_rep_btn:
                        st.download_button(
                            label="📥 تحميل التقرير الرسمي للتدقيق والاعتماد (Markdown)",
                            data=md_content.encode("utf-8"),
                            file_name=f"تقرير_اعتماد_{uploaded_bank.name.replace('.xlsx', '')}.md",
                            mime="text/markdown",
                            use_container_width=True
                        )

                    # معالجة التصحيح التلقائي إذا كان مفعلاً
                    if auto_fix_enabled:
                        fixer = ExamBankFixer(bank_path=bank_tmp_path)
                        fixed_path = fixer.fix_all()
                        with open(fixed_path, "rb") as f_fix:
                            fixed_bytes = f_fix.read()

                        with col_fix_btn:
                            st.download_button(
                                label="✨ تحميل بنك الأسئلة بعد التصحيح التلقائي الفوري (Excel)",
                                data=fixed_bytes,
                                file_name=f"{uploaded_bank.name.replace('.xlsx', '')}_مصحح_ومعتمد.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary",
                                use_container_width=True
                            )

                        if fixer.fixed_log:
                            with st.expander("🛠️ سجل عمليات التصحيح التلقائي التي تمت بنجاح"):
                                for log_item in fixer.fixed_log:
                                    st.write(f"• {log_item}")

                except Exception as e:
                    st.error(f"حدث خطأ أثناء إجراء التدقيق: {str(e)}")
                    st.exception(e)
