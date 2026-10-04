"""
Updates question banks strictly matching user-edited training program sheets:
1. Preserves user-edited training program sheets (hours distribution, topics) untouched!
2. Reads exact theoretical topics directly from the program sheets.
3. In question banks:
   - Header strictly:
     [نوع السؤال, نص السؤال, الشرح / التفسير, مستوى الصعوبة, الإجابة الصحيحة, الخيار أ (A), الخيار ب (B), الخيار ج (C), الخيار د (D), الدرس]
   - Column 'الإجابة الصحيحة': Written as 'الخيار أ (A)', 'الخيار ب (B)', 'الخيار ج (C)', 'الخيار د (D)' (not answer text).
   - Data validation applied to:
     - Column A (نوع السؤال): اختيار من متعدد,صواب أو خطأ
     - Column D (مستوى الصعوبة): سهل,متوسط,صعب
     - Column E (الإجابة الصحيحة): الخيار أ (A),الخيار ب (B),الخيار ج (C),الخيار د (D)
   - Column J (الدرس): Exactly matches the lesson title in the program sheet.
"""

import os
import sys
from copy import copy
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

sys.stdout.reconfigure(encoding='utf-8')

def create_thin_border():
    thin = Side(border_style='thin', color='A0A0A0')
    return Border(left=thin, right=thin, top=thin, bottom=thin)

def apply_rtl(ws):
    try:
        ws.sheet_view.rightToLeft = True
        ws.sheet_view.showGridLines = True
    except Exception:
        pass

def get_correct_option_label(correct_text, opt_a, opt_b, opt_c, opt_d):
    correct_str = str(correct_text or '').strip()
    if correct_str in ['أ', 'A', 'الخيار أ', 'الخيار أ (A)']: return 'الخيار أ (A)'
    if correct_str in ['ب', 'B', 'الخيار ب', 'الخيار ب (B)']: return 'الخيار ب (B)'
    if correct_str in ['ج', 'C', 'الخيار ج', 'الخيار ج (C)']: return 'الخيار ج (C)'
    if correct_str in ['د', 'D', 'الخيار د', 'الخيار د (D)']: return 'الخيار د (D)'

    def clean(s):
        s = str(s or '').strip()
        for ch in [' ', '،', '؟', '!', '.', ':', '-', '_', 'أ', 'إ', 'آ', 'ة', 'ه', 'ى', 'ي']:
            s = s.replace(ch, '')
        return s.lower()

    c_clean = clean(correct_str)
    opts = [
        ('الخيار أ (A)', opt_a, clean(opt_a)),
        ('الخيار ب (B)', opt_b, clean(opt_b)),
        ('الخيار ج (C)', opt_c, clean(opt_c)),
        ('الخيار د (D)', opt_d, clean(opt_d)),
    ]
    # Exact match
    for label, raw, cl in opts:
        if correct_str == str(raw or '').strip():
            return label
    # Normalized clean match
    for label, raw, cl in opts:
        if c_clean and c_clean == cl:
            return label
    # Substring match
    for label, raw, cl in opts:
        if c_clean and (c_clean in cl or cl in c_clean):
            return label
    return 'الخيار أ (A)'

def main():
    target_file = r"C:\Users\MaximuM-Tech\Downloads\بنوك\برنامج تدريب تخصص الضبع الاسود_النسخة_الرسمية_المعتمدة.xlsx"
    print(f"Loading user-edited workbook from: {target_file}")
    wb = openpyxl.load_workbook(target_file)

    # 1. Extract theoretical lesson titles directly from the program sheets
    # A) Prep
    ws_prep = wb['برنامج تدريب - القسم الإعدادي']
    prep_topics = []
    for r in range(5, ws_prep.max_row + 1):
        c = ws_prep.cell(r, 3).value
        d = ws_prep.cell(r, 4).value
        e = ws_prep.cell(r, 5).value
        if c == 'ف1' and d and (e == 2 or e == '2'):
            prep_topics.append(d)
    print(f"Extracted {len(prep_topics)} theoretical topics from برنامج تدريب - القسم الإعدادي")

    # B) Med
    ws_med = wb['برنامج تدريب - القسم المتوسط']
    med_t1_topics = []
    med_t2_topics = []
    for r in range(5, ws_med.max_row + 1):
        c = ws_med.cell(r, 3).value
        d = ws_med.cell(r, 4).value
        e = ws_med.cell(r, 5).value
        if (c in ['ف1', 'ف3']) and d and (e == 2 or e == '2'):
            if r <= 54:
                med_t1_topics.append(d)
            else:
                med_t2_topics.append(d)
    print(f"Extracted {len(med_t1_topics)} T1 topics and {len(med_t2_topics)} T2 topics from برنامج تدريب - القسم المتوسط")

    # C) Final
    ws_fin = wb['برنامج تدريب - القسم النهائي']
    fin_t1_topics = []
    fin_t2_topics = []
    for r in range(5, ws_fin.max_row + 1):
        c = ws_fin.cell(r, 3).value
        d = ws_fin.cell(r, 4).value
        e = ws_fin.cell(r, 5).value
        if (c in ['ف1', 'ف3']) and d and (e == 2 or e == '2'):
            if r <= 54:
                fin_t1_topics.append(d)
            else:
                fin_t2_topics.append(d)
    print(f"Extracted {len(fin_t1_topics)} T1 topics and {len(fin_t2_topics)} T2 topics from برنامج تدريب - القسم النهائي")

    # 2. Load Question Pools
    print("\nLoading question pools...")
    dhab_pool_file = os.path.join(r"C:\Users\MaximuM-Tech\Downloads\بنوك", "معدة الضبع الاسود.xlsx")
    wb_dhab = openpyxl.load_workbook(dhab_pool_file, data_only=True)
    ws_dhab_src = wb_dhab.active

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
                "correct_raw": ws_dhab_src.cell(r, 8).value or "",
                "difficulty": ws_dhab_src.cell(r, 9).value or "متوسط",
                "topic": ws_dhab_src.cell(r, 10).value or "",
                "q_type": ws_dhab_src.cell(r, 11).value or "اختيار من متعدد"
            })

    ejla_file = os.path.join(r"C:\Users\MaximuM-Tech\Downloads\بنوك", "نهائي ضبع اسود.xlsx")
    raw_ejla_qs = []
    if os.path.exists(ejla_file):
        wb_ejla = openpyxl.load_workbook(ejla_file, data_only=True)
        ws_ejla = wb_ejla.active
        for r in range(2, ws_ejla.max_row + 1):
            q_text = ws_ejla.cell(r, 2).value
            if q_text is not None:
                raw_ejla_qs.append({
                    "q_text": q_text,
                    "explanation": ws_ejla.cell(r, 3).value or "وفقاً للمرجع الموحد لمنظومة إيجلا",
                    "opt_a": ws_ejla.cell(r, 4).value or "",
                    "opt_b": ws_ejla.cell(r, 5).value or "",
                    "opt_c": ws_ejla.cell(r, 6).value or "",
                    "opt_d": ws_ejla.cell(r, 7).value or "",
                    "correct_raw": ws_ejla.cell(r, 8).value or "",
                    "difficulty": ws_ejla.cell(r, 9).value or "متوسط",
                    "topic": ws_ejla.cell(r, 10).value or "",
                    "q_type": ws_ejla.cell(r, 11).value or "اختيار من متعدد"
                })

    # 3. Helper to build bank rows
    def create_bank_rows(topics_list, pool):
        bank_rows = []
        pool_cursor = 0
        diff_cycle = ["سهل", "متوسط", "صعب", "متوسط"]

        for topic in topics_list:
            for q_i in range(25):
                src = pool[pool_cursor % len(pool)]
                pool_cursor += 1

                diff = src.get('difficulty')
                if not diff or diff not in ["سهل", "متوسط", "صعب"]:
                    diff = diff_cycle[q_i % len(diff_cycle)]

                q_type = src.get('q_type', 'اختيار من متعدد')
                if not q_type or q_type not in ["اختيار من متعدد", "صواب أو خطأ"]:
                    q_type = "اختيار من متعدد"

                opt_a = src['opt_a']
                opt_b = src['opt_b']
                opt_c = src['opt_c']
                opt_d = src['opt_d']
                correct_label = get_correct_option_label(src['correct_raw'], opt_a, opt_b, opt_c, opt_d)

                row = [
                    q_type,
                    src['q_text'],
                    src['explanation'],
                    diff,
                    correct_label,
                    opt_a,
                    opt_b,
                    opt_c,
                    opt_d,
                    topic  # Exactly matching the topic in program sheet!
                ]
                bank_rows.append(row)
        return bank_rows

    border_cell = create_thin_border()
    qb_header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    qb_data_font = Font(name="Calibri", size=10)
    qb_header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")

    def populate_qb_sheet(ws, q_rows):
        apply_rtl(ws)
        # Clear existing rows if any
        ws.delete_rows(1, ws.max_row + 1)

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
                if col_idx in [1, 4, 5]:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.alignment = Alignment(horizontal="right", vertical="center")

        widths = {1: 18, 2: 50, 3: 35, 4: 15, 5: 18, 6: 22, 7: 22, 8: 22, 9: 22, 10: 45}
        for col_idx, w in widths.items():
            ws.column_dimensions[get_column_letter(col_idx)].width = w

        max_r = max(len(q_rows) + 50, 200)

        # Data Validation 1: نوع السؤال (Col A)
        dv_type = DataValidation(type="list", formula1='"اختيار من متعدد,صواب أو خطأ"', allow_blank=True)
        ws.add_data_validation(dv_type)
        dv_type.add(f"A2:A{max_r}")

        # Data Validation 2: مستوى الصعوبة (Col D)
        dv_diff = DataValidation(type="list", formula1='"سهل,متوسط,صعب"', allow_blank=True)
        ws.add_data_validation(dv_diff)
        dv_diff.add(f"D2:D{max_r}")

        # Data Validation 3: الإجابة الصحيحة (Col E)
        dv_ans = DataValidation(type="list", formula1='"الخيار أ (A),الخيار ب (B),الخيار ج (C),الخيار د (D)"', allow_blank=True)
        ws.add_data_validation(dv_ans)
        dv_ans.add(f"E2:E{max_r}")

    # Build Question Bank sheets
    print("\nPopulating Question Bank sheets...")
    # 1. بنك القسم الإعدادي
    prep_rows = create_bank_rows(prep_topics, raw_dhab_qs)
    ws_qb_prep = wb['بنك القسم الإعدادي'] if 'بنك القسم الإعدادي' in wb.sheetnames else wb.create_sheet('بنك القسم الإعدادي')
    populate_qb_sheet(ws_qb_prep, prep_rows)
    print(f"  - بنك القسم الإعدادي: {len(prep_rows)} سؤال")

    # 2. بنك المتوسط - ترم أول
    med_t1_rows = create_bank_rows(med_t1_topics, raw_dhab_qs)
    ws_qb_med1 = wb['بنك المتوسط - ترم أول'] if 'بنك المتوسط - ترم أول' in wb.sheetnames else wb.create_sheet('بنك المتوسط - ترم أول')
    populate_qb_sheet(ws_qb_med1, med_t1_rows)
    print(f"  - بنك المتوسط - ترم أول: {len(med_t1_rows)} سؤال")

    # 3. بنك المتوسط - ترم ثاني
    med_t2_rows = create_bank_rows(med_t2_topics, raw_dhab_qs)
    ws_qb_med2 = wb['بنك المتوسط - ترم ثاني'] if 'بنك المتوسط - ترم ثاني' in wb.sheetnames else wb.create_sheet('بنك المتوسط - ترم ثاني')
    populate_qb_sheet(ws_qb_med2, med_t2_rows)
    print(f"  - بنك المتوسط - ترم ثاني: {len(med_t2_rows)} سؤال")

    # 4. بنك النهائي - ترم أول
    fin_pool = raw_ejla_qs if raw_ejla_qs else raw_dhab_qs
    fin_t1_rows = create_bank_rows(fin_t1_topics, fin_pool)
    ws_qb_fin1 = wb['بنك النهائي - ترم أول'] if 'بنك النهائي - ترم أول' in wb.sheetnames else wb.create_sheet('بنك النهائي - ترم أول')
    populate_qb_sheet(ws_qb_fin1, fin_t1_rows)
    print(f"  - بنك النهائي - ترم أول: {len(fin_t1_rows)} سؤال")

    # 5. بنك النهائي - ترم ثاني
    fin_t2_rows = create_bank_rows(fin_t2_topics, fin_pool)
    ws_qb_fin2 = wb['بنك النهائي - ترم ثاني'] if 'بنك النهائي - ترم ثاني' in wb.sheetnames else wb.create_sheet('بنك النهائي - ترم ثاني')
    populate_qb_sheet(ws_qb_fin2, fin_t2_rows)
    print(f"  - بنك النهائي - ترم ثاني: {len(fin_t2_rows)} سؤال")

    # Save to files
    wb.save(target_file)
    print(f"\nSUCCESS: Saved updated workbook to {target_file}")

    primary_file = r"C:\Users\MaximuM-Tech\Downloads\بنوك\برنامج تدريب تخصص الضبع الاسود.xlsx"
    try:
        wb.save(primary_file)
        print(f"SUCCESS: Also updated {primary_file}")
    except PermissionError:
        print(f"NOTE: {primary_file} is open in Excel.")

if __name__ == '__main__':
    main()
