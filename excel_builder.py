"""
Excel Builder Module for Question Banks, Training Programs & Specialty Curricula
Full RTL layout, Committee-approved templates, Data Validations, and (1 : 2 : 25) formula.
حقوق الملكية الفكرية وتطوير النظام: إسلام عبد البديع
"""

import io
from typing import List, Dict, Any, Optional
from collections import OrderedDict
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

# Difficulty Color Mapping
DIFFICULTY_COLORS = {
    "سهل": "E2EFDA",
    "متوسط": "FFF2CC",
    "صعب": "FCE4D6",
    "صعب جداً": "F8CBAD",
    "تفوق": "A9D08E"
}

# Neutral workbook headers
BRAND_NOTICE = "منظومة بنوك الأسئلة والتدريب التخصصي القياسية"
FOOTER_NOTICE = "تم إعداد وتنسيق هذا النموذج آلياً وفق المعايير القياسية للتقييم والتأهيل"

# Palette definitions
NAVY_HEADER = "1F4E78"
ACCENT_BLUE = "D9E1F2"
GRAY_ZEBRA = "F9FAFB"
BORDER_GRAY = "D3D3D3"
TOTAL_ROW_FILL = "EAEEF3"

ARABIC_DAYS = [
    "الأول", "الثاني", "الثالث", "الرابع", "الخامس", "السادس",
    "السابع", "الثامن", "التاسع", "العاشر", "الحادي عشر", "الثانى عشر",
    "الثالث عشر", "الرابع عشر", "الخامس عشر", "السادس عشر"
]


def create_thin_border(color: str = BORDER_GRAY) -> Border:
    thin = Side(border_style="thin", color=color)
    return Border(left=thin, right=thin, top=thin, bottom=thin)


def apply_rtl_and_grid(ws):
    """Configures Right-to-Left orientation and enables visible gridlines."""
    try:
        if hasattr(ws, 'sheet_view'):
            ws.sheet_view.rightToLeft = True
            ws.sheet_view.showGridLines = True
    except Exception:
        pass
    try:
        if hasattr(ws, 'views') and len(ws.views.sheetView) > 0:
            ws.views.sheetView[0].rightToLeft = True
            ws.views.sheetView[0].showGridLines = True
    except Exception:
        pass


def auto_fit_columns(ws, max_cols: int, min_width: int = 12, max_width: int = 65):
    """Adjusts column widths dynamically for Arabic text."""
    for col_idx in range(1, max_cols + 1):
        col_letter = get_column_letter(col_idx)
        max_len = 0
        for row in range(3, ws.max_row + 1):
            cell_val = ws.cell(row=row, column=col_idx).value
            if cell_val is not None:
                lines = str(cell_val).split("\n")
                line_len = max(len(l) for l in lines) if lines else 0
                max_len = max(max_len, line_len)
        calculated_width = max(min_width, min(max_len + 4, max_width))
        ws.column_dimensions[col_letter].width = calculated_width


def calc_page_range(idx: int, total_items: int, total_pages: int = 35):
    """Calculates realistic proportional page ranges for lectures."""
    if total_items <= 0:
        return 1, 1
    start_p = 1 + int((idx - 1) * max(1, total_pages - 1) / total_items)
    end_p = max(start_p, 1 + int(idx * max(1, total_pages - 1) / total_items))
    return start_p, end_p


def get_correct_option_label(correct_text: Any, opt_a: Any, opt_b: Any, opt_c: Any, opt_d: Any) -> str:
    """
    Standardizes the correct answer into one of the four strict option labels:
    'الخيار أ (A)', 'الخيار ب (B)', 'الخيار ج (C)', 'الخيار د (D)'
    """
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
    for label, raw, cl in opts:
        if correct_str == str(raw or '').strip():
            return label
    for label, raw, cl in opts:
        if c_clean and c_clean == cl:
            return label
    for label, raw, cl in opts:
        if c_clean and (c_clean in cl or cl in c_clean):
            return label
    return 'الخيار أ (A)'


# -----------------------------------------------------------------------------
# 1. Official Committee Training Program Sheet Builder
# -----------------------------------------------------------------------------
def add_training_program_sheet(
    wb: openpyxl.Workbook,
    sheet_title: str,
    section_label: str,
    specialty_name: str,
    items: Optional[List[Dict[str, Any]]] = None,
    terms_data: Optional[List[Dict[str, Any]]] = None,
    total_pages: int = 35,
    days_count: Optional[int] = None,
    midterm_day: Optional[int] = None,
    periods_per_day: Optional[int] = None,
    total_hours: Optional[float] = None,
    total_questions: Optional[int] = None
):
    """
    Builds an institutional training program sheet matching the official committee template:
    - Supports single term (Prep: 2 periods/day - ف1 نظري 2h, ف2 عملي 2h).
    - Supports two terms (Med & Final: 4 periods/day - ف1 نظري, ف2 عملي, ف3 نظري, ف4 عملي).
    - Split topic columns (نظري / عملي).
    - Hours (2.0), reference pages (من / الي), and questions count (25).
    - Integrated Midterm and Final exam rows with total formulas/values.
    """
    ws = wb.create_sheet(title=sheet_title)
    apply_rtl_and_grid(ws)

    border_cell = create_thin_border(BORDER_GRAY)
    font_title = Font(name='Times New Roman', size=22, bold=True)
    font_header = Font(name='Times New Roman', size=15, bold=True)
    font_data = Font(name='Times New Roman', size=13, bold=False)
    font_exam = Font(name='Times New Roman', size=15, bold=True)
    font_total = Font(name='Times New Roman', size=15, bold=True)

    header_fill = PatternFill(start_color='E8EEF5', end_color='E8EEF5', fill_type='solid')
    exam_fill = PatternFill(start_color='F2F4F7', end_color='F2F4F7', fill_type='solid')
    total_fill = PatternFill(start_color='DCE6F1', end_color='DCE6F1', fill_type='solid')

    # Row 1: Official Specialty & Section Banner
    ws.merge_cells('A1:I1')
    t_cell = ws['A1']
    clean_sec = section_label.replace("القسم ", "").strip()
    t_cell.value = f"برنامج محاضرات تخصص ( {specialty_name} ) للقسم ( {clean_sec} )"
    t_cell.font = font_title
    t_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 38
    ws.row_dimensions[2].height = 10

    # Normalizing terms data
    if not terms_data:
        # Build single term or auto-detect from section_label / periods_per_day
        p_day = periods_per_day if periods_per_day is not None else (2 if "إعداد" in section_label else 4)
        n_items = len(items or [])
        d_cnt = days_count if days_count is not None else (n_items if p_day == 2 else max(1, (n_items + 1) // 2))
        m_day = midterm_day if midterm_day is not None else max(1, d_cnt // 2)
        terms_data = [{
            "term_name": "الترم الأول",
            "items": items or [],
            "periods_per_day": p_day,
            "days_count": d_cnt,
            "midterm_day": m_day
        }]

    current_row = 3

    for term_idx, t_info in enumerate(terms_data):
        term_name = t_info.get("term_name", f"الترم {term_idx + 1}")
        t_items = t_info.get("items", [])
        p_per_day = t_info.get("periods_per_day", 4)
        t_days = t_info.get("days_count") or (len(t_items) if p_per_day == 2 else max(1, (len(t_items) + 1) // 2))
        t_midterm = t_info.get("midterm_day") or max(1, t_days // 2)

        # If not the first term, add a separator row
        if term_idx > 0:
            ws.row_dimensions[current_row].height = 12
            current_row += 1

        # Header Rows
        r_h1 = current_row
        r_h2 = current_row + 1

        ws.merge_cells(f'A{r_h1}:A{r_h2}')
        ws.merge_cells(f'B{r_h1}:B{r_h2}')
        ws.merge_cells(f'C{r_h1}:C{r_h2}')
        ws.merge_cells(f'D{r_h1}:D{r_h2}')
        ws.merge_cells(f'E{r_h1}:F{r_h1}')
        ws.merge_cells(f'G{r_h1}:H{r_h1}')
        ws.merge_cells(f'I{r_h1}:I{r_h2}')

        ws[f'A{r_h1}'] = 'الترم'
        ws[f'B{r_h1}'] = 'اليوم'
        ws[f'C{r_h1}'] = 'المحاضرة'
        ws[f'D{r_h1}'] = 'اسم الموضوع'
        ws[f'E{r_h1}'] = 'عدد الساعات'
        ws[f'G{r_h1}'] = 'الصفحة في المرجع الموحد'
        ws[f'I{r_h1}'] = 'عدد الاسئلة'

        ws[f'E{r_h2}'] = 'نظري'
        ws[f'F{r_h2}'] = 'عملي'
        ws[f'G{r_h2}'] = 'من'
        ws[f'H{r_h2}'] = 'الي'

        for r in [r_h1, r_h2]:
            ws.row_dimensions[r].height = 26
            for col in range(1, 10):
                c = ws.cell(r, col)
                c.font = font_header
                c.fill = header_fill
                c.alignment = Alignment(horizontal='center', vertical='center')
                c.border = border_cell

        current_row += 2
        term_start_data_row = current_row

        item_idx = 0
        theory_hrs_sum = 0
        prac_hrs_sum = 0
        questions_sum = 0

        for day_num in range(1, t_days + 1):
            day_label = ARABIC_DAYS[day_num - 1] if day_num <= len(ARABIC_DAYS) else f'اليوم {day_num}'
            day_start_row = current_row

            if p_per_day == 2:
                # 2 periods/day (Prep Section structure: 1 topic per day)
                it = t_items[item_idx] if item_idx < len(t_items) else None
                topic_title = (it.get('lesson_name') or it.get('title') if it else f'موضوع تدريبي {item_idx + 1}')
                p_from, p_to = (it.get('p_from', 1), it.get('p_to', 2)) if it and 'p_from' in it else calc_page_range(item_idx + 1, len(t_items), total_pages)

                # Period 1: Theory
                ws.cell(current_row, 3, value='ف1').alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 4, value=topic_title).alignment = Alignment(horizontal='right', vertical='center')
                ws.cell(current_row, 5, value=2).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 6, value='-').alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 7, value=p_from).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 8, value=p_to).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 9, value=25).alignment = Alignment(horizontal='center', vertical='center')
                theory_hrs_sum += 2
                questions_sum += 25
                ws.row_dimensions[current_row].height = 25
                for col in range(1, 10):
                    ws.cell(current_row, col).font = font_data
                    ws.cell(current_row, col).border = border_cell
                current_row += 1

                # Period 2: Practical
                ws.cell(current_row, 3, value='ف2').alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 4, value=topic_title).alignment = Alignment(horizontal='right', vertical='center')
                ws.cell(current_row, 5, value='-').alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 6, value=2).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 7, value=p_from).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 8, value=p_to).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 9, value='-').alignment = Alignment(horizontal='center', vertical='center')
                prac_hrs_sum += 2
                ws.row_dimensions[current_row].height = 25
                for col in range(1, 10):
                    ws.cell(current_row, col).font = font_data
                    ws.cell(current_row, col).border = border_cell
                current_row += 1

                item_idx += 1

                # Merge Day column (Col B) for 2 rows
                ws.merge_cells(start_row=day_start_row, start_column=2, end_row=day_start_row + 1, end_column=2)
                day_cell = ws.cell(day_start_row, 2, value=day_label)
                day_cell.font = font_header
                day_cell.alignment = Alignment(horizontal='center', vertical='center')
                for r_b in range(day_start_row, day_start_row + 2):
                    ws.cell(r_b, 2).border = border_cell

            else:
                # 4 periods/day (Medium & Final Section structure: 2 topics per day)
                # Topic 1: ف1 (نظري), ف2 (عملي)
                it1 = t_items[item_idx] if item_idx < len(t_items) else None
                topic1 = (it1.get('lesson_name') or it1.get('title') if it1 else f'موضوع تدريبي {item_idx + 1}')
                p_from1, p_to1 = (it1.get('p_from', 1), it1.get('p_to', 2)) if it1 and 'p_from' in it1 else calc_page_range(item_idx + 1, len(t_items), total_pages)
                item_idx += 1

                # Period 1: Theory
                ws.cell(current_row, 3, value='ف1').alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 4, value=topic1).alignment = Alignment(horizontal='right', vertical='center')
                ws.cell(current_row, 5, value=2).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 6, value='-').alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 7, value=p_from1).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 8, value=p_to1).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 9, value=25).alignment = Alignment(horizontal='center', vertical='center')
                theory_hrs_sum += 2
                questions_sum += 25
                ws.row_dimensions[current_row].height = 25
                for col in range(1, 10):
                    ws.cell(current_row, col).font = font_data
                    ws.cell(current_row, col).border = border_cell
                current_row += 1

                # Period 2: Practical
                ws.cell(current_row, 3, value='ف2').alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 4, value=topic1).alignment = Alignment(horizontal='right', vertical='center')
                ws.cell(current_row, 5, value='-').alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 6, value=2).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 7, value=p_from1).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 8, value=p_to1).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 9, value='-').alignment = Alignment(horizontal='center', vertical='center')
                prac_hrs_sum += 2
                ws.row_dimensions[current_row].height = 25
                for col in range(1, 10):
                    ws.cell(current_row, col).font = font_data
                    ws.cell(current_row, col).border = border_cell
                current_row += 1

                # Topic 2: ف3 (نظري), ف4 (عملي)
                it2 = t_items[item_idx] if item_idx < len(t_items) else None
                topic2 = (it2.get('lesson_name') or it2.get('title') if it2 else f'موضوع تدريبي {item_idx + 1}')
                p_from2, p_to2 = (it2.get('p_from', 1), it2.get('p_to', 2)) if it2 and 'p_from' in it2 else calc_page_range(item_idx + 1, len(t_items), total_pages)
                item_idx += 1

                # Period 3: Theory
                ws.cell(current_row, 3, value='ف3').alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 4, value=topic2).alignment = Alignment(horizontal='right', vertical='center')
                ws.cell(current_row, 5, value=2).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 6, value='-').alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 7, value=p_from2).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 8, value=p_to2).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 9, value=25).alignment = Alignment(horizontal='center', vertical='center')
                theory_hrs_sum += 2
                questions_sum += 25
                ws.row_dimensions[current_row].height = 25
                for col in range(1, 10):
                    ws.cell(current_row, col).font = font_data
                    ws.cell(current_row, col).border = border_cell
                current_row += 1

                # Period 4: Practical
                ws.cell(current_row, 3, value='ف4').alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 4, value=topic2).alignment = Alignment(horizontal='right', vertical='center')
                ws.cell(current_row, 5, value='-').alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 6, value=2).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 7, value=p_from2).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 8, value=p_to2).alignment = Alignment(horizontal='center', vertical='center')
                ws.cell(current_row, 9, value='-').alignment = Alignment(horizontal='center', vertical='center')
                prac_hrs_sum += 2
                ws.row_dimensions[current_row].height = 25
                for col in range(1, 10):
                    ws.cell(current_row, col).font = font_data
                    ws.cell(current_row, col).border = border_cell
                current_row += 1

                # Merge Day column (Col B) for 4 rows
                ws.merge_cells(start_row=day_start_row, start_column=2, end_row=day_start_row + 3, end_column=2)
                day_cell = ws.cell(day_start_row, 2, value=day_label)
                day_cell.font = font_header
                day_cell.alignment = Alignment(horizontal='center', vertical='center')
                for r_b in range(day_start_row, day_start_row + 4):
                    ws.cell(r_b, 2).border = border_cell

            # Midterm Exam row after midterm day
            if day_num == t_midterm:
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

        # Final Exam row for this term
        ws.row_dimensions[current_row].height = 28
        ws.merge_cells(start_row=current_row, start_column=2, end_row=current_row, end_column=4)
        f_cell = ws.cell(current_row, 2, value='امتحان ختامى الترم')
        f_cell.font = font_total
        f_cell.alignment = Alignment(horizontal='center', vertical='center')

        # Total Theory Hours in Col E
        c_th = ws.cell(current_row, 5, value=theory_hrs_sum)
        c_th.font = font_total
        c_th.alignment = Alignment(horizontal='center', vertical='center')

        # Total Practical Hours in Col F
        c_pr = ws.cell(current_row, 6, value=prac_hrs_sum)
        c_pr.font = font_total
        c_pr.alignment = Alignment(horizontal='center', vertical='center')

        # Page range merged dash in Col G:H
        ws.merge_cells(start_row=current_row, start_column=7, end_row=current_row, end_column=8)
        c_dash = ws.cell(current_row, 7, value='-')
        c_dash.font = font_total
        c_dash.alignment = Alignment(horizontal='center', vertical='center')

        # Total Questions in Col I
        c_qs = ws.cell(current_row, 9, value=questions_sum)
        c_qs.font = font_total
        c_qs.alignment = Alignment(horizontal='center', vertical='center')

        for col in range(1, 10):
            c = ws.cell(current_row, col)
            c.fill = total_fill
            c.border = border_cell

        # Merge Term column A from start of term data rows down to final exam row
        ws.merge_cells(start_row=term_start_data_row, start_column=1, end_row=current_row, end_column=1)
        a_term = ws.cell(term_start_data_row, 1, value=term_name)
        a_term.font = font_header
        a_term.alignment = Alignment(horizontal='center', vertical='center')
        for r_a in range(term_start_data_row, current_row + 1):
            ws.cell(r_a, 1).border = border_cell

        current_row += 1

    # Standard Column widths matching official template
    ws.column_dimensions['A'].width = 14.0
    ws.column_dimensions['B'].width = 15.0
    ws.column_dimensions['C'].width = 14.0
    ws.column_dimensions['D'].width = 55.0
    ws.column_dimensions['E'].width = 10.0
    ws.column_dimensions['F'].width = 13.0
    ws.column_dimensions['G'].width = 8.0
    ws.column_dimensions['H'].width = 13.0
    ws.column_dimensions['I'].width = 14.0


# -----------------------------------------------------------------------------
# 2. Institutional Question Bank Sheet Builder (with Data Validations)
# -----------------------------------------------------------------------------
def add_question_bank_sheet(
    wb: openpyxl.Workbook,
    sheet_title: str,
    questions: List[Dict[str, Any]],
    section_name: str = "",
    equipment_name: str = ""
):
    """
    Populates an LMS-compliant Question Bank sheet strictly adhering to the 10-column layout:
    [نوع السؤال | نص السؤال | الشرح / التفسير | مستوى الصعوبة | الإجابة الصحيحة | الخيار أ (A) | الخيار ب (B) | الخيار ج (C) | الخيار د (D) | الدرس]
    Includes Data Validation on Columns A, D, E.
    """
    if len(sheet_title) > 31:
        sheet_title = sheet_title[:31]

    ws = wb.create_sheet(title=sheet_title)
    apply_rtl_and_grid(ws)

    border_cell = create_thin_border("A0A0A0")
    header_font = Font(name="Times New Roman", size=14, bold=True, color="000000")
    data_font = Font(name="Times New Roman", size=12, bold=False, color="000000")
    header_fill = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")

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
    ws.row_dimensions[1].height = 28

    for col_idx in range(1, 11):
        c = ws.cell(1, col_idx)
        c.font = header_font
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.fill = header_fill
        c.border = border_cell

    for r_idx, q in enumerate(questions, start=2):
        q_type = str(q.get("question_type") or "اختيار من متعدد").strip()
        if q_type not in ["اختيار من متعدد", "صواب أو خطأ", "صواب وخطأ"]:
            q_type = "اختيار من متعدد"
        if q_type == "صواب وخطأ":
            q_type = "صواب أو خطأ"

        q_text = str(q.get("question_text") or "").strip()
        explanation = str(q.get("explanation") or "").strip()
        difficulty = str(q.get("difficulty") or "متوسط").strip()
        if difficulty not in ["سهل", "متوسط", "صعب"]:
            difficulty = "متوسط"

        opt_a = str(q.get("option_a") or "").strip()
        opt_b = str(q.get("option_b") or "").strip()
        opt_c = str(q.get("option_c") or "").strip()
        opt_d = str(q.get("option_d") or "").strip()

        # Format T/F options
        if q_type == "صواب أو خطأ":
            opt_a = "صواب"
            opt_b = "خطأ"
            opt_c = ""
            opt_d = ""

        # Normalize correct answer to label
        raw_correct = q.get("correct_answer") or q.get("correct_label") or ""
        correct_label = get_correct_option_label(raw_correct, opt_a, opt_b, opt_c, opt_d)
        lesson_name = str(q.get("lesson") or "").strip()

        row_vals = [
            q_type,
            q_text,
            explanation,
            difficulty,
            correct_label,
            opt_a,
            opt_b,
            opt_c,
            opt_d,
            lesson_name
        ]
        ws.append(row_vals)
        ws.row_dimensions[r_idx].height = 22

        for col_idx in range(1, 11):
            c = ws.cell(r_idx, col_idx)
            c.font = data_font
            c.border = border_cell

            # Diff highlight
            if col_idx == 4:
                diff_hex = DIFFICULTY_COLORS.get(difficulty, "FFF2CC")
                c.fill = PatternFill(start_color=diff_hex, end_color=diff_hex, fill_type="solid")

            # Alignments
            if col_idx in [1, 4, 5]:
                c.alignment = Alignment(horizontal="center", vertical="center")
                if col_idx == 5:
                    c.font = Font(name="Times New Roman", size=12, bold=True, color="003366")
            else:
                c.alignment = Alignment(horizontal="right", vertical="center", wrap_text=True)

    # Column widths
    widths = {1: 18, 2: 52, 3: 36, 4: 15, 5: 18, 6: 24, 7: 24, 8: 24, 9: 24, 10: 45}
    for col_idx, w in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    # Attach Excel Data Validations
    max_r = max(len(questions) + 100, 300)

    dv_type = DataValidation(type="list", formula1='"اختيار من متعدد,صواب أو خطأ"', allow_blank=True)
    ws.add_data_validation(dv_type)
    dv_type.add(f"A2:A{max_r}")

    dv_diff = DataValidation(type="list", formula1='"سهل,متوسط,صعب"', allow_blank=True)
    ws.add_data_validation(dv_diff)
    dv_diff.add(f"D2:D{max_r}")

    dv_ans = DataValidation(type="list", formula1='"الخيار أ (A),الخيار ب (B),الخيار ج (C),الخيار د (D)"', allow_blank=True)
    ws.add_data_validation(dv_ans)
    dv_ans.add(f"E2:E{max_r}")


# -----------------------------------------------------------------------------
# 3. Specialty Training Program Builder (برنامج تدريب تخصص)
# -----------------------------------------------------------------------------
def build_specialty_training_program_sheet(
    ws,
    specialty_name: str,
    sections_data: List[Dict[str, Any]]
):
    """
    Populates 'برنامج تدريب تخصص' sheet exactly matching the official template:
    - Deleted electrical column.
    - Headers: م | اسم الموضوع | السنة الدراسية | عدد الساعات (نظري / عملي) | يصلح للامتحان المنظومة
    - Section banner over each section across A:F.
    - Numbering restarts at 1 for each section.
    - Subtotals for First and Second Terms with =SUM(...) formulas.
    - Section Grand Total row with sum formula.
    - Note row across A:F.
    """
    apply_rtl_and_grid(ws)

    font_title = Font(name="Times New Roman", size=32, bold=True, color="000000")
    font_header = Font(name="Times New Roman", size=18, bold=True, color="000000")
    font_section_banner = Font(name="Times New Roman", size=20, bold=True, color="000000")
    font_term_total = Font(name="Times New Roman", size=16, bold=True, color="000000")
    font_data = Font(name="Times New Roman", size=16, bold=True, color="000000")
    font_note = Font(name="Times New Roman", size=16, bold=True, color="000000")
    font_col_f = Font(name="Times New Roman", size=24, bold=False, color="000000")

    fill_header = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    fill_section_banner = PatternFill(start_color="B8CCE4", end_color="B8CCE4", fill_type="solid")
    fill_term_total = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    fill_total = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")

    thin_side = Side(border_style="thin", color="000000")
    border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

    # Column widths
    widths = {1: 10, 2: 75, 3: 18, 4: 16, 5: 16, 6: 30}
    for col_idx, w in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    # Row 1: Title
    ws.row_dimensions[1].height = 60
    ws.merge_cells("A1:F1")
    cell_a1 = ws.cell(1, 1, f"برنامج تدريب تخصص ( {specialty_name} )")
    cell_a1.font = font_title
    cell_a1.alignment = Alignment(horizontal="center", vertical="center")
    for c in range(1, 7):
        ws.cell(1, c).border = border_all

    # Rows 2 & 3: Headers
    ws.row_dimensions[2].height = 33
    ws.row_dimensions[3].height = 38.25

    ws.merge_cells("A2:A3")
    ws.cell(2, 1, "م")

    ws.merge_cells("B2:B3")
    ws.cell(2, 2, "اسم الموضوع")

    ws.merge_cells("C2:C3")
    ws.cell(2, 3, "السنة الدراسية")

    ws.merge_cells("D2:E2")
    ws.cell(2, 4, "عدد الساعات")

    ws.cell(3, 4, "نظرى ")
    ws.cell(3, 5, "عملي")

    ws.merge_cells("F2:F3")
    ws.cell(2, 6, "يصلح للامتحان المنظومة")

    for r in [2, 3]:
        for c in range(1, 7):
            cell = ws.cell(r, c)
            cell.font = font_header
            cell.fill = fill_header
            cell.border = border_all
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    current_row = 4

    for sec in sections_data:
        sec_name = sec.get("name", "")
        sec_banner = sec.get("banner_title") or f"موضوعات برنامج تدريب ( {sec_name} ) - تخصص {sec.get('equipment_name', specialty_name)}"
        class_label = sec.get("class_label") or ("الإعدادي" if "إعداد" in sec_name else ("المتوسط" if "متوسط" in sec_name else "النهائي"))
        tot_label = sec.get("total_label") or f"إجمالي {sec_name}"

        # 1. Section Banner
        ws.row_dimensions[current_row].height = 35.0
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
        cell_banner = ws.cell(current_row, 1, sec_banner)
        cell_banner.font = font_section_banner
        cell_banner.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, 7):
            b_cell = ws.cell(current_row, c)
            b_cell.fill = fill_section_banner
            b_cell.border = border_all
        current_row += 1

        sec_topics = sec.get("topics", [])
        serial_num = 1

        # Group topics by term: { 'الترم الأول': [...], 'الترم الثاني': [...] }
        terms_grouped = OrderedDict()
        if isinstance(sec_topics, dict):
            for (term, topic), hours in sec_topics.items():
                t_key = term if term else "الترم الأول"
                terms_grouped.setdefault(t_key, []).append((topic, hours))
        elif isinstance(sec_topics, list):
            for item in sec_topics:
                t_key = item.get("term", "الترم الأول")
                t_name = item.get("topic") or item.get("lesson_name") or item.get("title") or ""
                t_hrs = {"th": item.get("th", 2), "pr": item.get("pr", 2)}
                terms_grouped.setdefault(t_key, []).append((t_name, t_hrs))

        term_total_rows = []
        has_multiple_terms = len(terms_grouped) > 1

        for term_name, term_items in terms_grouped.items():
            start_term_row = current_row

            for topic, hours in term_items:
                ws.row_dimensions[current_row].height = 30.75

                c_a = ws.cell(current_row, 1, serial_num)
                c_a.alignment = Alignment(horizontal="center", vertical="center")

                c_b = ws.cell(current_row, 2, topic)
                c_b.alignment = Alignment(horizontal="right", vertical="center")

                c_c = ws.cell(current_row, 3, class_label)
                c_c.alignment = Alignment(horizontal="center", vertical="center")

                th_val = hours.get("th", 2) if isinstance(hours, dict) else 2
                pr_val = hours.get("pr", 2) if isinstance(hours, dict) else 2

                c_d = ws.cell(current_row, 4, th_val if th_val > 0 else "-")
                c_d.alignment = Alignment(horizontal="center", vertical="center")

                c_e = ws.cell(current_row, 5, pr_val if pr_val > 0 else "-")
                c_e.alignment = Alignment(horizontal="center", vertical="center")

                for col_idx in range(1, 7):
                    cell = ws.cell(current_row, col_idx)
                    cell.font = font_data
                    cell.border = border_all

                current_row += 1
                serial_num += 1

            end_term_row = current_row - 1

            # Merge Col F across this term: يصلح لعدد ساعات النظرى فقط
            ws.merge_cells(start_row=start_term_row, start_column=6, end_row=end_term_row, end_column=6)
            cell_f = ws.cell(start_term_row, 6, "يصلح لعدد ساعات النظرى فقط")
            cell_f.font = font_col_f
            cell_f.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            if has_multiple_terms:
                # Term Subtotal Row
                ws.row_dimensions[current_row].height = 30.75
                ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
                c_term_label = ws.cell(current_row, 1, f"إجمالي ساعات {term_name}")
                c_term_label.alignment = Alignment(horizontal="center", vertical="center")

                c_term_th = ws.cell(current_row, 4, f"=SUM(D{start_term_row}:D{end_term_row})")
                c_term_th.alignment = Alignment(horizontal="center", vertical="center")

                c_term_pr = ws.cell(current_row, 5, f"=SUM(E{start_term_row}:E{end_term_row})")
                c_term_pr.alignment = Alignment(horizontal="center", vertical="center")

                for c in range(1, 7):
                    cell = ws.cell(current_row, c)
                    cell.font = font_term_total
                    cell.fill = fill_term_total
                    cell.border = border_all

                term_total_rows.append(current_row)
                current_row += 1

        # Section Grand Total Row
        ws.row_dimensions[current_row].height = 30.75
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
        c_tot_label = ws.cell(current_row, 1, tot_label)
        c_tot_label.alignment = Alignment(horizontal="center", vertical="center")

        if has_multiple_terms:
            th_terms_str = "+".join([f"D{r}" for r in term_total_rows])
            pr_terms_str = "+".join([f"E{r}" for r in term_total_rows])
            c_tot_th = ws.cell(current_row, 4, f"={th_terms_str}")
            c_tot_pr = ws.cell(current_row, 5, f"={pr_terms_str}")
        else:
            c_tot_th = ws.cell(current_row, 4, f"=SUM(D{start_term_row}:D{end_term_row})")
            c_tot_pr = ws.cell(current_row, 5, f"=SUM(E{start_term_row}:E{end_term_row})")

        c_tot_th.alignment = Alignment(horizontal="center", vertical="center")
        c_tot_pr.alignment = Alignment(horizontal="center", vertical="center")

        for c in range(1, 7):
            cell = ws.cell(current_row, c)
            cell.font = font_header
            cell.fill = fill_total
            cell.border = border_all

        current_row += 1

        # Note Row
        ws.row_dimensions[current_row].height = 30.75
        ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=6)
        c_note = ws.cell(current_row, 1, "ملحوظة : جميع ساعات العملي للتخصص لا يمكن الاختبار فيها عن طريق المنظومة ويتم اختبار النظري فقط.")
        c_note.font = font_note
        c_note.alignment = Alignment(horizontal="center", vertical="center")
        for c in range(1, 7):
            ws.cell(current_row, c).border = border_all

        current_row += 1


def build_specialty_program_workbook(
    specialty_name: str,
    sections_data: List[Dict[str, Any]]
) -> io.BytesIO:
    """Creates standalone 'برنامج تدريب تخصص.xlsx' workbook."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "التخصص"
    build_specialty_training_program_sheet(ws, specialty_name, sections_data)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    return out


# -----------------------------------------------------------------------------
# 4. Master Multi-Sheet Workbook Builder
# -----------------------------------------------------------------------------
def build_workbook(
    syllabus_data: Optional[Dict[str, Any]],
    question_banks: List[Dict[str, Any]],
    equipment_name: str,
    training_program_mode: bool = False,
    sections_setup: Optional[List[Dict[str, Any]]] = None
) -> io.BytesIO:
    """
    Builds the complete multi-sheet Excel workbook:
    - If training_program_mode is True (Scenario 4):
      Creates institutional training program sheets (Prep 2-period, Med & Final 4-period, multi-term),
      Question Banks (Col E option labels + data validations), and embeds the 'برنامج تدريب تخصص' sheet.
    - Otherwise (Scenarios 1-3):
      Sheet 1: الجدول الزمني وفهرس الدروس (Strict Single-Row: 1 Lecture : 2 Hours : 25 Questions)
      Sheets 2..N: بنوك الأسئلة القياسية.
    """
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    wb.properties.creator = "Eslam Abdelbadea"
    wb.properties.lastModifiedBy = "Eslam Abdelbadea"
    wb.properties.title = f"منظومة التدريب وبنوك الأسئلة - {equipment_name}"
    wb.properties.subject = "LMS Question Bank & Technical Syllabus Guide"
    wb.properties.description = f"{BRAND_NOTICE} | {FOOTER_NOTICE}"

    font_banner = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    fill_banner = PatternFill(start_color="0D233A", end_color="0D233A", fill_type="solid")
    border_cell = create_thin_border()

    all_items = syllabus_data.get("syllabus_items", []) if syllabus_data else []

    # ----------------------------------------------------
    # SCENARIO 4: INTEGRATED TRAINING & SPECIALTY SYSTEM
    # ----------------------------------------------------
    if training_program_mode:
        specialty_sections_data = []

        if sections_setup and len(sections_setup) > 0:
            cur_offset = 0
            for sec in sections_setup:
                sec_name = sec["name"]
                sec_lec = sec.get("lectures", 14)
                sec_eq = sec.get("equipment_name", equipment_name)
                sec_items = all_items[cur_offset : cur_offset + sec_lec]
                cur_offset += sec_lec

                # Configure terms for this section
                # If section specifies terms in 'terms', use them
                terms_conf = sec.get("terms")
                if not terms_conf:
                    if "إعداد" in sec_name:
                        # Prep: 1 term, 2 periods/day, 14 days
                        terms_conf = [{
                            "term_name": "الترم الأول",
                            "items": sec_items,
                            "periods_per_day": 2,
                            "days_count": len(sec_items),
                            "midterm_day": max(1, len(sec_items) // 2)
                        }]
                    elif "متوسط" in sec_name:
                        # Med: 2 terms (T1: 24 topics, T2: 28 topics)
                        t1_items = sec_items[:24]
                        t2_items = sec_items[24:]
                        terms_conf = [
                            {
                                "term_name": "الترم الأول",
                                "items": t1_items,
                                "periods_per_day": 4,
                                "days_count": max(1, (len(t1_items) + 1) // 2),
                                "midterm_day": 6
                            },
                            {
                                "term_name": "الترم الثاني",
                                "items": t2_items,
                                "periods_per_day": 4,
                                "days_count": max(1, (len(t2_items) + 1) // 2),
                                "midterm_day": 7
                            }
                        ]
                    else:
                        # Final: 2 terms (T1: 24 topics, T2: 20 topics)
                        t1_items = sec_items[:24]
                        t2_items = sec_items[24:]
                        terms_conf = [
                            {
                                "term_name": "الترم الأول",
                                "items": t1_items,
                                "periods_per_day": 4,
                                "days_count": max(1, (len(t1_items) + 1) // 2),
                                "midterm_day": 6
                            },
                            {
                                "term_name": "الترم الثاني",
                                "items": t2_items,
                                "periods_per_day": 4,
                                "days_count": max(1, (len(t2_items) + 1) // 2),
                                "midterm_day": 5
                            }
                        ]

                # 1. Add Training Program Sheet
                add_training_program_sheet(
                    wb=wb,
                    sheet_title=f"برنامج تدريب - {sec_name}",
                    section_label=sec_name,
                    specialty_name=sec_eq,
                    terms_data=terms_conf,
                    total_pages=35
                )

                # Collect data for 'برنامج تدريب تخصص'
                specialty_topics_list = []
                for t in terms_conf:
                    t_name = t.get("term_name", "الترم الأول")
                    for itm in t.get("items", []):
                        specialty_topics_list.append({
                            "term": t_name,
                            "topic": itm.get("lesson_name") or itm.get("title") or "",
                            "th": 2,
                            "pr": 2
                        })

                class_lbl = "الإعدادي" if "إعداد" in sec_name else ("المتوسط" if "متوسط" in sec_name else "النهائي")
                specialty_sections_data.append({
                    "name": sec_name,
                    "equipment_name": sec_eq,
                    "banner_title": f"موضوعات برنامج تدريب ( {sec_name} ) - تخصص {sec_eq}",
                    "class_label": class_lbl,
                    "total_label": f"إجمالي {sec_name}",
                    "topics": specialty_topics_list
                })

        # 2. Add Question Bank Sheets
        for qb in question_banks:
            b_name = qb.get("section_name", "عام")
            b_title = b_name if b_name.startswith("بنك") else f"بنك {b_name}"
            add_question_bank_sheet(
                wb=wb,
                sheet_title=b_title,
                questions=qb.get("questions", []),
                section_name=b_name,
                equipment_name=qb.get("equipment_name") or equipment_name
            )

        # 3. Add 'برنامج تدريب تخصص' Sheet inside master workbook
        if specialty_sections_data:
            ws_spec = wb.create_sheet(title="برنامج تدريب تخصص")
            build_specialty_training_program_sheet(ws_spec, equipment_name, specialty_sections_data)

    # ----------------------------------------------------
    # STANDARD SYLLABUS SHEET (Scenarios 1-3)
    # ----------------------------------------------------
    elif syllabus_data and syllabus_data.get("syllabus_items"):
        ws_syl = wb.create_sheet(title="الجدول الزمني وفهرس الدروس")
        apply_rtl_and_grid(ws_syl)

        ws_syl.merge_cells("A1:F1")
        banner_cell = ws_syl["A1"]
        banner_cell.value = f"الجدول الزمني وفهرس الدروس المعتمد | المنظومة: {equipment_name}"
        banner_cell.font = font_banner
        banner_cell.fill = fill_banner
        banner_cell.alignment = Alignment(horizontal="center", vertical="center")
        ws_syl.row_dimensions[1].height = 28

        headers_syl = [
            "م",
            "عنوان الدرس / المكون الفني",
            "عدد الساعات المخصصة",
            "عدد المحاضرات",
            "عدد الأسئلة المستهدفة",
            "المرجع المعتمد"
        ]
        ws_syl.row_dimensions[2].height = 28
        header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color=NAVY_HEADER, end_color=NAVY_HEADER, fill_type="solid")

        for col_idx, h_text in enumerate(headers_syl, start=1):
            c = ws_syl.cell(row=2, column=col_idx, value=h_text)
            c.font = header_font
            c.fill = header_fill
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = border_cell

        items = syllabus_data.get("syllabus_items", [])
        current_row = 3
        data_font = Font(name="Calibri", size=10)

        for it in items:
            ws_syl.row_dimensions[current_row].height = 24
            row_vals = [
                it.get("id", current_row - 2),
                it.get("lesson_name", ""),
                2,   # STRICTLY 2 hours
                1,   # STRICTLY 1 lecture
                25,  # STRICTLY 25 questions
                it.get("reference", "الدليل الفني المعتمد")
            ]
            for col_idx, val in enumerate(row_vals, start=1):
                c = ws_syl.cell(row=current_row, column=col_idx, value=val)
                c.font = data_font
                c.border = border_cell
                if col_idx in [1, 3, 4, 5]:
                    c.alignment = Alignment(horizontal="center", vertical="center")
                else:
                    c.alignment = Alignment(horizontal="right", vertical="center")
            current_row += 1

        # Summary Row
        ws_syl.row_dimensions[current_row].height = 26
        total_font = Font(name="Calibri", size=10, bold=True, color="000000")
        total_fill = PatternFill(start_color=TOTAL_ROW_FILL, end_color=TOTAL_ROW_FILL, fill_type="solid")

        c_blank = ws_syl.cell(row=current_row, column=1, value="")
        c_blank.fill = total_fill
        c_blank.border = border_cell

        c_total_lbl = ws_syl.cell(row=current_row, column=2, value="الإجمالي العام")
        c_total_lbl.font = total_font
        c_total_lbl.fill = total_fill
        c_total_lbl.alignment = Alignment(horizontal="center", vertical="center")
        c_total_lbl.border = border_cell

        c_hrs = ws_syl.cell(row=current_row, column=3, value=f"=SUM(C3:C{current_row-1})")
        c_hrs.font = total_font
        c_hrs.fill = total_fill
        c_hrs.alignment = Alignment(horizontal="center", vertical="center")
        c_hrs.border = border_cell

        c_lec = ws_syl.cell(row=current_row, column=4, value=f"=SUM(D3:D{current_row-1})")
        c_lec.font = total_font
        c_lec.fill = total_fill
        c_lec.alignment = Alignment(horizontal="center", vertical="center")
        c_lec.border = border_cell

        c_q = ws_syl.cell(row=current_row, column=5, value=f"=SUM(E3:E{current_row-1})")
        c_q.font = total_font
        c_q.fill = total_fill
        c_q.alignment = Alignment(horizontal="center", vertical="center")
        c_q.border = border_cell

        c_ref_total = ws_syl.cell(row=current_row, column=6, value="")
        c_ref_total.fill = total_fill
        c_ref_total.border = border_cell

        ws_syl.column_dimensions["A"].width = 6
        ws_syl.column_dimensions["B"].width = 46
        ws_syl.column_dimensions["C"].width = 20
        ws_syl.column_dimensions["D"].width = 16
        ws_syl.column_dimensions["E"].width = 22
        ws_syl.column_dimensions["F"].width = 30

        # Question banks for Scenarios 1-3
        for qb in question_banks:
            sec_name = qb.get("section_name", "عام")
            sec_eq = qb.get("equipment_name") or equipment_name
            sheet_title = f"بنك {sec_name} - {sec_eq}"
            add_question_bank_sheet(
                wb=wb,
                sheet_title=sheet_title,
                questions=qb.get("questions", []),
                section_name=sec_name,
                equipment_name=sec_eq
            )

    if len(wb.sheetnames) == 0:
        ws_empty = wb.create_sheet(title="تنبيه")
        ws_empty["A1"] = "لم يتم توليد أي محتوى"

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output
