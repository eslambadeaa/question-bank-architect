"""
Builds 'برنامج تدريب تخصص.xlsx' matching the template layout:
- Deletes the 'مخططات و دوائر كهربائية' column.
- Keeps columns:
    Col A: م
    Col B: اسم الموضوع
    Col C: السنة الدراسية
    Col D: عدد الساعات (نظري)
    Col E: عملي
    Col F: يصلح للامتحان المنظومة
- Populates all 3 sections:
    1. القسم الإعدادي (14 topics - السنة الدراسية: الإعدادية)
    2. القسم المتوسط (52 topics - السنة الدراسية: المتوسطة / الثانية)
    3. القسم النهائي (44 topics - السنة الدراسية: النهائية / الرابعة)
- Formats every row with the authentic Times New Roman fonts, borders, fills, and alignments.
- Includes total summary rows for each section and the general note row.
"""

import sys
import io
import os
import openpyxl
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from collections import OrderedDict

sys.stdout.reconfigure(encoding='utf-8')

# Source official workbook
src_official = r"C:\Users\MaximuM-Tech\Downloads\بنوك\برنامج تدريب تخصص الضبع الاسود_النسخة_الرسمية_المعتمدة.xlsx"
target_file = r"C:\Users\MaximuM-Tech\Downloads\برنامج تدريب تخصص.xlsx"

wb_official = openpyxl.load_workbook(src_official, data_only=True)

def extract_topics(sheet_name):
    ws = wb_official[sheet_name]
    topics = OrderedDict()
    term = ''
    for r in range(5, ws.max_row + 1):
        c1 = ws.cell(r, 1).value
        c4 = ws.cell(r, 4).value
        c5 = ws.cell(r, 5).value
        c6 = ws.cell(r, 6).value
        
        if c1 and 'الترم' in str(c1):
            term = str(c1).strip()
            
        if not c4 or str(c4).strip() in ['None', '', 'امتحان منتصف الترم', 'اسم الموضوع']:
            continue
            
        topic = str(c4).strip()
        h_th = int(c5) if str(c5).isdigit() else 0
        h_pr = int(c6) if str(c6).isdigit() else 0
        
        key = (term, topic)
        if key not in topics:
            topics[key] = {'th': 0, 'pr': 0}
        topics[key]['th'] += h_th
        topics[key]['pr'] += h_pr
    return topics

prep_topics = extract_topics('برنامج تدريب - القسم الإعدادي')
med_topics = extract_topics('برنامج تدريب - القسم المتوسط')
fin_topics = extract_topics('برنامج تدريب - القسم النهائي')

print(f"Extracted: Prep={len(prep_topics)}, Med={len(med_topics)}, Fin={len(fin_topics)}")

# Styles
font_title = Font(name="Times New Roman", size=32, bold=True, color="000000")
font_header = Font(name="Times New Roman", size=18, bold=True, color="000000")
font_sub_header = Font(name="Times New Roman", size=18, bold=True, color="000000")
font_data = Font(name="Times New Roman", size=16, bold=True, color="000000")
font_note = Font(name="Times New Roman", size=16, bold=True, color="000000")
font_col_f = Font(name="Times New Roman", size=24, bold=False, color="000000")

# Fills matching template theme tints
fill_header = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid") # ~25% darker
fill_total = PatternFill(start_color="BFBFBF", end_color="BFBFBF", fill_type="solid")  # ~35% darker

# Borders
thin_side = Side(border_style="thin", color="000000")
border_all = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)

# Create workbook
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "التخصص"
ws.sheet_view.rightToLeft = True
ws.sheet_view.showGridLines = True

# Column widths (A to F)
widths = {
    1: 10,   # م
    2: 75,   # اسم الموضوع
    3: 18,   # السنة الدراسية
    4: 16,   # نظرى
    5: 16,   # عملي
    6: 30    # يصلح للامتحان المنظومة
}
for col_idx, w in widths.items():
    ws.column_dimensions[get_column_letter(col_idx)].width = w

# Row 1: Title
ws.row_dimensions[1].height = 60
ws.merge_cells("A1:F1")
cell_a1 = ws.cell(1, 1, "برنامج تدريب تخصص ( الضبع الأسود / الإيجلا )")
cell_a1.font = font_title
cell_a1.alignment = Alignment(horizontal="center", vertical="center")
for c in range(1, 7):
    ws.cell(1, c).border = border_all

# Row 2 & 3: Headers
ws.row_dimensions[2].height = 33
ws.row_dimensions[3].height = 38.25

# A2:A3 -> م
ws.merge_cells("A2:A3")
ws.cell(2, 1, "م")

# B2:B3 -> اسم الموضوع
ws.merge_cells("B2:B3")
ws.cell(2, 2, "اسم الموضوع")

# C2:C3 -> السنة الدراسية
ws.merge_cells("C2:C3")
ws.cell(2, 3, "السنة الدراسية")

# D2:E2 -> عدد الساعات
ws.merge_cells("D2:E2")
ws.cell(2, 4, "عدد الساعات")

# D3 -> نظرى
ws.cell(3, 4, "نظرى ")

# E3 -> عملي
ws.cell(3, 5, "عملي")

# F2:F3 -> يصلح للامتحان المنظومة
ws.merge_cells("F2:F3")
ws.cell(2, 6, "يصلح للامتحان المنظومة")

# Apply formatting to Header cells
for r in [2, 3]:
    for c in range(1, 7):
        cell = ws.cell(r, c)
        cell.font = font_header
        cell.fill = fill_header
        cell.border = border_all
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

sections_data = [
    {
        "name": "القسم الإعدادي",
        "class_label": "الإعدادية",
        "total_label": "إجمالي القسم الإعدادي (السنة الأولى)",
        "topics": prep_topics
    },
    {
        "name": "القسم المتوسط",
        "class_label": "المتوسطة",
        "total_label": "إجمالي القسم المتوسط (السنة الثانية)",
        "topics": med_topics
    },
    {
        "name": "القسم النهائي",
        "class_label": "النهائية",
        "total_label": "إجمالي القسم النهائي (السنة الثالثة)",
        "topics": fin_topics
    }
]

current_row = 4

for sec_idx, sec in enumerate(sections_data):
    start_sec_row = current_row
    sec_topics = sec["topics"]
    serial_num = 1
    
    for (term, topic), hours in sec_topics.items():
        ws.row_dimensions[current_row].height = 30.75
        
        # Col A: م
        c_a = ws.cell(current_row, 1, serial_num)
        c_a.alignment = Alignment(horizontal="center", vertical="center")
        
        # Col B: اسم الموضوع
        c_b = ws.cell(current_row, 2, topic)
        c_b.alignment = Alignment(horizontal="right", vertical="center")
        
        # Col C: السنة الدراسية
        c_c = ws.cell(current_row, 3, sec["class_label"])
        c_c.alignment = Alignment(horizontal="center", vertical="center")
        
        # Col D: نظرى
        c_d = ws.cell(current_row, 4, hours["th"] if hours["th"] > 0 else "-")
        c_d.alignment = Alignment(horizontal="center", vertical="center")
        
        # Col E: عملي
        c_e = ws.cell(current_row, 5, hours["pr"] if hours["pr"] > 0 else "-")
        c_e.alignment = Alignment(horizontal="center", vertical="center")
        
        # Formatting
        for col_idx in range(1, 7):
            cell = ws.cell(current_row, col_idx)
            cell.font = font_data
            cell.border = border_all
            
        current_row += 1
        serial_num += 1

    end_sec_row = current_row - 1
    
    # Merge Col F across this section: يصلح لعدد ساعات النظرى فقط
    ws.merge_cells(start_row=start_sec_row, start_column=6, end_row=end_sec_row, end_column=6)
    cell_f = ws.cell(start_sec_row, 6, "يصلح لعدد ساعات النظرى فقط")
    cell_f.font = font_col_f
    cell_f.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Summary Row for this section
    ws.row_dimensions[current_row].height = 30.75
    ws.merge_cells(start_row=current_row, start_column=1, end_row=current_row, end_column=3)
    c_tot_label = ws.cell(current_row, 1, sec["total_label"])
    c_tot_label.alignment = Alignment(horizontal="center", vertical="center")
    
    # Formulas for totals
    c_tot_th = ws.cell(current_row, 4, f"=SUM(D{start_sec_row}:D{end_sec_row})")
    c_tot_th.alignment = Alignment(horizontal="center", vertical="center")
    
    c_tot_pr = ws.cell(current_row, 5, f"=SUM(E{start_sec_row}:E{end_sec_row})")
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

# Save target workbook
print(f"Saving updated workbook to: {target_file}")
wb.save(target_file)
print("SUCCESS: Updated successfully!")
