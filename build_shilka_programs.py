# -*- coding: utf-8 -*-
"""
توليد برنامجي التدريب لشيلكا (القسم المتوسط 52 درساً والقسم النهائي 44 درساً)
بشكل احترافي يطابق المعايير العسكرية وجداول الساعات المعتمدة
"""

import os
import sys
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from shilka_lessons_data import INTERMEDIATE_LESSONS, FINAL_LESSONS

if sys.stdout:
    sys.stdout.reconfigure(encoding='utf-8')

def style_program_sheet(ws, title, stage_name, lessons, total_hours, practical_hours):
    # Colors
    header_fill = PatternFill(start_color="1E3A8A", end_color="1E3A8A", fill_type="solid") # Dark Navy
    sub_fill = PatternFill(start_color="3B82F6", end_color="3B82F6", fill_type="solid") # Blue
    section_fill = PatternFill(start_color="DBEAFE", end_color="DBEAFE", fill_type="solid") # Light Blue
    total_fill = PatternFill(start_color="FEF3C7", end_color="FEF3C7", fill_type="solid") # Warm Yellow
    zebra_fill = PatternFill(start_color="F8FAFC", end_color="F8FAFC", fill_type="solid")

    font_title = Font(name="Calibri", size=15, bold=True, color="FFFFFF")
    font_header = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    font_section = Font(name="Calibri", size=11, bold=True, color="1E3A8A")
    font_data = Font(name="Calibri", size=10, bold=False, color="000000")
    font_bold = Font(name="Calibri", size=10, bold=True, color="000000")

    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )
    thick_bottom = Border(bottom=Side(style='medium', color='1E3A8A'))

    align_center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    align_right = Alignment(horizontal='right', vertical='center', wrap_text=True)

    # Row 1: Title
    ws.merge_cells("A1:G1")
    ws["A1"] = f"برنامج تدريب تخصص ( الشيلكا ZSU-23-4 ) - {stage_name}"
    ws["A1"].font = font_title
    ws["A1"].fill = header_fill
    ws["A1"].alignment = align_center
    ws.row_dimensions[1].height = 40

    # Row 2 & 3: Table Headers
    ws.merge_cells("A2:A3")
    ws["A2"] = "م"
    ws.merge_cells("B2:B3")
    ws["B2"] = "اسم الموضوع"
    ws.merge_cells("C2:C3")
    ws["C2"] = "السنة الدراسية"
    ws.merge_cells("D2:E2")
    ws["D2"] = "عدد الساعات"
    ws["D3"] = "نظرى"
    ws["E3"] = "عملي"
    ws.merge_cells("F2:F3")
    ws["F2"] = "يصلح للامتحان المنظومة"
    ws.merge_cells("G2:G3")
    ws["G2"] = "عدد الأسئلة"

    for cell_ref in ["A2", "B2", "C2", "D2", "D3", "E3", "F2", "G2"]:
        cell = ws[cell_ref]
        cell.font = font_header
        cell.fill = sub_fill
        cell.alignment = align_center

    ws.row_dimensions[2].height = 24
    ws.row_dimensions[3].height = 24

    # Row 4: Section Banner
    ws.merge_cells("A4:G4")
    ws["A4"] = f"موضوعات برنامج تدريب ( {stage_name} ) - تخصص الشيلكا ZSU-23-4"
    ws["A4"].font = font_section
    ws["A4"].fill = section_fill
    ws["A4"].alignment = align_center
    ws.row_dimensions[4].height = 28

    # Rows 5+: Lessons
    curr_r = 5
    for idx, (code, topic) in enumerate(lessons, start=1):
        ws.cell(curr_r, 1, idx).alignment = align_center
        ws.cell(curr_r, 2, topic).alignment = align_right
        ws.cell(curr_r, 3, stage_name).alignment = align_center
        ws.cell(curr_r, 4, 2).alignment = align_center # 2 hours theory
        ws.cell(curr_r, 5, 2).alignment = align_center # 2 hours practical
        ws.cell(curr_r, 6, "يصلح لعدد ساعات النظرى فقط" if idx == 1 else "").alignment = align_center
        ws.cell(curr_r, 7, 25).alignment = align_center # 25 Qs per lesson

        fill = zebra_fill if idx % 2 == 0 else PatternFill(fill_type=None)
        for c in range(1, 8):
            cell = ws.cell(curr_r, c)
            cell.font = font_data
            cell.border = thin_border
            if fill.fill_type: cell.fill = fill

        ws.row_dimensions[curr_r].height = 22
        curr_r += 1

    # Total Row
    ws.merge_cells(f"A{curr_r}:C{curr_r}")
    ws.cell(curr_r, 1, f"إجمالي {stage_name}").alignment = align_center
    ws.cell(curr_r, 4, len(lessons) * 2).alignment = align_center # Total theory
    ws.cell(curr_r, 5, len(lessons) * 2).alignment = align_center # Total practical
    ws.cell(curr_r, 6, "").alignment = align_center
    ws.cell(curr_r, 7, len(lessons) * 25).alignment = align_center # Total Qs

    for c in range(1, 8):
        cell = ws.cell(curr_r, c)
        cell.font = font_bold
        cell.fill = total_fill
        cell.border = thin_border

    ws.row_dimensions[curr_r].height = 26

    # Column widths
    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 58
    ws.column_dimensions['C'].width = 18
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 26
    ws.column_dimensions['G'].width = 14

    ws.views.sheetView[0].rightToLeft = True

def build_training_programs():
    out_dir = r"e:\bank"
    
    # 1. Intermediate Program
    wb_med = openpyxl.Workbook()
    ws_med = wb_med.active
    ws_med.title = "التخصص"
    style_program_sheet(
        ws_med,
        "برنامج تدريب تخصص الشيلكا",
        "القسم المتوسط",
        INTERMEDIATE_LESSONS,
        total_hours=104,
        practical_hours=104
    )
    med_path = os.path.join(out_dir, "برنامج_تدريب_تخصص_شيلكا_القسم_المتوسط.xlsx")
    wb_med.save(med_path)
    print(f"Created: {med_path}")

    # 2. Final Program
    wb_fin = openpyxl.Workbook()
    ws_fin = wb_fin.active
    ws_fin.title = "التخصص"
    style_program_sheet(
        ws_fin,
        "برنامج تدريب تخصص الشيلكا",
        "القسم النهائي",
        FINAL_LESSONS,
        total_hours=88,
        practical_hours=88
    )
    fin_path = os.path.join(out_dir, "برنامج_تدريب_تخصص_شيلكا_القسم_النهائي.xlsx")
    wb_fin.save(fin_path)
    print(f"Created: {fin_path}")

if __name__ == "__main__":
    build_training_programs()
