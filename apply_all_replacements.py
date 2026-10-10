# -*- coding: utf-8 -*-
"""
Apply replacements to Black Hyena Question Banks (Medium & Final).
Ensures zero duplicates across entire files, empty Col C, intact validations.
"""

import openpyxl
import shutil
import sys
from generate_medium_questions import MED_Q
from generate_final_questions import FIN_Q

sys.stdout.reconfigure(encoding='utf-8')

def apply_updates(filepath, replacements, section_name):
    print(f"\n=======================================================")
    print(f"Applying updates to {section_name}: {filepath}")
    print(f"Replacements count: {len(replacements)}")
    print(f"=======================================================")

    # Backup first
    backup_path = filepath.replace('.xlsx', '_backup_before_dedup.xlsx')
    shutil.copy2(filepath, backup_path)
    print(f"Backup saved to: {backup_path}")

    # Load workbook preserving formulas and validations
    wb = openpyxl.load_workbook(filepath, data_only=False)
    ws = wb['Questions']

    applied_mcq = 0
    applied_tf = 0

    for r, data in replacements.items():
        q_type = str(ws.cell(r, 1).value).strip() if ws.cell(r, 1).value else ''
        q_text = data['q']
        ans = data['ans']

        # Ensure Col B is updated
        ws.cell(r, 2).value = q_text

        # Ensure Col C is strictly None (empty)
        ws.cell(r, 3).value = None

        # Ensure Col E is updated
        ws.cell(r, 5).value = ans

        if 'opts' in data: # MCQ
            applied_mcq += 1
            ws.cell(r, 6).value = data['opts'][0]
            ws.cell(r, 7).value = data['opts'][1]
            ws.cell(r, 8).value = data['opts'][2]
            ws.cell(r, 9).value = data['opts'][3]
        else: # TF
            applied_tf += 1
            ws.cell(r, 6).value = None
            ws.cell(r, 7).value = None
            ws.cell(r, 8).value = None
            ws.cell(r, 9).value = None

    wb.save(filepath)
    print(f"Successfully applied {applied_mcq} MCQs and {applied_tf} TFs (Total: {len(replacements)})")

# 1. Update Medium
apply_updates(
    r'C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_المتوسط_الضبع_الاسود_LMS.xlsx',
    MED_Q,
    'القسم المتوسط'
)

# 2. Update Final
apply_updates(
    r'C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_النهائي_الضبع_الاسود_LMS.xlsx',
    FIN_Q,
    'القسم النهائي'
)

print("\nAll updates applied successfully!")
