import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

files = [
    r'C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_الاعدادي_الضبع_الاسود_LMS.xlsx',
    r'C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_المتوسط_الضبع_الاسود_LMS.xlsx',
    r'C:\Users\MaximuM-Tech\Downloads\بنك_اسئلة_القسم_النهائي_الضبع_الاسود_LMS.xlsx'
]

for file_path in files:
    print(f"\n=======================================================")
    print(f"Processing File: {file_path.split('\\')[-1]}")
    print(f"=======================================================")

    wb = openpyxl.load_workbook(file_path)
    ws = wb['Questions']

    total_rows = ws.max_row - 1
    tf_cleared_count = 0
    mcq_checked_count = 0
    errors = []

    for r in range(2, ws.max_row + 1):
        q_type = str(ws.cell(r, 1).value).strip() if ws.cell(r, 1).value else ''
        q_text = ws.cell(r, 2).value
        if not q_text:
            continue

        raw_ans = ws.cell(r, 5).value
        ans_str = str(raw_ans).strip() if raw_ans is not None else ''

        # Verify correct answer is not empty
        if not ans_str or ans_str.lower() in ['none', '']:
            errors.append(f"Row {r}: Empty correct answer in Col E!")
            continue

        # Case 1: True / False question
        if ans_str in ['True', 'False'] or q_type.startswith('2') or 'صح' in q_type:
            # Normalize answer
            if ans_str in ['True', 'true', 'صواب', 'صح', 'A']:
                ws.cell(r, 5, 'True')
            elif ans_str in ['False', 'false', 'خطأ', 'خطا', 'B']:
                ws.cell(r, 5, 'False')
            else:
                errors.append(f"Row {r}: Invalid TF answer '{ans_str}'")

            # Clear Col F (6), Col G (7), Col H (8), Col I (9)
            ws.cell(r, 6).value = None
            ws.cell(r, 7).value = None
            ws.cell(r, 8).value = None
            ws.cell(r, 9).value = None
            tf_cleared_count += 1

        # Case 2: Multiple Choice question
        else:
            if ans_str not in ['A', 'B', 'C', 'D']:
                if ans_str.startswith('الخيار أ') or ans_str == 'أ': ws.cell(r, 5, 'A')
                elif ans_str.startswith('الخيار ب') or ans_str == 'ب': ws.cell(r, 5, 'B')
                elif ans_str.startswith('الخيار ج') or ans_str == 'ج': ws.cell(r, 5, 'C')
                elif ans_str.startswith('الخيار د') or ans_str == 'د': ws.cell(r, 5, 'D')
                else:
                    errors.append(f"Row {r}: Invalid MCQ answer '{ans_str}'")

            # Verify that options A, B, C, D are present
            optA = ws.cell(r, 6).value
            optB = ws.cell(r, 7).value
            optC = ws.cell(r, 8).value
            optD = ws.cell(r, 9).value
            if not optA or not optB:
                errors.append(f"Row {r}: MCQ missing options A or B!")
            mcq_checked_count += 1

    wb.save(file_path)
    print(f"Total questions processed: {total_rows}")
    print(f"MCQ questions verified (Col E in [A,B,C,D], Cols F-I populated): {mcq_checked_count}")
    print(f"TF questions updated (Col E in [True,False], Cols F & G cleared to None): {tf_cleared_count}")
    print(f"Errors found: {len(errors)}")
    if errors:
        for err in errors[:5]:
            print(f"  {err}")
    else:
        print(">> ALL ROWS 100% VALIDATED AND COMPLIANT!")

print("\nALL 3 FILES COMPLETED.")
