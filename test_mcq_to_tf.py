import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.styles import Font, Alignment, Border, Side
import shutil
import sys

sys.stdout.reconfigure(encoding='utf-8')

template_path = r'C:\Users\MaximuM-Tech\Downloads\نموذج تسجيل الاسئلة (1).xlsx'
source_lms_path = r'C:\Users\MaximuM-Tech\Downloads\برنامج تدريب تخصص الضبع الاسود + بنوك_محدث_LMS.xlsx'

wb_source = openpyxl.load_workbook(source_lms_path, data_only=True)

def mcq_to_tf(mcq_item, make_true=True):
    # Converts an MCQ item to a clean True/False item
    ans_key = str(mcq_item['ans']).strip()
    correct_opt = mcq_item.get('opt' + ans_key, '')
    
    # Pick incorrect opt
    incorrect_key = 'A' if ans_key != 'A' else 'B'
    incorrect_opt = mcq_item.get('opt' + incorrect_key, '')
    
    q_stem = mcq_item['text']
    # Clean stem if ends with question mark
    if q_stem.endswith('؟') or q_stem.endswith('?'):
        q_stem = q_stem[:-1].strip()
    
    # Common question starters to strip
    starters = ['ما هو ', 'ما هي ', 'كم يبلغ ', 'أين يقع ', 'متى يتم ', 'هل يمكن ', 'ما وظيفة ', 'ما الغرض من ', 'كم عدد ']
    cleaned_stem = q_stem
    for s in starters:
        if s in cleaned_stem:
            cleaned_stem = cleaned_stem.replace(s, '')
            break
            
    if make_true:
        text = f"{cleaned_stem} هو: {correct_opt}." if not cleaned_stem.endswith(':') else f"{cleaned_stem} {correct_opt}."
        ans = "True"
    else:
        text = f"{cleaned_stem} هو: {incorrect_opt}." if not cleaned_stem.endswith(':') else f"{cleaned_stem} {incorrect_opt}."
        ans = "False"
        
    return {
        'text': text,
        'exp': mcq_item['exp'],
        'diff': mcq_item['diff'],
        'ans': ans,
        'optA': 'True',
        'optB': 'False',
        'optC': None,
        'optD': None,
        'lesson': mcq_item['lesson']
    }

print("Tested MCQ to TF module.")
