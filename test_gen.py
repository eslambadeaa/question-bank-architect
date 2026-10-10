import sys, io, os, json
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import google.genai as genai
from pydantic import BaseModel, Field
from typing import List

with open('.api_key_config', 'r', encoding='utf-8') as f:
    key = f.read().strip()

client = genai.Client(api_key=key)

class QuestionModel(BaseModel):
    question_type: str = Field(description='اختيار من متعدد أو صواب أو خطأ')
    question_text: str = Field(description='نص السؤال التخصصي')
    explanation: str = Field(description='الشرح والتفسير العلمي')
    difficulty: str = Field(description='سهل أو متوسط أو صعب')
    correct_answer: str = Field(description='الخيار أ (A) أو الخيار ب (B) أو الخيار ج (C) أو الخيار د (D)')
    option_a: str = Field(description='الخيار أ')
    option_b: str = Field(description='الخيار ب')
    option_c: str = Field(description='الخيار ج')
    option_d: str = Field(description='الخيار د')

class BatchModel(BaseModel):
    questions: List[QuestionModel]

prompt = '''أنت خبير نظم تسليح دفاع جوي. قم بصياغة 3 أسئلة تخصصية دقيقة جداً عن موضوع:
[إجراءات التعامل مع عدم خروج الصاروخ بعد الضغطة الثانية في قاذف الضبع الأسود]
بناء على تعليمات الأمان: تثبيت المعدة على الكتف بزاوية لا تقل عن 20 درجة لمدة 2 دقيقة، ثم نزع مج الإطلاق ووضع القاذف على الدروة لمدة 15 دقيقة بعيداً عن الأمام والخلف.
أخرج كائن JSON بالهيكل المحدد.'''

resp = client.models.generate_content(
    model='gemini-2.5-flash',
    contents=prompt,
    config={
        'response_mime_type': 'application/json',
        'response_schema': BatchModel,
        'temperature': 0.2
    }
)

data = json.loads(resp.text)
qs = data.get("questions", [])
print(f"Successfully generated {len(qs)} questions:")
for q in qs:
    print("-", q["question_text"])
    print("  Ans:", q["correct_answer"], "| A:", q["option_a"])
