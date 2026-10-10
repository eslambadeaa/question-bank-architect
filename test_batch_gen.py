import sys, io, os, json, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')
import google.genai as genai

with open('.api_key_config', 'r', encoding='utf-8') as f:
    key = f.read().strip()

client = genai.Client(api_key=key)

prompt = """أنت خبير نظم تسليح دفاع جوي وبنوك أسئلة عسكرية.
المطلوب: صياغة 25 سؤالاً دقيقاً ومتبايناً بنسبة 100% (15 اختيار من متعدد و10 صواب أو خطأ) لمعدة: [الضبع الأسود].
الموضوع: [مفاتيح التنشيط والتتك ومسارات التوصيل بمجموعة الإطلاق]
المرجع الفني:
- تتكون مجموعة الأمان من عمود مشطوف وياي وذراع أمان وتيلة، وله وضعان: الوضع C أمان، والوضع B عمل.
- التتك يتكون من جسم ومحور وياي خاص وجلبة أسطوانية، وعند الإطلاق يؤخذ ضغطتان:
  الضغطة الأولى: بعد ظهور إشارتي الصوت والضوء لتحرير الجيروسكوب وتتبع الهدف.
  الضغطة الثانية: لإطلاق الصاروخ.
- ريشة تثبيت التتك تثبته في وضع الضغطة الثانية ويتم فك التثبيت عند نزع مج الإطلاق.
- تلامسات مجموعة الإطلاق: 1 و 2 توفر حرية حركة الجيرو عند الضغطة الأولى، 3 و 4 تعمل على إطلاق الصاروخ عند الضغطة الثانية (جهد 40 فولت)، 5 و 6 و 7 و 8 تعمل على التثبيت الكهربي.

الشروط:
1. توزيع الصعوبة: 10 سهل، 10 متوسط، 5 صعب.
2. حقل الإجابة الصحيحة (correct_answer) يجب أن يكون بالضبط واحداً من:
   "الخيار أ (A)", "الخيار ب (B)", "الخيار ج (C)", "الخيار د (D)".
3. في أسئلة صواب أو خطأ، يكون الخيار أ: "صواب" والخيار ب: "خطأ" والخياران ج و د فارغين.
4. أخرج مصفوفة JSON تحتوي على 25 كائناً بالحقول:
   question_type, question_text, explanation, difficulty, correct_answer, option_a, option_b, option_c, option_d.
"""

for model_name in ['gemini-2.5-flash', 'gemini-2.0-flash']:
    try:
        t0 = time.time()
        resp = client.models.generate_content(
            model=model_name,
            contents=prompt,
            config={'response_mime_type': 'application/json', 'temperature': 0.2}
        )
        data = json.loads(resp.text)
        print(f"{model_name} generated {len(data)} questions in {time.time()-t0:.2f}s")
        print("Sample Q1:", data[0]['question_text'])
        print("Sample Q1 Ans:", data[0]['correct_answer'])
        break
    except Exception as e:
        print(f"{model_name} error: {e}")
