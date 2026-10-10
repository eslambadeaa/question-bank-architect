import re, sys

sys.stdout.reconfigure(encoding='utf-8')

with open(r'e:\bank\igla_manual_text.txt', 'r', encoding='utf-8') as f:
    igla_txt = f.read()

keywords_final = [
    'مكونات النظام الصاروخي',
    'الخواص التكتيكية ومعدلات المناورة',
    'حساسات الرؤية',
    'المولد التوربيني',
    'موتور التوجيه الابتدائي',
    'وحدة إفقاد الاستقرار',
    'الإحساس بالهدف واستجابة الصدمة',
    'الأمان الكهروميكانيكي',
    'السماعة ومؤشرات الإطلاق',
    'مصدر التغذية الأرضي'
]

print("=== Search in igla_manual_text.txt ===")
for kw in keywords_final:
    matches = [m.start() for m in re.finditer(re.escape(kw), igla_txt)]
    print(f"Keyword '{kw}': {len(matches)} matches")
    if not matches:
        # try words
        subwords = kw.split()
        for w in subwords:
            m = len(re.findall(re.escape(w), igla_txt))
            print(f"   part '{w}': {m} matches")
