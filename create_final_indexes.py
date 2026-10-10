import os
import sys
import docx
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import parse_xml
from docx.oxml.ns import nsdecls

sys.stdout.reconfigure(encoding='utf-8')

# RTL XML Helper Functions
def make_table_rtl(table):
    tblPr = table._tbl.tblPr
    tblPr.append(parse_xml('<w:bidiVisual %s/>' % nsdecls('w')))

def set_paragraph_bidi(p):
    pPr = p._p.get_or_add_pPr()
    pPr.append(parse_xml('<w:bidi %s/>' % nsdecls('w')))

def set_cell_background(cell, fill_hex):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = parse_xml('<w:shd %s w:fill="%s"/>' % (nsdecls('w'), fill_hex))
    tcPr.append(shd)

def set_cell_margins(cell, top=100, bottom=100, left=140, right=140):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = parse_xml('<w:tcMar %s><w:top w:w="%s" w:type="dxa"/><w:bottom w:w="%s" w:type="dxa"/><w:left w:w="%s" w:type="dxa"/><w:right w:w="%s" w:type="dxa"/></w:tcMar>' % (nsdecls('w'), top, bottom, left, right))
    tcPr.append(tcMar)

def set_cell_borders(cell, top='single', bottom='single', left='single', right='single', color='D0D0D0', sz='4'):
    tcPr = cell._tc.get_or_add_tcPr()
    tcBorders = parse_xml('<w:tcBorders %s><w:top w:val="%s" w:sz="%s" w:space="0" w:color="%s"/><w:bottom w:val="%s" w:sz="%s" w:space="0" w:color="%s"/><w:left w:val="%s" w:sz="%s" w:space="0" w:color="%s"/><w:right w:val="%s" w:sz="%s" w:space="0" w:color="%s"/></w:tcBorders>' % (nsdecls('w'), top, sz, color, bottom, sz, color, left, sz, color, right, sz, color))
    tcPr.append(tcBorders)

def add_page_number_field(run):
    fldChar1 = parse_xml('<w:fldChar %s w:fldCharType="begin"/>' % nsdecls('w'))
    instrText = parse_xml('<w:instrText %s xml:space="preserve"> PAGE </w:instrText>' % nsdecls('w'))
    fldChar2 = parse_xml('<w:fldChar %s w:fldCharType="separate"/>' % nsdecls('w'))
    fldChar3 = parse_xml('<w:fldChar %s w:fldCharType="end"/>' % nsdecls('w'))
    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

def generate_index_document(system_title, subtitle, rows, output_path):
    doc = docx.Document()
    sec = doc.sections[0]
    sec.top_margin = Inches(0.7)
    sec.bottom_margin = Inches(0.7)
    sec.left_margin = Inches(0.7)
    sec.right_margin = Inches(0.7)

    # Footer
    footer = sec.footer
    p_ftr = footer.paragraphs[0]
    set_paragraph_bidi(p_ftr)
    p_ftr.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_ftr = p_ftr.add_run(f'فهرس مرجع {system_title}  |  صفحة ')
    r_ftr.font.name = 'Arial'
    r_ftr.font.size = Pt(9.5)
    r_ftr.font.color.rgb = RGBColor(0x7F, 0x7F, 0x7F)
    r_page = p_ftr.add_run()
    r_page.font.name = 'Arial'
    r_page.font.size = Pt(9.5)
    r_page.font.bold = True
    add_page_number_field(r_page)

    # Document Header Title
    p_title = doc.add_paragraph()
    set_paragraph_bidi(p_title)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_title.paragraph_format.space_before = Pt(4)
    p_title.paragraph_format.space_after = Pt(2)
    rt = p_title.add_run(f'فهرس الموضوعات الفنية والتدريبية لمعدة {system_title}')
    rt.font.name = 'Arial'
    rt.font.size = Pt(16)
    rt.font.bold = True
    rt.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

    p_sub = doc.add_paragraph()
    set_paragraph_bidi(p_sub)
    p_sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p_sub.paragraph_format.space_after = Pt(12)
    rs = p_sub.add_run(subtitle)
    rs.font.name = 'Arial'
    rs.font.size = Pt(10.5)
    rs.font.bold = True
    rs.font.color.rgb = RGBColor(0x59, 0x59, 0x59)

    # Table: 4 columns: م | الموضوع / المحتوى العلمي في المرجع | من صفحة | إلى صفحة
    table = doc.add_table(rows=1, cols=4)
    make_table_rtl(table)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.autofit = False

    col_widths = [Inches(0.6), Inches(4.8), Inches(0.85), Inches(0.85)]

    # Header Row
    headers = ['م', 'الموضوع / المحتوى في المرجع', 'من صفحة', 'إلى صفحة']
    for idx, (cell, h_text) in enumerate(zip(table.rows[0].cells, headers)):
        cell.width = col_widths[idx]
        set_cell_background(cell, '1F4E78')
        set_cell_margins(cell, top=130, bottom=130, left=140, right=140)
        set_cell_borders(cell, color='1F4E78')
        p = cell.paragraphs[0]
        set_paragraph_bidi(p)
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
        run = p.add_run(h_text)
        run.font.name = 'Arial'
        run.font.size = Pt(11)
        run.font.bold = True
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    # Data Rows
    for item in rows:
        row = table.add_row()
        cells = row.cells
        for idx in range(4):
            cells[idx].width = col_widths[idx]
            bg_color = 'F2F5F9' if item['num'] % 2 == 0 else 'FFFFFF'
            set_cell_background(cells[idx], bg_color)
            set_cell_margins(cells[idx], top=85, bottom=85, left=120, right=120)
            set_cell_borders(cells[idx], color='D9D9D9')

        # Cell 0: م
        p0 = cells[0].paragraphs[0]
        set_paragraph_bidi(p0)
        p0.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p0.paragraph_format.space_before = Pt(0)
        p0.paragraph_format.space_after = Pt(0)
        r0 = p0.add_run(str(item['num']))
        r0.font.name = 'Arial'
        r0.font.size = Pt(10)
        r0.font.bold = True

        # Cell 1: اسم الموضوع المدمج
        p1 = cells[1].paragraphs[0]
        set_paragraph_bidi(p1)
        p1.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        p1.paragraph_format.space_before = Pt(0)
        p1.paragraph_format.space_after = Pt(0)
        r1 = p1.add_run(item['name'])
        r1.font.name = 'Arial'
        r1.font.size = Pt(10)
        r1.font.bold = True

        # If there were multiple original topics merged, list them in light note if helpful
        if " / " in item.get('orig_topics', ''):
            p1_sub = cells[1].add_paragraph()
            set_paragraph_bidi(p1_sub)
            p1_sub.alignment = WD_ALIGN_PARAGRAPH.RIGHT
            p1_sub.paragraph_format.space_before = Pt(2)
            p1_sub.paragraph_format.space_after = Pt(0)
            r1_sub = p1_sub.add_run(f"• الموضوعات المدمجة: {item['orig_topics']}")
            r1_sub.font.name = 'Arial'
            r1_sub.font.size = Pt(8.5)
            r1_sub.font.color.rgb = RGBColor(0x5A, 0x6B, 0x82)

        # Cell 2: من صفحة
        p2 = cells[2].paragraphs[0]
        set_paragraph_bidi(p2)
        p2.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p2.paragraph_format.space_before = Pt(0)
        p2.paragraph_format.space_after = Pt(0)
        r2 = p2.add_run(str(item['p_start']))
        r2.font.name = 'Arial'
        r2.font.size = Pt(10.5)
        r2.font.bold = True
        r2.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

        # Cell 3: إلى صفحة
        p3 = cells[3].paragraphs[0]
        set_paragraph_bidi(p3)
        p3.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p3.paragraph_format.space_before = Pt(0)
        p3.paragraph_format.space_after = Pt(0)
        r3 = p3.add_run(str(item['p_end']))
        r3.font.name = 'Arial'
        r3.font.size = Pt(10.5)
        r3.font.bold = True
        r3.font.color.rgb = RGBColor(0x1F, 0x4E, 0x78)

    try:
        doc.save(output_path)
        print(f"Successfully generated: {output_path}")
    except PermissionError:
        base, ext = os.path.splitext(output_path)
        alt = f"{base}_جديد{ext}"
        doc.save(alt)
        print(f"File locked, saved to alternative: {alt}")

# 1. Black Hyena Data
hyena_rows = [
    {"num": 1, "name": "التركيب العام ومكونات المنظومة والخواص الفنية والتكتيكية للضبع الأسود", "orig_topics": "الخواص الفنية والتكتيكية لمنظومة الضبع الأسود / التركيب العام ومكونات المنظومة", "p_start": 1, "p_end": 2},
    {"num": 2, "name": "مكونات المنظومة الأساسية (الصاروخ، مصدر التغذية الأرضي، مجموعة الإطلاق)", "orig_topics": "الصاروخ ومصدر التغذية الأرضي ومجموعة الإطلاق", "p_start": 2, "p_end": 3},
    {"num": 3, "name": "جداول الخواص الفنية والمواصفات العامة والتكتيكية لمنظومة الضبع الأسود", "orig_topics": "جداول الخواص الفنية والمواصفات العامة", "p_start": 3, "p_end": 4},
    {"num": 4, "name": "احتياطات الأمان الواجب اتباعها وإجراءات التعامل مع الأعطال الميدانية المعقدة للمعدة", "orig_topics": "احتياطات الأمان عند العمل على المعدة / إجراءات التعامل مع الأعطال الميدانية المعقدة", "p_start": 4, "p_end": 5},
    {"num": 5, "name": "الصاروخ وأقسامه الرئيسية والغرض العملياتي", "orig_topics": "الصاروخ وأقسامه الرئيسية والغرض العملياتي", "p_start": 5, "p_end": 6},
    {"num": 6, "name": "رأس التوجيه الذاتي ومكوناته الأساسية", "orig_topics": "رأس التوجيه الذاتي ومكوناته الأساسية", "p_start": 6, "p_end": 7},
    {"num": 7, "name": "العضو الدوار للجيروسكوب ونظرية عمله", "orig_topics": "العضو الدوار للجيروسكوب ونظرية عمله", "p_start": 7, "p_end": 8},
    {"num": 8, "name": "المسار البصري والمرآة المغناطيسية وقرص التعديل وتوليد الإشارة الكهربية بالمقاومة الضوئية", "orig_topics": "المسار البصري والمرآة وقرص التعديل / المقاومة الضوئية وتوليد الإشارة الكهربية", "p_start": 8, "p_end": 9},
    {"num": 9, "name": "ملفات التوليد (البيلنج) والتردد الفعلي لدوران العضو الدوار", "orig_topics": "ملفات التوليد (البيلنج) والتردد الفعلي للدوران", "p_start": 9, "p_end": 10},
    {"num": 10, "name": "وحدة قياس التردد ومكبر الإشارة برأس التوجيه الذاتي", "orig_topics": "وحدة قياس التردد ومكبر الإشارة بالرأس", "p_start": 10, "p_end": 11},
    {"num": 11, "name": "ملفات التصحيح وأنظمة رأس التوجيه الذاتي والتثبيت الميكانيكي والكهربي", "orig_topics": "ملفات التصحيح والتوجيه / نظام التثبيت الميكانيكي والكهربي / أنظمة رأس التوجيه", "p_start": 11, "p_end": 12},
    {"num": 12, "name": "الدليل الآلي ونظرية عمله ومكوناته الأساسية", "orig_topics": "الدليل الآلي ونظرية عمله ومكوناته", "p_start": 12, "p_end": 13},
    {"num": 13, "name": "نظام التوجيه والتتبع الآلي وفكرة عمله والدوائر الإلكترونية لتتبع الهدف", "orig_topics": "الدوائر الإلكترونية لتتبع الهدف بالدليل الآلي / نظام التوجيه والتتبع الآلي", "p_start": 13, "p_end": 14},
    {"num": 14, "name": "الدليل الآلي ودائرة العمل الكهربية وجهاز إنتاج أمر التوجيه لتصحيح المسار", "orig_topics": "الدليل الآلي ودائرة العمل الكهربية ومكوناتها", "p_start": 14, "p_end": 15},
    {"num": 15, "name": "جهاز تحديد السرعة الزاوية وقنطرة هيوستن والجزء الخاص بالدفات ومكوناته", "orig_topics": "جهاز تحديد السرعة الزاوية وقنطرة هيوستن / الجزء الخاص بالدفات ومكوناته ومواصفاته", "p_start": 15, "p_end": 16},
    {"num": 16, "name": "جهاز تحريك الدفات ومستودع الغاز المضغوط", "orig_topics": "جهاز تحريك الدفات ومستودع الغاز المضغوط", "p_start": 16, "p_end": 17},
    {"num": 17, "name": "مصدر التغذية المحمول على سطح الصاروخ وطريقة عمل التوربين الغازي", "orig_topics": "مصدر التغذية المحمول على سطح الصاروخ / طريقة عمل التوربين الغازي", "p_start": 17, "p_end": 18},
    {"num": 18, "name": "القسم القتالي الغرض العام ومكوناته الأساسية", "orig_topics": "القسم القتالي الغرض العام ومكوناته", "p_start": 18, "p_end": 19},
    {"num": 19, "name": "جهاز التفجير ومكوناته ووسائل الأمان وطريقة عمله وحالات التعمير والانفجار الذاتي", "orig_topics": "جهاز التفجير ومكوناته ووسائل الأمان / طريقة عمل جهاز التفجير وحالات التعمير والانفجار الذاتي", "p_start": 19, "p_end": 20},
    {"num": 20, "name": "رأس التدمير والعبوة المتفجرة وتأثير الشظايا ومحركات وأجنحة الصاروخ والحلقتان المركزيتان", "orig_topics": "رأس التدمير ومكوناتها / تأثير الشظايا والموجة الانفجارية / الجزء الخاص بالمحركات والأجنحة", "p_start": 20, "p_end": 21},
    {"num": 21, "name": "المحرك الدافع والمحرك الرئيسي ومواصفاتهما وأزمنة عمل المحركات", "orig_topics": "المحرك الدافع ومواصفاته وطريقة عمله / المحرك الرئيسي وأزمنة عمل المحركات", "p_start": 21, "p_end": 22},
    {"num": 22, "name": "القاذف الغرض منه والمواصفات العامة ومكونات أنبوب القاذف والتجهيزات البصرية", "orig_topics": "القاذف الغرض منه والمواصفات العامة / مكونات أنبوب القاذف والتجهيزات البصرية", "p_start": 22, "p_end": 23},
    {"num": 23, "name": "مصدر التغذية الأرضي ومواصفاته العامة والغرض منه", "orig_topics": "مصدر التغذية الأرضي ومواصفاته العامة", "p_start": 23, "p_end": 24},
    {"num": 24, "name": "بطارية مصدر التغذية الأرضي والتفاعل الكيميائي وطريقة تشغيل المصدر", "orig_topics": "بطارية مصدر التغذية الأرضي والتفاعل الكيميائي", "p_start": 24, "p_end": 25},
    {"num": 25, "name": "مجموعة الإطلاق الغرض والخواص الفنية", "orig_topics": "مجموعة الإطلاق الغرض والخواص الفنية", "p_start": 25, "p_end": 26},
    {"num": 26, "name": "مكونات مجموعة الإطلاق وذراع الأمان ومفاتيح التنشيط والتتك ومسارات التوصيل", "orig_topics": "مكونات مجموعة الإطلاق وذراع الأمان / مفاتيح التنشيط والتتك ومسارات التوصيل", "p_start": 26, "p_end": 27},
    {"num": 27, "name": "التعاون بين عناصر المعدة منذ التشغيل حتى تجهيز الصاروخ للإطلاق", "orig_topics": "التعاون بين عناصر المعدة منذ التشغيل حتى تجهيز الصاروخ", "p_start": 27, "p_end": 28},
    {"num": 28, "name": "تسلسل الإشارات الكهربية والصوتية عند التقاط الهدف", "orig_topics": "تسلسل الإشارات الكهربية والصوتية عند التقاط الهدف", "p_start": 28, "p_end": 29},
    {"num": 29, "name": "وحدة التأخير ووحدة تجهيز جهاز التفجير ومراحل خروج وانطلاق الصاروخ من القاذف", "orig_topics": "وحدة التأخير ووحدة تجهيز جهاز التفجير / مراحل خروج وانطلاق الصاروخ من القاذف", "p_start": 29, "p_end": 30},
    {"num": 30, "name": "مراحل طيران الصاروخ حتى الاصطدام بالهدف وآلية التدمير الذاتي والتطبيق التكتيكي للاشتباك", "orig_topics": "مراحل طيران الصاروخ حتى الاصطدام / آلية التدمير الذاتي / التطبيق العملي التكتيكي المتكامل", "p_start": 30, "p_end": 31},
    {"num": 31, "name": "تكتيكات الرماية الميدانية والاشتباك في بيئة المعركة", "orig_topics": "تكتيكات الرماية الميدانية والاشتباك في بيئة المعركة", "p_start": 31, "p_end": 32},
    {"num": 32, "name": "الصيانة الدورية والتفتيش الفني، وقواعد تخزين الصواريخ واختبار الجاهزية والتقييم الشامل للمعدة", "orig_topics": "الصيانة الدورية والتفتيش الفني / قواعد تخزين الصواريخ / اختبار الجاهزية الفنية / استخلاص النتائج والتقييم الفني", "p_start": 32, "p_end": 33}
]

# 2. Igla Data
igla_rows = [
    {"num": 1, "name": "الخواص الفنية والتكتيكية للنظام الصاروخي إيجلا ومكوناته الأساسية", "orig_topics": "الخواص الفنية والتكتيكية للنظام الصاروخي ايجلا / مكونات النظام الصاروخي ايجلا الأساسية", "p_start": 1, "p_end": 2},
    {"num": 2, "name": "الخواص التكتيكية ومعدلات المناورة والسرعة للصاروخ إيجلا", "orig_topics": "الخواص التكتيكية ومعدلات المناورة والسرعة للصاروخ", "p_start": 2, "p_end": 3},
    {"num": 3, "name": "دور دائرة الإزاحة وتكتيكات الاشتباك الصاروخي ضد الأهداف الجوية في ظروف التشويش", "orig_topics": "دور دائرة الإزاحة في زيادة فاعلية الاشتباك / تكتيكات الاشتباك الصاروخي ضد الطائرات في ظروف التشويش", "p_start": 3, "p_end": 4},
    {"num": 4, "name": "مناطق الإطلاق والتدمير والعوامل المؤثرة عليها", "orig_topics": "مناطق الإطلاق والتدمير والعوامل المؤثرة عليها", "p_start": 4, "p_end": 5},
    {"num": 5, "name": "الوصف التفصيلي لمكونات الصاروخ إيجلا ووحدة التحكم", "orig_topics": "الوصف التفصيلي لمكونات الصاروخ إيجلا", "p_start": 5, "p_end": 6},
    {"num": 6, "name": "رأس التوجيه الذاتي والمسمار الإيروديناميكي وحساسات الرؤية", "orig_topics": "رأس التوجيه الذاتي والمسمار الإيروديناميكي / مكونات رأس التوجيه الذاتي وحساسات الرؤية", "p_start": 6, "p_end": 7},
    {"num": 7, "name": "وحدة منسق التتبع والجيروسكوب ونظرية عمله", "orig_topics": "وحدة منسق التتبع والجيروسكوب وطريقة عمله", "p_start": 7, "p_end": 8},
    {"num": 8, "name": "المقاومات الفوتوغرافية والفلتر المجزأ للإشارات", "orig_topics": "المقاومات الفوتوغرافية والفلتر المجزأ للإشارات", "p_start": 8, "p_end": 9},
    {"num": 9, "name": "مكونات وحدة التحكم الإلكترونية ومسارات الإشارة والمشغل الميكانيكي وأسطح التحكم بالدفات", "orig_topics": "مكونات وحدة التحكم الإلكترونية ومسارات الإشارة / المشغل الميكانيكي وأسطح التحكم بالدفات", "p_start": 9, "p_end": 10},
    {"num": 10, "name": "مصدر التغذية المحمول ومكوناته بالصاروخ وطريقة عمل المولد التوربيني", "orig_topics": "مصدر التغذية المحمول ومكوناته بالصاروخ / طريقة عمل المولد التوربيني لمصدر التغذية المحمول", "p_start": 11, "p_end": 12},
    {"num": 11, "name": "جهاز الإحساس بالتغير الزاوي الغرض والمكونات", "orig_topics": "جهاز الإحساس بالتغير الزاوي الغرض والمكونات", "p_start": 12, "p_end": 13},
    {"num": 12, "name": "مكبر جهاز الإحساس بالتغير الزاوي ووحدة الإمداد بالغاز", "orig_topics": "مكبر جهاز الإحساس بالتغير الزاوي ووحدة الإمداد بالغاز", "p_start": 13, "p_end": 14},
    {"num": 13, "name": "موتور توجيه الصاروخ في المرحلة الابتدائية (SPCM) وطريقة عمله وتغذيته", "orig_topics": "موتور توجيه الصاروخ في المرحلة الابتدائية / طريقة عمل وتغذية موتور التوجيه الابتدائي", "p_start": 14, "p_end": 15},
    {"num": 14, "name": "وحدة التسليح ومراحل الأمان ووحدة إفقاد الاستقرار وآليات التفعيل التكتيكي بالصاروخ", "orig_topics": "وحدة التسليح ومراحل الأمان بالصاروخ / وحدة إفقاد الاستقرار وآليات التفعيل التكتيكي", "p_start": 15, "p_end": 16},
    {"num": 15, "name": "رأس التدمير الغرض والمكونات والعبوة الشديدة الانفجار", "orig_topics": "رأس التدمير الغرض والمكونات والعبوة الشديدة الانفجار", "p_start": 16, "p_end": 17},
    {"num": 16, "name": "وسائل الأمان والتفجير ومولد الانفجار برأس التدمير", "orig_topics": "وسائل الأمان والتفجير ومولد الانفجار بالرأس", "p_start": 17, "p_end": 18},
    {"num": 17, "name": "الوحدة الإضافية للإحساس بالهدف واستجابة الصدمة", "orig_topics": "الوحدة الإضافية للإحساس بالهدف واستجابة الصدمة", "p_start": 18, "p_end": 19},
    {"num": 18, "name": "منظومة الأمان الكهروميكانيكي المتكاملة للصاروخ", "orig_topics": "منظومة الأمان الكهروميكانيكي المتكاملة للصاروخ", "p_start": 19, "p_end": 20},
    {"num": 19, "name": "منظومة المحركات بالصاروخ إيجلا ومكوناتها العامة", "orig_topics": "منظومة المحركات بالصاروخ إيجلا ومكوناتها", "p_start": 20, "p_end": 21},
    {"num": 20, "name": "المحرك الدافع الغرض والمكونات وطريقة العمل", "orig_topics": "المحرك الدافع الغرض والمكونات وطريقة العمل", "p_start": 21, "p_end": 22},
    {"num": 21, "name": "المحرك ثنائي الوظيفة الرافع والحافظ وأزمنة الاحتراق", "orig_topics": "المحرك ثنائي الوظيفة الرافع والحافظ", "p_start": 22, "p_end": 23},
    {"num": 22, "name": "مؤخر عمل الشحنة ووحدة الأجنحة الخلفية الغرض والمكونات والفتح الميكانيكي", "orig_topics": "مؤخر عمل الشحنة والغرض وطريقة العمل / وحدة الأجنحة الخلفية الغرض والمكونات والفتح الميكانيكي", "p_start": 23, "p_end": 24},
    {"num": 23, "name": "عمل مكونات المعدة قبل مغادرة الصاروخ للقاذف", "orig_topics": "عمل مكونات المعدة قبل مغادرة الصاروخ للقاذف", "p_start": 24, "p_end": 25},
    {"num": 24, "name": "عمل مكونات الصاروخ أثناء الطيران والتوجيه والتطبيق الميداني للرماية بمنظومة إيجلا", "orig_topics": "عمل مكونات الصاروخ أثناء الطيران والتوجيه نحو الهدف / التطبيق الميداني التكتيكي المتكامل للرماية", "p_start": 25, "p_end": 27},
    {"num": 25, "name": "مجموعة الأدوات والأجزاء الاحتياطية والتعبئة وصناديق التخزين والحفظ", "orig_topics": "مجموعة الأدوات والأجزاء الاحتياطية والتعبئة والتخزين", "p_start": 27, "p_end": 28},
    {"num": 26, "name": "مجموعة الإطلاق المكونات ونظرية العمل والتلامسات ومؤشرات الإطلاق الصوتية والضوئية", "orig_topics": "مجموعة الإطلاق المكونات ونظرية العمل والتلامسات / السماعة ومؤشرات الإطلاق الصوتية والضوئية", "p_start": 28, "p_end": 29},
    {"num": 27, "name": "خطوات وشروط إطلاق الصاروخ وإجراءات التعامل مع الأعطال وعدم خروج الصاروخ", "orig_topics": "خطوات إطلاق الصاروخ وشروطها بمجموعة الإطلاق / إجراءات التعامل مع الأعطال وعدم خروج الصاروخ", "p_start": 29, "p_end": 30},
    {"num": 28, "name": "مصدر التغذية الأرضي المكونات ونظرية العمل وخطوات التركيب والتفريغ", "orig_topics": "مصدر التغذية الأرضي المكونات ونظرية العمل / تركيب وتوصيل وتفريغ مصدر التغذية الأرضي", "p_start": 30, "p_end": 31},
    {"num": 29, "name": "إجراءات الصيانة الدورية والتفتيش الفني على منظومة إيجلا", "orig_topics": "إجراءات الصيانة الدورية والتفتيش الفني على منظومة إيجلا", "p_start": 31, "p_end": 32},
    {"num": 30, "name": "القاذف وأجهزة التسديد والناشنكاهات البصرية ووحدات التوصيل والتقييم العملياتي الشامل", "orig_topics": "القاذف المكونات ونظرية العمل / أجهزة التسديد والناشنكاهات البصرية / التقييم العملياتي والدروس المستفادة", "p_start": 32, "p_end": 33}
]

# Paths
out_hyena = r'C:\Users\MaximuM-Tech\Downloads\فهرس_مرجع_الضبع_الاسود.docx'
out_igla = r'C:\Users\MaximuM-Tech\Downloads\فهرس_مرجع_الايجلا.docx'

generate_index_document(
    "الضبع الأسود (9K32M)",
    "فهرس تفصيلي للموضوعات الفنية بالمرجع مرتبة تسلسلياً حسب الصفحات (33-1)",
    hyena_rows,
    out_hyena
)

generate_index_document(
    "الإيجلا (9K38 Igla)",
    "فهرس تفصيلي للموضوعات الفنية بالمرجع مرتبة تسلسلياً حسب الصفحات (33-1)",
    igla_rows,
    out_igla
)

print("Both standalone index documents created successfully!")
