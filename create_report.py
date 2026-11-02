from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_SECTION
from pathlib import Path

OUT = Path(__file__).parent / 'تقرير_مشروع_مطعم_صحة.docx'
GREEN = '315E48'
PALE = 'EAF0E3'
INK = '253A31'
MUTED = '64766C'


def bidi(p, on=True):
    pPr = p._p.get_or_add_pPr()
    tag = qn('w:bidi')
    node = pPr.find(tag)
    if on and node is None:
        pPr.append(OxmlElement('w:bidi'))
    elif not on and node is not None:
        pPr.remove(node)
    return p


def set_run_font(run, name='Arial', size=11, bold=None, color=INK, italic=None):
    run.font.name = name
    run._element.rPr.rFonts.set(qn('w:eastAsia'), name)
    run._element.rPr.rFonts.set(qn('w:cs'), name)
    run.font.size = Pt(size)
    if bold is not None: run.bold = bold
    if italic is not None: run.italic = italic
    if color: run.font.color.rgb = RGBColor.from_string(color)


def shade(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)


def set_cell_text(cell, text, bold=False, color=INK, size=9, align=WD_ALIGN_PARAGRAPH.RIGHT):
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = align
    bidi(p, align == WD_ALIGN_PARAGRAPH.RIGHT or align == WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.space_before = Pt(2)
    r = p.add_run(str(text))
    set_run_font(r, size=size, bold=bold, color=color)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER


def set_repeat_table_header(row):
    trPr = row._tr.get_or_add_trPr()
    tblHeader = OxmlElement('w:tblHeader')
    tblHeader.set(qn('w:val'), 'true')
    trPr.append(tblHeader)


def set_table_rtl(table):
    tblPr = table._tbl.tblPr
    bidi_visual = OxmlElement('w:bidiVisual')
    bidi_visual.set(qn('w:val'), '1')
    tblPr.append(bidi_visual)


def add_page_field(paragraph):
    run = paragraph.add_run()
    begin = OxmlElement('w:fldChar'); begin.set(qn('w:fldCharType'), 'begin')
    instr = OxmlElement('w:instrText'); instr.set(qn('xml:space'), 'preserve'); instr.text = ' PAGE '
    separate = OxmlElement('w:fldChar'); separate.set(qn('w:fldCharType'), 'separate')
    text = OxmlElement('w:t'); text.text = '1'
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end')
    for el in (begin, instr, separate, text, end): run._r.append(el)
    set_run_font(run, size=9, color=MUTED)


def add_toc_field(paragraph):
    run = paragraph.add_run()
    for typ, value in [('begin', None), ('instr', ' TOC \\o "1-2" \\h \\z \\u '), ('separate', None)]:
        if typ == 'begin':
            el = OxmlElement('w:fldChar'); el.set(qn('w:fldCharType'), 'begin')
        elif typ == 'instr':
            el = OxmlElement('w:instrText'); el.set(qn('xml:space'), 'preserve'); el.text = value
        else:
            el = OxmlElement('w:fldChar'); el.set(qn('w:fldCharType'), 'separate')
        run._r.append(el)
    fallback = OxmlElement('w:t'); fallback.text = 'يُحدّث فهرس الصفحات تلقائيًا عند فتح التقرير في Word.'; run._r.append(fallback)
    end = OxmlElement('w:fldChar'); end.set(qn('w:fldCharType'), 'end'); run._r.append(end)
    set_run_font(run, size=10, color=MUTED)


def add_para(doc, text='', style=None, size=11, bold=False, color=INK, align=WD_ALIGN_PARAGRAPH.RIGHT, after=6, before=0, italic=False):
    p = doc.add_paragraph(style=style)
    p.alignment = align
    bidi(p, align == WD_ALIGN_PARAGRAPH.RIGHT or align == WD_ALIGN_PARAGRAPH.CENTER)
    p.paragraph_format.space_after = Pt(after)
    p.paragraph_format.space_before = Pt(before)
    p.paragraph_format.line_spacing = 1.15
    if text:
        r = p.add_run(text)
        set_run_font(r, size=size, bold=bold, color=color, italic=italic)
    return p


def add_heading(doc, text, level=1):
    p = doc.add_paragraph(style=f'Heading {level}')
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    bidi(p)
    r = p.add_run(text)
    sizes = {1:20, 2:14, 3:12}
    set_run_font(r, size=sizes.get(level,12), bold=True, color=GREEN)
    p.paragraph_format.space_before = Pt(12 if level == 1 else 8)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.keep_with_next = True
    return p


def add_bullet(doc, text):
    p = doc.add_paragraph(style='List Bullet')
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    bidi(p)
    p.paragraph_format.space_after = Pt(3)
    r = p.add_run(text); set_run_font(r, size=10.5, color=INK)
    return p


def add_code(doc, code):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    bidi(p, False)
    p.paragraph_format.left_indent = Inches(.18)
    p.paragraph_format.right_indent = Inches(.18)
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.keep_together = True
    pPr = p._p.get_or_add_pPr()
    shd = OxmlElement('w:shd'); shd.set(qn('w:fill'), 'F1F3EE'); pPr.append(shd)
    r = p.add_run(code)
    set_run_font(r, name='Consolas', size=9, color='31483A')
    return p


doc = Document()
section = doc.sections[0]
section.top_margin = Inches(.68); section.bottom_margin = Inches(.65)
section.left_margin = Inches(.72); section.right_margin = Inches(.72)
section.page_width = Inches(8.27); section.page_height = Inches(11.69)
section.header_distance = Inches(.32); section.footer_distance = Inches(.32)

# Global style defaults
normal = doc.styles['Normal']
normal.font.name = 'Arial'; normal.font.size = Pt(11); normal.font.color.rgb = RGBColor.from_string(INK)
normal._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial'); normal._element.rPr.rFonts.set(qn('w:cs'), 'Arial')
for sname, pts in [('Heading 1',20),('Heading 2',14),('Heading 3',12)]:
    st = doc.styles[sname]; st.font.name = 'Arial'; st.font.size = Pt(pts); st.font.bold = True; st.font.color.rgb = RGBColor.from_string(GREEN)
    st._element.rPr.rFonts.set(qn('w:eastAsia'), 'Arial'); st._element.rPr.rFonts.set(qn('w:cs'), 'Arial')

# Header and live page-number field
header = section.header.paragraphs[0]
header.alignment = WD_ALIGN_PARAGRAPH.RIGHT; bidi(header)
r = header.add_run('صِحّة  |  SMART HEALTHY FOOD  —  تقرير المشروع'); set_run_font(r, size=8, color=MUTED)
footer = section.footer.paragraphs[0]
footer.alignment = WD_ALIGN_PARAGRAPH.CENTER; bidi(footer)
r = footer.add_run('مادة تطوير تطبيقات الأجهزة الذكية   •   صفحة '); set_run_font(r, size=9, color=MUTED)
add_page_field(footer)

# Cover page
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; bidi(p); p.paragraph_format.space_before = Pt(54); p.paragraph_format.space_after = Pt(12)
r = p.add_run('✳ صِحّة'); set_run_font(r, size=34, bold=True, color=GREEN)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; bidi(p); p.paragraph_format.space_after = Pt(7)
r = p.add_run('SMART HEALTHY FOOD'); set_run_font(r, name='DM Sans', size=15, bold=True, color='7E9C79')
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; bidi(p); p.paragraph_format.space_before = Pt(16); p.paragraph_format.space_after = Pt(11)
r = p.add_run('تقرير مشروع الموقع الإلكتروني'); set_run_font(r, size=27, bold=True, color=INK)
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER; bidi(p); p.paragraph_format.space_after = Pt(19)
r = p.add_run('مطعم الوجبات الصحية — HTML5 · CSS · JavaScript'); set_run_font(r, size=14, color=MUTED)
cover = doc.add_table(rows=4, cols=2); cover.alignment = WD_TABLE_ALIGNMENT.CENTER; cover.autofit = False
cover.columns[0].width = Inches(2.2); cover.columns[1].width = Inches(3.7)
for row,(key,val) in enumerate([('المادة','تطوير تطبيقات الأجهزة الذكية'),('الجامعة','يُستكمل باسم الجامعة'),('المجموعة','يُستكمل باسم المجموعة'),('أسماء الطالبات','تُستكمل الأسماء في جدول توزيع المهام')]):
    set_cell_text(cover.cell(row,0), key, bold=True, color=GREEN, size=10); shade(cover.cell(row,0),PALE)
    set_cell_text(cover.cell(row,1), val, size=10)
set_table_rtl(cover)
add_para(doc,'ملخص المشروع',size=14,bold=True,color=GREEN,align=WD_ALIGN_PARAGRAPH.CENTER,before=22,after=8)
add_para(doc,'موقع عربي متجاوب لمطعم الوجبات الصحية، يتكون من عشر صفحات مستقلة ومترابطة. يعرض قائمة الطعام والقيم الغذائية، ويضم نماذج تسجيل وطلب وحاسبة سعر ضمن تجربة سهلة الاستخدام.',size=12,align=WD_ALIGN_PARAGRAPH.CENTER,after=15)
add_para(doc,'نسخة قابلة للتعديل — تُستكمل أسماء الطالبات والجامعة واسم المجموعة قبل التسليم النهائي.',size=10,color=MUTED,align=WD_ALIGN_PARAGRAPH.CENTER,italic=True)
add_para(doc,'مخرجات المشروع',size=13,bold=True,color=GREEN,align=WD_ALIGN_PARAGRAPH.CENTER,before=14,after=5)
add_para(doc,'عشر صفحات HTML مترابطة  •  ملف CSS خارجي متجاوب  •  تفاعلات JavaScript للتحقق والحساب  •  تقرير توزيع وخطوات وأكواد',size=11,align=WD_ALIGN_PARAGRAPH.CENTER,after=10)
doc.add_page_break()

# Contents page
add_heading(doc,'محتويات التقرير',1)
add_para(doc,'فهرس تلقائي لعناوين التقرير (يمكن تحديثه في Word بالنقر بزر الفأرة الأيمن ثم «تحديث الحقل»).',size=10,color=MUTED)
p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.RIGHT; bidi(p); add_toc_field(p)
add_heading(doc,'أقسام التقرير',2)
for text in ['1. فكرة المشروع وأهدافه','2. بنية الموقع والصفحات العشر','3. توزيع المهام على الطالبات','4. تفاصيل عمل كل طالبة وخطوات صفحتها','5. خطوات إنشاء الصفحات المشتركة','6. الأكواد والمهارات المستخدمة','7. تشغيل الموقع والتحقق قبل التسليم']:
    add_bullet(doc,text)
add_heading(doc,'تنبيه قبل التسليم',2)
add_para(doc,'بيانات الجامعة واسم المجموعة وأسماء الطالبات غير مذكورة في التكليف. تركنا خانات واضحة قابلة للتحرير في الجدول التالي لإدخالها من المجموعة قبل طباعة التقرير.',size=10.5)
doc.add_page_break()

add_heading(doc,'1. فكرة المشروع وأهدافه',1)
add_para(doc,'«صِحّة» موقع تعريفي وخدمي لمطعم وجبات صحية. يقدم تجربة عربية من اليمين إلى اليسار لعرض الوجبات والأسعار والقيم الغذائية، ويتيح التسجيل وإعداد ملخص طلب التوصيل وحساب قيمة الأصناف. تتحقق النماذج من البيانات وتعرض الملخص، لكنها لا ترسل البيانات إلى خادم ولا تنفذ دفعًا حقيقيًا.')
add_heading(doc,'أهداف الموقع',2)
for text in ['إنشاء صفحات HTML5 مستقلة وربطها باستخدام عنصر الرابط a.','توظيف العناوين والفقرات والصور والجداول والقوائم والنماذج.','فصل التنسيق في ملف CSS خارجي موحد للموقع.','استخدام JavaScript للتحقق من المدخلات وإجراء العمليات الحسابية وعرض الرسائل.','تقسيم المشروع إلى مهام واضحة، صفحة رئيسية واحدة لكل طالبة.']:
    add_bullet(doc,text)
add_heading(doc,'2. بنية الموقع والصفحات العشر',1)
page_rows=[('01','index.html','الرئيسية','هوية المطعم، صورة رئيسية، روابط الأقسام ورسالة ترحيب.'),('02','about.html','من نحن','الرسالة والأهداف والقوائم وتدرج العناوين h1 إلى h6.'),('03','nutrition.html','القيم الغذائية','جدول أسعار وسعرات وبروتين وكربوهيدرات ودهون.'),('04','register.html','إنشاء حساب','حقول نص وبريد وكلمة مرور وقائمة منسدلة وملاحظات.'),('05','order.html','طلب توصيل','بيانات العميل والوجبة والكمية والعنوان وملخص السعر.'),('06','calculator.html','حاسبة الطلب','حساب لحظي للكميات والإجمالي برسائل شرطية.'),('07','bowls.html','الأطباق','عرض أطباق الأوعية الصحية وأسعارها وصورها.'),('08','wraps.html','الراب','عرض الراب والسندويشات الكاملة.'),('09','smoothies.html','المشروبات','عرض السموذي والمشروبات الطازجة.'),('10','snacks.html','الوجبات الخفيفة','عرض الوجبات الخفيفة ومعلوماتها.')]
t=doc.add_table(rows=1, cols=4); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.style='Table Grid'
set_table_rtl(t)
for i,h in enumerate(['التسلسل','الملف','الصفحة','المحتوى']): set_cell_text(t.rows[0].cells[i],h,bold=True,color='FFFFFF',size=9); shade(t.rows[0].cells[i],GREEN)
set_repeat_table_header(t.rows[0])
for row in page_rows:
    cells=t.add_row().cells
    for i,val in enumerate(row): set_cell_text(cells[i],val,size=8.5)

doc.add_page_break()
add_heading(doc,'3. جدول توزيع المهام على الطالبات',1)
add_para(doc,'تُسند صفحة أساسية لكل طالبة. اكتبي اسم الجامعة واسم المجموعة واسم الطالبة في الخانات؛ تبقى المهمة والصفحة ثابتتين ويمكن تعديلهما حسب توزيع المجموعة الفعلي.',size=10.5)
assignments=[
('01','الرئيسية — index.html','إعداد هيكل HTML5 العام، محتوى التعريف، الصورة الرئيسية، وروابط الانتقال إلى أقسام الموقع.'),
('02','من نحن — about.html','صياغة قصة المطعم ورسالته وأهدافه، وتنظيم العناوين h1–h6 والفقرات والقائمة ul.'),
('03','القيم الغذائية — nutrition.html','إنشاء جدول HTML منظم للأسعار والسعرات والبروتين والكربوهيدرات والدهون.'),
('04','التسجيل — register.html','بناء نموذج تسجيل يتضمن حقول text وemail وpassword وselect وtextarea والإقرار.'),
('05','طلب توصيل — order.html','بناء نموذج العميل والوجبة والكمية والعنوان وموعد التوصيل والملاحظات.'),
('06','الحاسبة — calculator.html','إنشاء واجهة اختيار الأصناف وربطها بحساب JavaScript للكمية والإجمالي.'),
('07','الأطباق — bowls.html','عرض فئة الأطباق المتوازنة بالصور والوصف والأسعار وروابط الطلب.'),
('08','الراب — wraps.html','عرض الراب والسندويشات ووصف المكونات والسعر وربط صفحة الطلب.'),
('09','المشروبات — smoothies.html','عرض المشروبات والسموذي بصورة ووصف وسعر لكل صنف.'),
('10','الوجبات الخفيفة — snacks.html','عرض الوجبات الخفيفة وربطها بجدول المعلومات والحاسبة ونموذج الطلب.')]
t=doc.add_table(rows=1, cols=5); t.alignment=WD_TABLE_ALIGNMENT.CENTER; t.style='Table Grid'; t.autofit=False
set_table_rtl(t)
widths=[.62,1.28,1.15,1.15,2.15]
for i,w in enumerate(widths): t.columns[i].width=Inches(w)
for i,h in enumerate(['الرقم','الجامعة','اسم المجموعة','اسم الطالبة','المهمة والصفحة']): set_cell_text(t.rows[0].cells[i],h,bold=True,color='FFFFFF',size=8); shade(t.rows[0].cells[i],GREEN)
set_repeat_table_header(t.rows[0])
for num,task,detail in assignments:
    cells=t.add_row().cells
    vals=[num,'يُستكمل','يُستكمل','يُستكمل',task+'\n'+detail]
    for i,val in enumerate(vals): set_cell_text(cells[i],val,size=7.6)
    if int(num)%2==0:
        for c in cells: shade(c,'F5F7F2')

add_heading(doc,'4. تفاصيل عمل كل طالبة وخطوات صفحتها',1)
add_para(doc,'يوضح هذا الجزء مساهمة كل عضوة، وخطوات بناء الصفحة المسندة إليها، ومقتطفًا من الكود المستخدم. أسماء العضوات الفعلية تُضاف في جدول التوزيع بعد اعتماد المجموعة.',size=10.5)
individual_work=[
('01 — الرئيسية','جهزت صفحة الدخول إلى الموقع، وعرّفت هوية المطعم وقدمت صورة الطعام وأزرار الانتقال للطلب والقيم الغذائية.','بناء هيكل HTML5؛ إضافة عنوان تعريفي وفقرة؛ إدراج الصورة مع alt؛ ربط أزرار الدعوة بالصفحات الداخلية.','<main>\n  <h1>لذيذ ومتوازن</h1>\n  <img src="meal.webp" alt="وعاء صحي بالخضار">\n  <a href="order.html">اطلبي الآن</a>\n</main>'),
('02 — من نحن','كتبت قصة المطعم ورسالته ورؤيته وأهدافه، ورتبت المحتوى بالعناوين والقوائم.','كتابة فقرات تعريفية؛ ترتيب العناوين؛ تحويل الأهداف إلى قائمة ul؛ إضافة روابط ذات صلة.','<h2>رسالتنا</h2>\n<p>وجبات طازجة ومعلومة واضحة.</p>\n<ul><li>مكونات طازجة</li><li>اختيار متوازن</li></ul>'),
('03 — القيم الغذائية','أنشأت جدولًا يقارن الأسعار والسعرات والعناصر الغذائية لوجبات مختارة.','تحديد أسماء الأعمدة؛ إضافة thead وtbody؛ إدخال بيانات كل وجبة؛ ضبط عناوين الأعمدة باستخدام th.','<table>\n  <thead><tr><th>الصنف</th><th>السعرات</th></tr></thead>\n  <tbody><tr><td>وعاء الدجاج</td><td>420</td></tr></tbody>\n</table>'),
('04 — التسجيل','جهزت نموذجًا لإنشاء حساب مع مدخلات الاسم والبريد وكلمة المرور والهدف والملاحظات.','إنشاء form؛ إقران label بكل حقل؛ إضافة input وselect وtextarea؛ ربط التحقق برسالة نجاح أو خطأ.','<input type="text" name="fullName">\n<input type="password" name="password">\n<select name="goal"><option>وجبات متوازنة</option></select>\n<textarea name="preferences"></textarea>'),
('05 — طلب التوصيل','أنشأت نموذج بيانات المستلمة والوجبة والكمية والعنوان، مع ملخص قيمة الطلب ورسوم التوصيل.','إنشاء حقول العميل؛ إضافة خيارات وجبات وأسعارها؛ ربط الكمية بالمجموع؛ التحقق من الهاتف والكمية والبريد والعنوان قبل القبول.','<label for="order-meal">الوجبة</label>\n<select id="order-meal"><option data-price="42">وعاء السلمون</option></select>\n<input id="order-quantity" type="number" value="1">'),
('06 — الحاسبة','بنت صفحة تفاعل تجمع أسعار الأصناف وتعرض عددها وإجماليها مع رسالة تتغير بحسب قيمة السلة.','تهيئة المتغيرات؛ المرور على الأصناف بحلقة for؛ جمع السعر مضروبًا بالكمية؛ إظهار نتيجة if/else في الصفحة.','var total = 0;\nfor (var i = 0; i < rows.length; i++) { total += price[i] * quantity[i]; }\nif (total >= 100) { message = "طلب كبير"; }\nelse { message = "أضيفي صنفًا"; }'),
('07 — أطباق متوازنة','عرضت فئة الأطباق في بطاقات تضم اسم الوجبة وصورتها ووصفًا موجزًا وسعرًا ورابطًا للطلب.','إنشاء بطاقة article؛ إدراج صورة alt؛ إضافة اسم ووصف وسعر؛ ربط زر البطاقة بنموذج الطلب.','<article class="meal-card">\n  <img src="bowls.webp" alt="طبق صحي">\n  <h3>وعاء الدجاج الأخضر</h3><a href="order.html">اطلبيها</a>\n</article>'),
('08 — الراب والسندويشات','جهزت صفحة فئة الراب بحشوات متنوعة، مع أسعار وأوصاف وروابط لبقية أجزاء الموقع.','نسخ بنية البطاقة المشتركة؛ تغيير أسماء الأصناف والأوصاف؛ تعيين صورة الراب؛ ربط البطاقة بصفحة الطلب والبيانات.','<h1>الراب والسندويشات</h1>\n<img src="wraps.webp" alt="راب حبوب كاملة">\n<a href="nutrition.html">القيم الغذائية</a>'),
('09 — المشروبات الطازجة','عرضت أصناف السموذي والمشروبات مع صورة مناسبة ووصف المكونات والسعر لكل خيار.','اختيار بيانات الفئة؛ كتابة اسم ووصف كل مشروب؛ إدراج صورة مع بديل وصفي؛ إضافة رابط الطلب.','<article>\n  <img src="smoothies.webp" alt="سموذي توت طازج">\n  <h3>سموذي التوت</h3><p>توت وزبادي وموز.</p>\n</article>'),
('10 — الوجبات الخفيفة','أكملت فئات القائمة بوجبات خفيفة، وربطت محتواها بجدول القيم والحاسبة ونموذج الطلب.','إنشاء بطاقات الأصناف؛ تحديد الوصف والسعر؛ إضافة رابط حساب الطلب؛ مراجعة التنقل الموحد مع بقية الصفحات.','<article class="meal-card">\n  <h3>كرات الطاقة</h3><p>تمر وشوفان وكاكاو.</p>\n  <a href="calculator.html">احسبي طلبك</a>\n</article>')]
for student, contribution, steps_text, snippet in individual_work:
    add_heading(doc,student,3)
    add_para(doc,'ما أُنجز: '+contribution,size=9.5,after=3)
    add_para(doc,'خطوات الإنشاء: '+steps_text,size=9.5,after=3)
    add_para(doc,'مثال الكود:',size=9,color=MUTED,after=1)
    add_code(doc,snippet)

add_heading(doc,'5. خطوات إنشاء الصفحات المشتركة',1)
steps=[('الخطوة 1 — التخطيط','حصر الصفحات العشر، وتحديد روابط التنقل، ومحتوى كل صفحة، وصور الوجبات والبيانات التي ستعرضها.'),('الخطوة 2 — بناء الهيكل','إنشاء ملف HTML مستقل لكل صفحة. إضافة <!doctype html>، وتحديد اللغة العربية والاتجاه dir="rtl"، ثم بناء الرأس والتنقل والمحتوى والتذييل.'),('الخطوة 3 — ربط الصفحات','إضافة روابط a href إلى الصفحات في قائمة التنقل والتذييل، والتأكد من أن كل اسم ملف يطابق مسار الرابط.'),('الخطوة 4 — إدراج الوسائط والبيانات','إضافة الصور باستخدام img مع وصف بديل alt، وبناء جدول HTML للقيم الغذائية وقوائم ul للأهداف وروابط الأقسام.'),('الخطوة 5 — إنشاء النماذج','إضافة حقول النص والبريد وكلمة المرور، والقائمة select، وtextarea، وعناصر label لتوضيح كل حقل للمستخدمة.'),('الخطوة 6 — التنسيق الخارجي','كتابة القواعد في styles.css، وتطبيق نظام الألوان والخطوط والمحاذاة والمسافات وتصميم مرن للهاتف.'),('الخطوة 7 — إضافة JavaScript','ربط app.js. إضافة التحقق من النموذج والحساب عند تغيير الكميات ورسالة النجاح وحلقة for والشرط if...else.'),('الخطوة 8 — المراجعة والتسليم','تشغيل الموقع بخادم محلي، وفحص الروابط والحقول وحساب السعر والعرض على الهاتف، ثم إكمال أسماء المجموعة في التقرير.')]
for title,desc in steps:
    add_heading(doc,title,2); add_para(doc,desc,size=10.5)

add_heading(doc,'6. الأكواد والمهارات المستخدمة',1)
add_heading(doc,'أ. هيكل صفحة HTML5 ورابط وصورة',2)
add_code(doc,'<!doctype html>\n<html lang="ar" dir="rtl">\n  <head><link rel="stylesheet" href="styles.css"></head>\n  <body>\n    <a href="nutrition.html">القيم الغذائية</a>\n    <img src="images/meal.jpg" alt="وعاء صحي بالخضار الطازجة">\n  </body>\n</html>')
add_para(doc,'في المشروع تُستخدم الروابط بين الصفحات العشر، ويُحمّل ملف CSS خارجي. الصور الحية في نسخة الموقع مخزنة عبر مسارات وسائط المشروع.',size=10)
add_heading(doc,'ب. قائمة وجدول ونموذج',2)
add_code(doc,'<ul><li>مكونات طازجة</li><li>وجبات متوازنة</li></ul>\n<table><thead><tr><th>الصنف</th><th>السعرات</th></tr></thead>\n<tbody><tr><td>وعاء الدجاج</td><td>420</td></tr></tbody></table>\n<input type="text" name="fullName">\n<input type="password" name="password">\n<select name="goal"><option>وجبات متوازنة</option></select>\n<textarea name="notes"></textarea>')
add_heading(doc,'ج. قواعد CSS المطلوبة',2)
add_code(doc,'body { font-family: Arial, sans-serif; font-size: 16px; color: #19352c;\n       background-color: #f7f6ef; text-align: right; direction: rtl; }\n.hero-image { float: left; }\n.float-note { float: left; font-style: italic; }')
add_para(doc,'تُستخدم في styles.css خصائص font-family وfont-size وfont-style وcolor وbackground-color وtext-align وfloat، مع قواعد استجابة للشاشات الأصغر.',size=10)
add_heading(doc,'د. JavaScript للمتغيرات والشرط والحلقة',2)
add_code(doc,'var total = 0;\nfor (var i = 0; i < rows.length; i++) {\n  total = total + price[i] * quantity[i];\n}\nif (total >= 100) {\n  message = "طلبك يتجاوز 100 ريال";\n} else {\n  message = "أضيفي أصنافًا حسب رغبتك";\n}\nwindow.alert("تم التحقق من التسجيل بنجاح");')
add_para(doc,'التطبيق الفعلي في app.js يجمع أسعار الأطباق والكميات، ويعرض إجماليًا لحظيًا، ويتحقق من حقول التسجيل والطلب ويظهر رسائل بالعربية.',size=10)

add_heading(doc,'7. تشغيل الموقع والتحقق قبل التسليم',1)
add_para(doc,'يتكون المشروع من ملفات HTML مستقلة، وملف styles.css، وملف app.js. صور الطعام تُستدعى عبر مسارات الوسائط التي توفرها معاينة المشروع. لا يحتاج نموذج الطلب أو التسجيل إلى قاعدة بيانات؛ ولا ترسل النماذج البيانات إلى خادم أو تحفظها؛ كما لا تتم معالجة دفع.')
add_code(doc,'1. افتحي مجلد المشروع في الطرفية.\n2. شغلي: python3 -m http.server 3000\n3. افتحي: http://localhost:3000\n4. تنقلي بين صفحات الموقع من القائمة والتذييل.\n5. جرّبي نموذج التسجيل والحاسبة ونموذج الطلب قبل العرض.')
add_heading(doc,'مراجعة التسليم',2)
for text in ['أكملي اسم الجامعة واسم المجموعة وأسماء الطالبات في الجدول.','تأكدي من ظهور عشر صفحات HTML وروابط التنقل والصور.','جرّبي رسالة رفض الحقول الفارغة ورسالة القبول في نموذج التسجيل.','غيّري الكميات في الحاسبة وتأكدي من تحديث الإجمالي ورسالة الشرط.','تحققي من أن نموذج الطلب يعرض رسوم التوصيل والإجمالي بطريقة صحيحة.','وضحي أن القيم الغذائية تقريبية وأن النماذج لا ترسل الطلب أو تعالج الدفع.']:
    add_bullet(doc,text)
add_para(doc,'تم إعداد هذا التقرير لشرح مشروع المادة وتوزيع العمل المقترح؛ يعتمد التوزيع النهائي على موافقة المجموعة وإشراف المعلمة.',size=9,color=MUTED,before=8)

doc.core_properties.title='تقرير مشروع مطعم صحة — الوجبات الصحية'
doc.core_properties.subject='مشروع HTML5 وCSS وJavaScript لمادة تطوير تطبيقات الأجهزة الذكية'
doc.core_properties.author='مجموعة مشروع صِحّة'
doc.save(OUT)
print(OUT)
