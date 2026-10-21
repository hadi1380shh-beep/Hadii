#!/usr/bin/env python3
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

def set_rtl(p):
    pPr=p._p.get_or_add_pPr()
    if pPr.find(qn('w:bidi')) is None:
        b=OxmlElement('w:bidi'); b.set(qn('w:val'),'1'); pPr.append(b)
    jc=pPr.find(qn('w:jc'))
    if jc is None:
        jc=OxmlElement('w:jc'); pPr.append(jc)
    jc.set(qn('w:val'),'right')
    p.alignment=WD_ALIGN_PARAGRAPH.RIGHT
def set_font(run,name="B Nazanin",size=13,bold=False,color=None,italic=False):
    run.font.name=name; run.font.size=Pt(size); run.bold=bold; run.italic=italic
    if color: run.font.color.rgb=RGBColor.from_string(color)
    r=run._element; rPr=r.get_or_add_rPr()
    rFonts=rPr.find(qn('w:rFonts'))
    if rFonts is None: rFonts=OxmlElement('w:rFonts'); rPr.append(rFonts)
    rFonts.set(qn('w:ascii'),name); rFonts.set(qn('w:hAnsi'),name); rFonts.set(qn('w:cs'),name); rFonts.set(qn('w:eastAsia'),name)
def sp(p,before=0,after=6,ls=1.45):
    pf=p.paragraph_format; pf.space_before=Pt(before); pf.space_after=Pt(after); pf.line_spacing=ls; pf.alignment=WD_ALIGN_PARAGRAPH.RIGHT
    pPr=p._p.get_or_add_pPr()
    if pPr.find(qn('w:widowControl')) is None:
        w=OxmlElement('w:widowControl'); w.set(qn('w:val'),'1'); pPr.append(w)
def add_heading(doc,text,level=1,color="8B1A1A",size=15):
    p=doc.add_paragraph(); set_rtl(p); r=p.add_run(text); set_font(r,"B Nazanin",size,True,color); sp(p,before=14,after=8)
    if level==1:
        pPr=p._p.get_or_add_pPr(); bdr=OxmlElement('w:pBdr'); b=OxmlElement('w:bottom'); b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'4'); b.set(qn('w:space'),'6'); b.set(qn('w:color'),color); bdr.append(b); pPr.append(bdr)
        shd=OxmlElement('w:shd'); shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'FFF2F2'); pPr.append(shd)
    return p
def add_sub(doc,text,size=12.5,color="333333"):
    p=doc.add_paragraph(); set_rtl(p); r=p.add_run(text); set_font(r,"B Nazanin",size,True,color); sp(p,before=8,after=4); return p
def add_p(doc,text,size=11.5,bold=False,color="1A1A1A",after=6,italic=False):
    p=doc.add_paragraph(); set_rtl(p); r=p.add_run(text); set_font(r,"B Nazanin",size,bold,color,italic); sp(p,after=after); return p
def add_bullet(doc,text,size=11):
    p=doc.add_paragraph(); set_rtl(p); p.clear(); r=p.add_run("•  "+text); set_font(r,"B Nazanin",size,color="1A1A1A"); sp(p,after=3,ls=1.35); p.paragraph_format.right_indent=Cm(1); return p
def add_quote(doc,text,source="",size=11,color="004400"):
    p=doc.add_paragraph(); set_rtl(p); pPr=p._p.get_or_add_pPr(); bdr=OxmlElement('w:pBdr'); rt=OxmlElement('w:right'); rt.set(qn('w:val'),'single'); rt.set(qn('w:sz'),'12'); rt.set(qn('w:space'),'8'); rt.set(qn('w:color'),'8B1A1A'); bdr.append(rt); pPr.append(bdr); shd=OxmlElement('w:shd'); shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'F9F5E8'); pPr.append(shd); r=p.add_run(text); set_font(r,"B Nazanin",size,italic=True,color=color)
    if source: r=p.add_run("  — "+source); set_font(r,"B Nazanin",9,color="666666")
    sp(p,before=6,after=8); pf=p.paragraph_format; pf.right_indent=Cm(0.8); pf.left_indent=Cm(0.5); return p

path="/home/user/Hadii/کتاب_ایام_فاطمیه/ایام_فاطمیه_کتاب_جامع_طلاب.docx"
doc=Document(path)
print(f"Before {len(doc.paragraphs)}")

add_heading(doc, "پیوست ۸: وصیت‌نامهٔ نورانی حضرت زهرا سلام‌الله‌علیها", level=1, size=14)
add_p(doc, "وصیت‌نامهٔ حضرت زهرا، سندی تاریخی و عاطفی است که مظلومیت او را ثابت می‌کند. متن کامل وصیت را از بحارالأنوار، ج۴۳، ص۲۱۴ می‌آوریم تا در مجالس بخوانید؛ هر جمله‌اش دریایی از معرفت است:")
add_quote(doc, "بِسْمِ اللهِ الرَّحْمنِ الرَّحیمِ، هذا ما أَوْصَتْ بِهِ فاطِمَةُ بِنْتُ مُحَمَّدٍ، أَوْصَتْ وَ هِیَ تَشْهَدُ أَنْ لا إِلهَ إِلاَّ اللهُ وَ أَنَّ مُحَمَّداً عَبْدُهُ وَ رَسُولُهُ وَ أَنَّ الْجَنَّةَ حَقٌّ وَ النّارَ حَقٌّ وَ أَنَّ السّاعَةَ آتِیَةٌ لا رَیْبَ فیها وَ أَنَّ اللهَ یَبْعَثُ مَنْ فِی الْقُبُورِ. یا عَلِیُّ! أَنَا فاطِمَةُ بِنْتُ مُحَمَّدٍ، زَوَّجَنی اللهُ مِنْکَ لِأَکُونَ لَکَ فِی الدُّنْیا وَ الْآخِرَةِ، أَنْتَ أَوْلی بِی مِنْ غَیْری، حَنِّطْنی وَ غَسِّلْنی وَ کَفِّنْنی بِاللَّیْلِ وَ صَلِّ عَلَیَّ وَ ادْفِنّی بِاللَّیْلِ وَ لا تُعْلِمْ أَحَداً وَ أَسْتَوْدِعُکَ اللهَ وَ أَقْرَأُ عَلی وُلْدی السَّلامَ إِلی یَوْمِ الْقِیامَةِ.", "بحارالأنوار، ج۴۳، ص۲۱۴")
add_p(doc, "ترجمه: به نام خدای بخشندهٔ مهربان، این وصیت فاطمه دختر محمد است؛ وصیت می‌کند در حالی که گواهی می‌دهد خدایی جز الله نیست و محمد بنده و فرستادهٔ اوست و بهشت و دوزخ حق است و قیامت آمدنی است و خدا مردگان را برمی‌انگیزد. ای علی! من فاطمه دختر محمدم، خدا مرا به همسری تو درآورد تا در دنیا و آخرت برای تو باشم، تو از دیگران به من سزاوارتری، مرا شب حنوط کن، غسل بده، کفن کن، بر من نماز بخوان، شب دفن کن، و هیچ‌کس را خبر نکن، تو را به خدا می‌سپارم و بر فرزندانم تا روز قیامت سلام می‌رسانم.")
add_p(doc, "نکتهٔ مهم: حضرت تأکید می‌کند «لا تعلم احداً»؛ به هیچ‌کس خبر نده. چرا؟ چون نمی‌خواست کسانی که او را آزردند، بر پیکرش نماز بخوانند. این وصیت، اعتراض ابدی است. هر کس وصیت فاطمه را می‌خواند، می‌فهمد که فاطمه راضی نبود. طلبه باید این وصیت را با بغض بخواند؛ بغضی که معرفت بیاورد نه کینهٔ کور.")

add_heading(doc, "پیوست ۹: داستان‌های فاطمی برای کودکان و نوجوانان", level=1, size=14)
add_p(doc, "کودکان، آیندهٔ فاطمیه هستند. این سه داستان کوتاه را برای آنان بگویید؛ زبان ساده و شیرین دارد:")
add_sub(doc, "داستان اول: تسبیح هدیهٔ پیامبر")
add_p(doc, "روزی فاطمهٔ کوچک نزد پدر آمد و گفت: بابا جان! دست‌هایم از آسیاب کردن گندم زخم شده است. پیامبر لبخندی زد و فرمود: دخترم! می‌خواهی چیزی به تو یاد بدهم که از کنیز بهتر باشد؟ فاطمه گفت: بله پدر. پیامبر فرمود: بعد از هر نماز بگو: ۳۴ بار الله‌اکبر، ۳۳ بار الحمدلله، ۳۳ بار سبحان‌الله. فاطمه خوشحال شد و این ذکر را همیشه گفت. از آن روز، هر کس این تسبیح را بگوید، به یاد فاطمه می‌افتد. بچه‌ها! شما هم بعد از نماز، تسبیح فاطمه را بگویید تا خدا شما را دوست داشته باشد.")
add_sub(doc, "داستان دوم: نان فاطمه برای فقیر")
add_p(doc, "حسن و حسین مریض شده بودند. فاطمه نذر کرد اگر خوب شوند، سه روز روزه بگیرد. آن‌ها خوب شدند و خانواده روزه گرفتند. هنگام افطار، فقیری در زد و گفت: گرسنه‌ام. فاطمه نان خود را به او داد. روز دوم، یتیمی آمد و نان را به او دادند. روز سوم، اسیری آمد و باز نان را بخشیدند و خود با آب افطار کردند. خدا این ایثار را دوست داشت و سورهٔ «هل اتی» را نازل کرد: «و یطعمون الطعام علی حبه مسکیناً و یتیماً و اسیراً». بچه‌ها! درس این قصه چیست؟ اینکه اول به فکر دیگران باشیم، بعد خودمان؛ مثل فاطمه.")
add_sub(doc, "داستان سوم: چادر فاطمه")
add_p(doc, "فاطمه همیشه چادر مشکی می‌پوشید و حجابش کامل بود. روزی مرد نابینایی به خانه آمد. فاطمه پشت پرده رفت. پیامبر پرسید: او که تو را نمی‌بیند، چرا حجاب کردی؟ فاطمه گفت: پدر جان! او مرا نمی‌بیند، اما من او را می‌بینم و بوی مرا حس می‌کند. پیامبر خوشحال شد و فرمود: آفرین بر تو، تو پارهٔ تن منی. بچه‌ها! حجاب فاطمه فقط چادر نیست؛ حجاب چشم، حجاب زبان و حجاب دل است. دختر فاطمی، در مدرسه و کوچه، با حجابش شناخته می‌شود.")

add_heading(doc, "پیوست ۱۰: فاطمیه و مهدویت؛ پیوند ناگسستنی", level=1, size=14)
add_p(doc, "فاطمیه بدون مهدویت ناقص است. در دعای ندبه می‌خوانیم: «أین الطالب بدم المقتول بکربلاء؟» منتقم فاطمه، مهدی است. امام صادق فرمود: «هنگامی که مهدی ظهور کند، به کعبه تکیه می‌دهد و می‌فرماید: ای مردم! من فرزند فاطمه‌ام، من انتقام مادرم را می‌گیرم». (بحارالأنوار، ج۵۳، ص۱۲)")
add_p(doc, "سرباز مهدی باید فاطمی باشد. کسی که فاطمیه را بزرگ نمی‌دارد، نمی‌تواند منتظر واقعی باشد. در فاطمیه، برای فرج دعا کنیم؛ دعای «اللهم عجل لولیک الفرج» را با اشک فاطمی بخوانیم. هر اشک فاطمی، یک بیعت با مهدی است. به امید روزی که مهدی بیاید و قبر مخفی مادرش را نشان دهد و انتقام سیلی را بگیرد. «متی ترانا و نراک و قد نشرت لواء النصر تری».")
add_quote(doc, "یا مَوْلای یا صاحِبَ الزَّمانِ، الْغَوْثَ الْغَوْثَ الْغَوْثَ، أَدْرِکْنی أَدْرِکْنی أَدْرِکْنی", "دعای استغاثه به امام زمان")

add_heading(doc, "پیوست ۱۱: چک‌لیست برگزاری مجلس فاطمیه (برای هیئت‌ها)", level=1, size=14)
add_p(doc, "این چک‌لیست را پرینت بگیرید و در اختیار هیئت‌ها قرار دهید؛ هر مجلس فاطمیه باید این موارد را داشته باشد:")
add_bullet(doc, "تعیین سخنران و مداح حداقل دو هفته قبل؛ هماهنگی موضوع منبر با کتاب حاضر.")
add_bullet(doc, "سیاه‌پوش کردن مسجد، نصب بنر «یا فاطمة الزهراء» و پرچم مشکی در کوچه‌ها.")
add_bullet(doc, "چاپ بروشور ۴۰ حدیث فاطمی و توزیع بین مردم؛ هدیهٔ فرهنگی ماندگار است.")
add_bullet(doc, "تهیهٔ صوت و تصویر با کیفیت برای انتشار در فضای مجازی؛ هر مجلس باید دیده شود.")
add_bullet(doc, "پذیرایی ساده و با کرامت؛ پرهیز از اسراف؛ نذری پاکیزه و بهداشتی.")
add_bullet(doc, "جداسازی محل خواهران و برادران؛ رعایت حجاب فاطمی در همهٔ بخش‌ها.")
add_bullet(doc, "شروع مجلس با قرآن، ادامه با سخنرانی مستند (۲۰ دقیقه)، روضهٔ مستند (۱۵ دقیقه)، سینه‌زنی (۱۰ دقیقه) و دعای فرج.")
add_bullet(doc, "پایان مجلس با زیارت حضرت زهرا و صلوات هدیه به مادر.")
add_p(doc, "اگر این چک‌لیست رعایت شود، هر مجلس فاطمیه، کلاس معرفت می‌شود نه فقط عزاداری احساسی. هدف، تربیت نسل فاطمی است.")

add_heading(doc, "کتابنامهٔ تکمیلی و منابع دیجیتال", level=1, size=14)
add_p(doc, "علاوه بر منابع اصلی، این منابع تکمیلی نیز توصیه می‌شود:")
add_bullet(doc, "صحیفهٔ فاطمیه: مجموعهٔ ادعیه و زیارات حضرت زهرا با ترجمهٔ روان - نشر مسجد جمکران.")
add_bullet(doc, "سوگنامهٔ فاطمه: مقتل مستند شهادت با ذکر منابع سنی و شیعه - سید حسن افتخارزاده.")
add_bullet(doc, "فاطمیه در آینهٔ اسناد: پژوهشی تطبیقی از منابع دست اول - پژوهشگاه باقرالعلوم.")
add_bullet(doc, "مستند «بانوی بی‌نشان»: روایت تصویری تاریخ فاطمیه - شبکهٔ افق (قابل دانلود از سایت حوزه).")
add_bullet(doc, "نرم‌افزار «کوثر»: ۴۰ حدیث، ۴۰ داستان، ۴۰ شعر فاطمی برای موبایل - تولید مرکز نور قم.")

add_heading(doc, "سخن آخر: امتحان بزرگ", level=1, size=15)
add_p(doc, "طلبهٔ عزیز! کتاب را خواندی، اما امتحان اصلی در عمل است. آیا در فاطمیهٔ امسال، یک مجلس خواهی گرفت؟ آیا به یک جوان، حقیقت فاطمیه را خواهی گفت؟ آیا حجاب فاطمی را در خانواده‌ات ترویج خواهی کرد؟ فاطمه، تو را می‌بیند. او مادر است و نگران فرزندانش است. اگر برای او قدم برداری، او برای تو قدم برمی‌دارد.")
add_p(doc, "ما در حوزهٔ علمیه، پیمان بسته‌ایم: تا جان داریم، نام فاطمه را زنده نگه داریم. تو نیز بپیوند. این کتاب را به پنج طلبهٔ دیگر هدیه بده، در مسجدت تدریس کن، و در فضای مجازی منتشر کن. نشر معارف فاطمی، صدقهٔ جاریه است.")
add_quote(doc, "یا فاطِمَةُ اشْفَعی لَنا فِی الْجَنَّةِ فَإِنَّ لَکِ عِنْدَ اللهِ شَأْناً مِنَ الشَّأْنِ", "زیارت حضرت زهرا سلام‌الله‌علیها")
add_p(doc, "خدایا! به حق پهلوی شکستهٔ فاطمه، به حق اشک حسن و حسین، به حق صبر علی، فرج مهدی را برسان، حوزه‌ها را حفظ کن، رهبر انقلاب را سلامت بدار، و ما را فاطمی بمیران. آمین یا رب العالمین.", size=12, color="8B1A1A", italic=True)
add_p(doc, "پایان نگارش: قم المقدسه - جوار حرم حضرت معصومه سلام‌الله‌علیها - شب شهادت حضرت زهرا سلام‌الله‌علیها - ۱۴۰۳ هـ.ش", size=10, color="666666", italic=True)

doc.save(path)
print(f"After expand3 paras: {len(doc.paragraphs)}")
total_text = "\n".join([p.text for p in doc.paragraphs])
print(f"Word count: {len(total_text.split())}")
print(f"Chars: {len(total_text)}")
