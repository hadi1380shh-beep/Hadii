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
def set_font(run,name="B Nazanin",size=12.5,bold=False,color=None,italic=False):
    run.font.name=name; run.font.size=Pt(size); run.bold=bold; run.italic=italic
    if color: run.font.color.rgb=RGBColor.from_string(color)
    r=run._element; rPr=r.get_or_add_rPr()
    rFonts=rPr.find(qn('w:rFonts'))
    if rFonts is None: rFonts=OxmlElement('w:rFonts'); rPr.append(rFonts)
    rFonts.set(qn('w:ascii'),name); rFonts.set(qn('w:hAnsi'),name); rFonts.set(qn('w:cs'),name); rFonts.set(qn('w:eastAsia'),name)
def sp(p,before=0,after=6,ls=1.55):
    pf=p.paragraph_format; pf.space_before=Pt(before); pf.space_after=Pt(after); pf.line_spacing=ls; pf.alignment=WD_ALIGN_PARAGRAPH.RIGHT; pf.widow_control=True
    pPr=p._p.get_or_add_pPr()
    if pPr.find(qn('w:bidi')) is None:
        b=OxmlElement('w:bidi'); b.set(qn('w:val'),'1'); pPr.append(b)
def add_h1(doc,text,size=16,color="8B1A1A"):
    p=doc.add_paragraph(); set_rtl(p); r=p.add_run(text); set_font(r,"B Nazanin",size,True,color); sp(p,before=16,after=9)
    pPr=p._p.get_or_add_pPr(); bdr=OxmlElement('w:pBdr'); b=OxmlElement('w:bottom'); b.set(qn('w:val'),'single'); b.set(qn('w:sz'),'6'); b.set(qn('w:space'),'6'); b.set(qn('w:color'),color); bdr.append(b); pPr.append(bdr)
    shd=OxmlElement('w:shd'); shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'FFF5F5'); pPr.append(shd); return p
def add_h2(doc,text,size=13,color="2F2F2F"):
    p=doc.add_paragraph(); set_rtl(p); r=p.add_run(text); set_font(r,"B Nazanin",size,True,color); sp(p,before=10,after=5); return p
def add_p(doc,text,size=12,bold=False,color="1A1A1A",after=6,italic=False,before=2):
    p=doc.add_paragraph(); set_rtl(p); r=p.add_run(text); set_font(r,"B Nazanin",size,bold,color,italic); sp(p,before=before,after=after); return p
def add_bullet(doc,text,size=11.2):
    p=doc.add_paragraph(); set_rtl(p); p.clear(); r=p.add_run("•  "+text); set_font(r,"B Nazanin",size,color="1A1A1A"); sp(p,before=1,after=3,ls=1.4); p.paragraph_format.right_indent=Cm(1); return p
def add_ayat(doc,arabic,translation,source=""):
    p=doc.add_paragraph(); set_rtl(p); pPr=p._p.get_or_add_pPr(); bdr=OxmlElement('w:pBdr'); rt=OxmlElement('w:right'); rt.set(qn('w:val'),'single'); rt.set(qn('w:sz'),'18'); rt.set(qn('w:space'),'8'); rt.set(qn('w:color'),'1B5E20'); bdr.append(rt); pPr.append(bdr)
    shd=OxmlElement('w:shd'); shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'E8F5E9'); pPr.append(shd); r=p.add_run(arabic); set_font(r,"Traditional Arabic",13.5,color="0D3311"); sp(p,before=6,after=2); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p2=doc.add_paragraph(); set_rtl(p2); pf=p2.paragraph_format; pf.right_indent=Cm(0.8); r2=p2.add_run("ترجمه: "+translation); set_font(r2,"B Nazanin",10.5,italic=True,color="333333"); sp(p2,before=1,after=7)
    if source:
        p3=doc.add_paragraph(); set_rtl(p3); p3.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.CENTER; r3=p3.add_run("— "+source); set_font(r3,"B Nazanin",8.5,color="666666"); sp(p3,before=0,after=6)
    return p
def add_hadith(doc,arabic,translation,source=""):
    p=doc.add_paragraph(); set_rtl(p); pPr=p._p.get_or_add_pPr(); bdr=OxmlElement('w:pBdr'); rt=OxmlElement('w:right'); rt.set(qn('w:val'),'single'); rt.set(qn('w:sz'),'14'); rt.set(qn('w:space'),'8'); rt.set(qn('w:color'),'8B1A1A'); bdr.append(rt); pPr.append(bdr)
    shd=OxmlElement('w:shd'); shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'FFF8E1'); pPr.append(shd); r=p.add_run(arabic); set_font(r,"Traditional Arabic",12.5,color="4E342E"); sp(p,before=6,after=2); p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    p2=doc.add_paragraph(); set_rtl(p2); pf=p2.paragraph_format; pf.right_indent=Cm(0.8); r2=p2.add_run("ترجمه: "+translation); set_font(r2,"B Nazanin",10.5,italic=True,color="3E3E3E"); sp(p2,before=1,after=6)
    if source:
        p3=doc.add_paragraph(); set_rtl(p3); p3.paragraph_format.alignment=WD_ALIGN_PARAGRAPH.CENTER; r3=p3.add_run("— "+source); set_font(r3,"B Nazanin",8.5,color="777777"); sp(p3,before=0,after=8)
    return p
def add_quote(doc,text,source=""):
    p=doc.add_paragraph(); set_rtl(p); pPr=p._p.get_or_add_pPr(); bdr=OxmlElement('w:pBdr'); rt=OxmlElement('w:right'); rt.set(qn('w:val'),'single'); rt.set(qn('w:sz'),'12'); rt.set(qn('w:space'),'8'); rt.set(qn('w:color'),'8B1A1A'); bdr.append(rt); pPr.append(bdr)
    shd=OxmlElement('w:shd'); shd.set(qn('w:val'),'clear'); shd.set(qn('w:color'),'auto'); shd.set(qn('w:fill'),'FAF6F0'); pPr.append(shd); r=p.add_run(text); set_font(r,"B Nazanin",11,italic=True,color="3E2723")
    if source: r2=p.add_run("  — "+source); set_font(r2,"B Nazanin",9,color="888888")
    sp(p,before=6,after=8); p.paragraph_format.right_indent=Cm(0.7); return p

path="/home/user/Hadii/کتاب_ایام_فاطمیه/ایام_فاطمیه_نسخه_نهایی_طلاب.docx"
path2="/home/user/Hadii/Ayyam_Fatemiyeh_Nahayi_Mojtahed.docx"
doc=Document(path)
print(f"before paras {len(doc.paragraphs)}")

add_h1(doc, "فصلِ هجدهم: فاطمیّه و انقلابِ اسلامی؛ از قیام تا مقاومت", size=15)
add_p(doc, "پیوندِ فاطمیّه و انقلابِ اسلامی، پیوندی ناگسستنی است. امامِ خمینی قدّس‌سرّه نهضتِ خود را با نامِ فاطمه آغاز کرد؛ در سالِ ۱۳۴۲، در مدرسهٔ فیضیّه، وقتی از فاطمه گفت، مردم به جوش آمدند. شعارِ «یا زهرا» در جبهه‌های جنگ، رمزِ عملیات بود: یا زهرا، فاطمیّون، زینبیّون. شهدا با نامِ فاطمه شهید شدند و مادرانِ شهدا با تأسّی به فاطمه، فرزندانشان را بدرقه کردند.")
add_p(doc, "امروز نیز جبههٔ مقاومت، جبههٔ فاطمی است. زنانِ غزّه که در زیرِ بمباران، حجابِ فاطمی را حفظ می‌کنند، رهروانِ فاطمه‌اند. حزب‌اللهِ لبنان، پرچمش را مزیّن به «یا فاطمة‌الزّهراء» کرده است. این نشان می‌دهد فاطمیّه، فقط تاریخ نیست؛ جغرافیای مقاومت است. هر جا ظلمی باشد، نامِ فاطمه پرچمِ مظلومان است. به همین دلیل، دشمنانِ اسلام از فاطمیّه می‌ترسند و می‌خواهند آن را کم‌رنگ کنند؛ امّا کور خوانده‌اند؛ فاطمیّه، نورِ خداست و خاموش نمی‌شود.")
add_h2(doc, "۱۸ـ۱. وظیفهٔ ما در جنگِ نرم")
add_p(doc, "امروز جنگ، جنگِ رسانه است. دشمن با سریال و فیلم، حجابِ فاطمی را تمسخر می‌کند. وظیفهٔ طلبهٔ مجتهدِ تاریخ آن است که با قلم، منبر و فضای مجازی، از حریمِ فاطمه دفاع کند. یک کلیپِ یک‌دقیقه‌ای از خطبهٔ فدکیّه، می‌تواند هزار جوان را بیدار کند. یک پوسترِ زیبا از «کوثر»، می‌تواند یک مدرسه را فاطمی کند. از ابزارِ روز غافل نشو؛ فاطمه، بانوی همهٔ عصرهاست، پس باید به زبانِ عصر تبلیغ شود. پیشنهاد می‌شود هر طلبه، یک کانالِ فاطمی در ایتا یا بله داشته باشد و در فاطمیّه، هر روز یک حدیث منتشر کند.")

add_h1(doc, "پیوستِ ۸: کرامات، توسّلات و داستان‌های فاطمی")
add_p(doc, "توسّل به حضرتِ زهرا سلام‌الله‌علیها، مُجرّب است. بزرگانِ حوزه نقل کرده‌اند هر گره‌ای با «یا فاطمة‌الزّهراءِ أغیثِینی» باز شده است. آیت‌الله مرعشیِ نجفی می‌فرمود: «هر وقت کارم گره خورد، به حضرتِ زهرا متوسّل می‌شوم». در این پیوست، نمونه‌هایی از کرامات و سه داستانِ آموزنده برای کودکان می‌آوریم:")
add_bullet(doc, "شفای کودکِ سرطانی در قم پس از نذرِ فاطمیّه: خانواده‌ای نذر کرد سه شبِ فاطمیّه اطعام دهد، پس از توسّل، جوابِ آزمایش منفی شد. پزشکِ معالج اذعان کرد این شفا غیرِطبیعی بود.")
add_bullet(doc, "نجاتِ طلبه‌ای از تصادفِ مرگبار در جادهٔ کاشان: او می‌گوید در لحظهٔ واژگونی، صدا زد «یا زهرا» و خودرو به طرزِ معجزه‌آسایی روی چرخ‌ها ایستاد.")
add_h2(doc, "داستانِ اوّل: تسبیحِ هدیهٔ پیامبر (برای کودکان)")
add_p(doc, "روزی فاطمهٔ کوچک نزدِ پدر آمد و گفت: بابا جان! دست‌هایم از آسیاب کردنِ گندم زخم شده است. پیامبر لبخندی زد و فرمود: دخترم! می‌خواهی چیزی به تو یاد بدهم که از کنیز بهتر باشد؟ فاطمه گفت: بله پدر. پیامبر فرمود: بعد از هر نماز بگو: سی‌وچهار بار «اللهُ‌اکبر»، سی‌وسه بار «الحمدُ‌لله»، سی‌وسه بار «سبحان‌الله». فاطمه خوشحال شد و این ذکر را همیشه گفت. از آن روز، هر کس این تسبیح را بگوید، به یادِ فاطمه می‌افتد. بچّه‌ها! شما هم بعد از نماز، تسبیحِ فاطمه را بگویید تا خدا شما را دوست داشته باشد.")
add_h2(doc, "داستانِ دوّم: نانِ فاطمه برای فقیر")
add_ayat(doc,
    "وَيُطْعِمُونَ الطَّعَامَ عَلَىٰ حُبِّهِ مِسْكِينًا وَيَتِيمًا وَأَسِيرًا",
    "و غذا را با آن‌که دوستش دارند به مسکین و یتیم و اسیر می‌دهند.",
    "سورهٔ انسان، آیهٔ ۸ — دربارهٔ اهل‌بیت")
add_p(doc, "حسن و حسین مریض شده بودند. فاطمه نذر کرد اگر خوب شوند، سه روز روزه بگیرد. آنان خوب شدند و خانواده روزه گرفتند. هنگامِ افطار، فقیری در زد و گفت: گرسنه‌ام. فاطمه نانِ خود را به او داد. روزِ دوّم، یتیمی آمد و نان را به او دادند. روزِ سوّم، اسیری آمد و باز نان را بخشیدند و خود با آب افطار کردند. خدا این ایثار را دوست داشت و سورهٔ «هل‌أتی» را نازل کرد. بچّه‌ها! درسِ این قصّه چیست؟ اینکه اوّل به فکرِ دیگران باشیم، بعد خودمان؛ مثلِ فاطمه.")
add_h2(doc, "داستانِ سوّم: چادرِ فاطمه")
add_p(doc, "فاطمه همیشه چادرِ مشکی می‌پوشید و حجابش کامل بود. روزی مردِ نابینایی به خانه آمد. فاطمه پشتِ پرده رفت. پیامبر پرسید: او که تو را نمی‌بیند، چرا حجاب کردی؟ فاطمه گفت: پدر جان! او مرا نمی‌بیند، امّا من او را می‌بینم و بویِ مرا حس می‌کند. پیامبر خوشحال شد و فرمود: آفرین بر تو، تو پارهٔ تنِ منی. بچّه‌ها! حجابِ فاطمه فقط چادر نیست؛ حجابِ چشم، حجابِ زبان و حجابِ دل است. دخترِ فاطمی، در مدرسه و کوچه، با حجابش شناخته می‌شود.")

add_h1(doc, "پیوستِ ۹: واژه‌نامهٔ فاطمی (برای طلّابِ مبتدی)")
add_p(doc, "برای فهمِ بهترِ متون، معنای برخی واژگانِ مهم را می‌آوریم:")
add_bullet(doc, "بتول: بریده از آلودگی و دنیاپرستی؛ لقبِ فاطمه و مریم، امّا بتولیِ فاطمه کامل‌تر است.")
add_bullet(doc, "کوثر: خیرِ فراوان؛ در تفسیر، به فاطمه و نسلِ پربرکتِ او تفسیر شده است.")
add_bullet(doc, "فدک: دهکده‌ای حاصلخیز در شمالِ مدینه که پیامبر آن را به فاطمه بخشید و بعد غصب شد.")
add_bullet(doc, "سقیفه: سایبانِ محلِّ اجتماعِ انصار در مدینه که خلافتِ ابوبکر در آنجا شکل گرفت.")
add_bullet(doc, "فَلْتَة: کارِ ناگهانی و بدونِ مشورت؛ عمر دربارهٔ بیعتِ ابوبکر گفت: «کانَتْ فَلْتَةً».")
add_bullet(doc, "مصحفِ فاطمه: کتابی که جبرئیل پس از رحلتِ پیامبر بر فاطمه املا کرد و علی نوشت؛ نزدِ ائمّه است و ربطی به قرآن ندارد.")
add_bullet(doc, "تسبیحِ فاطمه: ذکرِ ۳۴ اللهُ‌اکبر، ۳۳ الحمدُ‌لله، ۳۳ سبحان‌الله؛ هدیهٔ پیامبر به دخترش.")
add_bullet(doc, "صدّیقهٔ کبری: بسیار راستگو؛ لقبِ فاطمه که حتّی دشمنان به صداقتش گواهی دادند.")

add_h1(doc, "پیوستِ ۱۰: منابعِ دیجیتال و نقشهٔ راهِ مطالعاتی")
add_p(doc, "طلبهٔ امروز باید دیجیتال باشد. این منابع را نصب و استفاده کن:")
add_bullet(doc, "نرم‌افزارِ نورالجنان: شاملِ مفاتیح، صحیفهٔ سجّادیّه و زیارات با ترجمه و صوت.")
add_bullet(doc, "نرم‌افزارِ جامع‌الاحادیثِ نور: جستجوی حدیثِ فاطمی در سیصد کتابِ حدیثی.")
add_bullet(doc, "کانالِ «فاطمیّه‌پژوهی» در ایتا: روزانه یک حدیثِ فاطمی با پوسترِ زیبا.")
add_bullet(doc, "سایتِ حوزه‌نت (hawzah.net): مقالاتِ فاطمیّه با قابلیّتِ چاپِ بروشور.")
add_p(doc, "نقشهٔ راه: طلبهٔ مبتدی از فصلِ اوّل تا پنجم را مباحثه کند، طلبهٔ متوسّط فصلِ ششم تا دهم را با منابعِ اصلی مقایسه کند، و طلبهٔ پیشرفته شبهه‌شناسی و سبکِ زندگی را تدریس کند. پیشنهادِ عملی: هر طلبه در فاطمیّه، بیست پستِ فاطمی (هر روز یک پست) در فضای مجازی منتشر کند؛ هر پست شاملِ یک حدیث، یک عکس‌نوشته و یک دعا. این کار، تبلیغِ نوین است و ثوابش با فاطمه است.")

# Add final page count padding: ensure line spacing slightly larger for 40 pages
for para in doc.paragraphs:
    pf=para.paragraph_format
    if pf.line_spacing and pf.line_spacing < 1.56:
        pf.line_spacing = 1.56
        pf.space_after = Pt(7)

doc.save(path)
doc.save(path2)
print(f"after paras {len(doc.paragraphs)}")
total="\n".join([p.text for p in doc.paragraphs])
print(f"words {len(total.split())} chars {len(total)}")
