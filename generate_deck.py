"""
Eco-Drain WEFE Nexus - 5-minute executive pitch generator (3 slides, 16:9).

Features
- Real RTL handling for Arabic (paragraph rtl="1", complex-script font, mirrored layout).
- Rounded cards, rounded pictures, KPI chips, funding-fit badges.
- Figures from /images embedded in rounded frames.
- Slide transitions (fade) + staggered entrance animations (fade / wipe).
- Auto-fit estimator so text never overflows its card.
- Speaker notes with a 5-minute timing script.
"""
import math
import os
import shutil

from PIL import Image
from pptx import Presentation
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls, qn
from pptx.util import Inches, Pt

BASE = os.path.dirname(os.path.abspath(__file__))
IMG = os.path.join(BASE, "images")

# ----------------------------------------------------------------------------
# Palette
# ----------------------------------------------------------------------------
BG_TOP = RGBColor(13, 29, 54)
BG_BOTTOM = RGBColor(6, 14, 28)
BG_CARD = RGBColor(15, 34, 64)
BG_CARD_ALT = RGBColor(24, 49, 88)
TEAL = RGBColor(21, 181, 164)
GOLD = RGBColor(240, 200, 96)
GOLD_DEEP = RGBColor(212, 168, 67)
WHITE = RGBColor(248, 250, 252)
SOFT = RGBColor(226, 232, 240)
GRAY = RGBColor(148, 163, 184)
GREEN = RGBColor(16, 185, 129)
RED = RGBColor(248, 113, 113)
BLUE = RGBColor(96, 165, 250)
NAVY_TEXT = RGBColor(10, 22, 40)

FONT = "Segoe UI"
SW, SH = 13.333, 7.5
MARGIN = 0.8
MEDIA_W = 3.9
GAP = 0.12
CONTENT_L = MARGIN + MEDIA_W + GAP
CONTENT_W = (SW - MARGIN) - CONTENT_L
BODY_TOP = 1.7

# ----------------------------------------------------------------------------
# Content
# ----------------------------------------------------------------------------
CONTENT = {
    "en": {
        "pill": "SLIDE {n} OF 3  |  5-MINUTE EXECUTIVE PITCH",
        "footer": "Eco-Drain  |  Group 10  |  EPADP Technical Office, Lower Egypt",
        "s1": {
            "title": "Eco-Drain: Invisible, Tamper-Proof Subsurface Drainage",
            "sub": "Climate-resilient retrofit of 100,000 feddans in Egypt's Nile Delta",
            "img": "fig1_subsurface_drainage.png",
            "cap": "Subsurface drainage retrofit concept",
            "tiles": [
                ("560 m³", "water per capita per year, below the 1,000 m³ poverty line"),
                ("2.3M", "feddans of drainage networks past their design life"),
                ("25–40%", "of drainage failures caused by tampering"),
            ],
            "a_title": "Problem addressed and invisible solution",
            "a": [
                ("Ageing networks:", "Delta drainage is past its 20–25 year design life, causing waterlogging, salinity and 10–30% yield loss."),
                ("Hidden failures:", "buried pipes cannot be inspected, so faults go unseen for years."),
                ("Tampering:", "farmers block surface manholes to pool rice water, causing 25–40% of failures."),
                ("Eco-Drain:", "chambers buried 30–50 cm deep, secured by RFID/GPS, with shock sensors (>2.5g) sending encrypted NB-IoT alerts."),
            ],
            "b_title": "Lead institution, partners and governance",
            "b": [
                ("Lead:", "Ministry of Water Resources and Irrigation (MWRI) through EPADP, Technical Office for Lower Egypt."),
                ("Partners:", "DRI (verification, MRV), MALR (crop advisory), NREA and EgyptERA (50 MW net metering), Water User Associations."),
                ("Governance:", "inter-ministerial Steering Committee, dedicated PIU in EPADP, local platforms in Kafr El-Sheikh and Dakahlia."),
                ("Alignment:", "Irrigation 2.0, Egypt Vision 2030, National Climate Strategy 2050."),
            ],
        },
        "s2": {
            "title": "Integrated WEFE Interventions and Bankability",
            "sub": "Four closed-loop interventions, a formal screening score and a resilience scorecard",
            "img": "fig2_wefe_nexus.png",
            "cap": "Water–Energy–Food–Ecosystems nexus",
            "score_title": "WEFE screening and resilience",
            "score_big": "92/100",
            "score_rating": "Resilience: HIGH",
            "score_dims": "Water 10 · Energy 8 · Food 10 · Ecosystems 8 · Digital 9",
            "cards": [
                ("WATER", "150M m³/yr saved", "HDPE and geotextile retrofit with smart gate valves using capillary rise: 25–30% less irrigation water.", BLUE),
                ("ENERGY", "50 MW solar · 90 GWh/yr", "Distributed PV at 25 pump stations cuts pumping bills by 40%; surplus exported via net metering.", GOLD),
                ("FOOD", "15–25% higher yields", "Wheat, rice and maize restored for 200,000 farmers, with automatic flushing when salinity rises.", GREEN),
                ("ECOSYSTEMS", "100k tCO₂/yr · 4,760 feddans", "20 constructed wetlands treat 200M m³/yr and protect Delta lakes (Verra blue carbon).", TEAL),
            ],
            "bank_title": "How WEFE integration improves bankability",
            "bank": [
                ("Self-funded OPEX:", "solar net metering pays pump power and protects the DSCR (1.35x)."),
                ("Avoided cost and revenue:", "150M m³/yr is worth about $52.5M; smart-valve payback 0.48 years; $8M energy, $3M export, $5M water, $0.5M carbon per year."),
                ("No displacement:", "in-situ retrofit with zero resettlement or community conflict."),
            ],
        },
        "s3": {
            "title": "Capital Stack, Readiness and Expected Results",
            "sub": "USD 420M blended finance, a ready 10-feddan pilot and a three-stage scaling pathway",
            "img": "fig5_capital_stack.png",
            "cap": "Blended-finance capital stack",
            "fit_title": "Funding call fit",
            "fit": [
                ("14/14", "Green Climate Fund", "Transformational adaptation and digital MRV. Decision: GO"),
                ("13/14", "Adaptation Fund", "Ideal for the pilot component. Decision: GO"),
                ("High", "PRIMA and EU Global Gateway", "Water-agri innovation and Euro-Med cooperation"),
            ],
            "bud_title": "Budget: USD 420M ($4,200 per feddan)",
            "bud": [
                ("15% grants ($63M):", "GCF, AfDB, EU for pilot, digital twin and capacity."),
                ("20% government ($84M):", "in-kind land, pump stations, EPADP staff."),
                ("40% concessional debt ($168M):", "World Bank, AfDB, 20–25 year tenor."),
                ("25% PPP / commercial ($105M):", "solar developers under PPA. EIRR 21.4%."),
            ],
            "rdy_title": "Current readiness status",
            "rdy": [
                ("Track record:", "EPADP's 50 years, 6M feddans installed, TRL 8–9 technologies."),
                ("Pilot ready:", "10-feddan unit on an isolated collector to validate before scale-up."),
                ("Next step:", "$1.5M PPF grant request to GCF/AfDB for Kafr El-Sheikh feasibility."),
            ],
            "res_title": "Expected results (KPIs) and scaling pathway",
            "kpis": [
                ("150M m³", "water saved / yr"),
                ("15–25%", "yield increase"),
                ("90 GWh", "clean energy / yr"),
                ("80%", "less tampering"),
            ],
            "phases": ["1 · Pilot 10 feddans (Yr 1)", "2 · Core 100k feddans (Yr 2–5)", "3 · Scale 4.3M feddans"],
        },
        "notes": [
            "00:00-01:40  Problem: 560 m3 per capita, 2.3M feddans past design life, 25-40% of failures from tampering. Solution: invisible buried chambers with RFID/GPS and shock alerts. Lead: MWRI/EPADP with DRI, MALR, NREA, EgyptERA and WUAs.",
            "01:40-03:30  Four WEFE interventions: 150M m3 water saved, 50 MW solar, +15-25% yields, 20 wetlands. Composite WEFE score 92/100. Bankability: self-funded OPEX, avoided cost, diversified revenue, no displacement.",
            "03:30-05:00  USD 420M blended finance (15/20/40/25). Funding fit: GCF 14/14, Adaptation Fund 13/14. Ready: 10-feddan pilot, TRL 8-9, PPF request. Scaling: 10 feddans, 100k feddans, 4.3M feddans in the Delta, then Iraq, Jordan and Pakistan.",
        ],
    },
    "ar": {
        "pill": "الشريحة {n} من 3  |  عرض تنفيذي لمدة 5 دقائق",
        "footer": "إيكو-درين  |  المجموعة العاشرة  |  المكتب الفني لهيئة الصرف بالوجه البحري",
        "s1": {
            "title": "إيكو-درين: صرف مغطى ذكي غير مرئي ومقاوم للتلاعب",
            "sub": "تجديد مرن مناخياً لـ 100,000 فدان في دلتا النيل بمصر",
            "img": "fig1_subsurface_drainage.png",
            "cap": "تصور تجديد شبكات الصرف المغطى",
            "tiles": [
                ("560 م³", "نصيب الفرد من المياه سنوياً، أقل من خط الفقر المائي"),
                ("2.3 مليون", "فدان من شبكات الصرف تجاوزت عمرها التصميمي"),
                ("25–40%", "من أعطال الصرف سببها التلاعب"),
            ],
            "a_title": "المشكلة والحل غير المرئي",
            "a": [
                ("تقادم الشبكات:", "الصرف بالدلتا تجاوز عمره التصميمي (20–25 سنة) مسبباً التغدق والتملح وفقد 10–30% من المحاصيل."),
                ("أعطال خفية:", "المواسير مدفونة ولا يمكن فحصها، فتبقى الأعطال مجهولة لسنوات."),
                ("التلاعب:", "سد المزارعين للمناهل السطحية لحبس مياه الأرز يتسبب في 25–40% من الأعطال."),
                ("حل إيكو-درين:", "غرف مدفونة على عمق 30–50 سم، مؤمنة بـ RFID وGPS، وحساسات اهتزاز (>2.5g) ترسل إنذاراً مشفراً عبر NB-IoT."),
            ],
            "b_title": "الجهة القائدة والشركاء والحوكمة",
            "b": [
                ("الجهة القائدة:", "وزارة الموارد المائية والري (MWRI) عبر هيئة الصرف (EPADP) — المكتب الفني للوجه البحري."),
                ("الشركاء:", "معهد بحوث الصرف (التحقق وMRV)، وزارة الزراعة، هيئة الطاقة المتجددة وجهاز تنظيم الكهرباء (50 MW)، روابط مستخدمي المياه."),
                ("الحوكمة:", "لجنة توجيهية بين الوزارات، وحدة تنفيذ مخصصة (PIU) بهيئة الصرف، ومنصات محلية في كفر الشيخ والدقهلية."),
                ("التوافق:", "منظومة الري 2.0، ورؤية مصر 2030، واستراتيجية المناخ 2050."),
            ],
        },
        "s2": {
            "title": "تدخلات WEFE المتكاملة والجدوى البنكية",
            "sub": "أربعة تدخلات مترابطة ودرجة فرز رسمية وبطاقة قياس للمرونة",
            "img": "fig2_wefe_nexus.png",
            "cap": "ترابط المياه والطاقة والغذاء والنظم البيئية",
            "score_title": "فرز WEFE ومرونة المشروع",
            "score_big": "92/100",
            "score_rating": "المرونة: عالية",
            "score_dims": "المياه 10 · الطاقة 8 · الغذاء 10 · النظم البيئية 8 · الرقمنة 9",
            "cards": [
                ("المياه", "توفير 150 مليون م³ سنوياً", "مواسير HDPE وفلاتر جيوتكستيل مع صمامات ذكية تستفيد من الصعود الشعري: توفير 25–30% من مياه الري.", BLUE),
                ("الطاقة", "50 MW شمسية · 90 GWh سنوياً", "طاقة موزعة على 25 محطة رفع تخفض فاتورة الضخ 40% وتصدّر الفائض بنظام صافي القياس.", GOLD),
                ("الغذاء", "زيادة الإنتاجية 15–25%", "استعادة غلة القمح والأرز والذرة لـ 200,000 مزارع مع غسيل آلي عند ارتفاع الملوحة.", GREEN),
                ("النظم البيئية", "100 ألف طن CO₂ · 4,760 فدان", "20 أرضاً رطبة تعالج 200 مليون م³ سنوياً وتحمي بحيرات الدلتا (كربون أزرق Verra).", TEAL),
            ],
            "bank_title": "كيف يعزز تكامل WEFE الجدوى البنكية",
            "bank": [
                ("تشغيل ذاتي التمويل:", "الطاقة الشمسية تسدد كهرباء الطلمبات وتحمي نسبة خدمة الدين 1.35x."),
                ("تكلفة متجنبة وإيرادات:", "توفير 150 مليون م³ ≈ 52.5 مليون دولار سنوياً، واسترداد خلال 0.48 سنة، وإيرادات طاقة وتصدير ومياه وكربون."),
                ("بلا تهجير:", "تطوير موضعي دون إعادة توطين أو نزاعات مجتمعية."),
            ],
        },
        "s3": {
            "title": "هيكل التمويل والجاهزية والنتائج المتوقعة",
            "sub": "تمويل مختلط بقيمة 420 مليون دولار وحقل تجريبي جاهز ومسار توسع من ثلاث مراحل",
            "img": "fig5_capital_stack.png",
            "cap": "هيكل رأس المال المختلط",
            "fit_title": "ملاءمة جهات التمويل",
            "fit": [
                ("14/14", "صندوق المناخ الأخضر (GCF)", "تكيف تحويلي ورقمنة MRV. القرار: المضي"),
                ("13/14", "صندوق التكيف", "مثالي لتمويل الحقل التجريبي. القرار: المضي"),
                ("مرتفعة", "PRIMA وبوابة الاتحاد الأوروبي", "ابتكار مائي زراعي وتعاون متوسطي"),
            ],
            "bud_title": "الميزانية: 420 مليون دولار (4,200 دولار للفدان)",
            "bud": [
                ("15% منح (63 مليون دولار):", "GCF وAfDB والاتحاد الأوروبي للتجريبي والتوأم الرقمي."),
                ("20% مساهمة حكومية (84 مليون دولار):", "أراضٍ ومحطات رفع وكوادر هيئة الصرف."),
                ("40% قروض ميسرة (168 مليون دولار):", "البنك الدولي وAfDB لمدة 20–25 سنة."),
                ("25% شراكة خاصة (105 مليون دولار):", "مطورو طاقة شمسية بعقود PPA. EIRR 21.4%."),
            ],
            "rdy_title": "حالة الجاهزية الحالية",
            "rdy": [
                ("سجل التنفيذ:", "50 عاماً لهيئة الصرف و6 ملايين فدان منفذة وتقنيات TRL 8–9."),
                ("حقل تجريبي جاهز:", "وحدة 10 أفدنة على مجمع معزول للتحقق قبل التوسع."),
                ("الخطوة التالية:", "طلب منحة تحضيرية 1.5 مليون دولار من GCF/AfDB لدراسة كفر الشيخ."),
            ],
            "res_title": "النتائج المتوقعة (KPIs) ومسار التوسع",
            "kpis": [
                ("150 مليون م³", "وفر مائي سنوياً"),
                ("15–25%", "زيادة الإنتاجية"),
                ("90 GWh", "طاقة نظيفة سنوياً"),
                ("80%", "انخفاض التعديات"),
            ],
            "phases": ["1 · تجريبي 10 أفدنة (سنة 1)", "2 · ريادي 100 ألف فدان (2–5)", "3 · تعميم 4.3 مليون فدان"],
        },
        "notes": [
            "00:00-01:40  المشكلة: نصيب الفرد 560 م³، و2.3 مليون فدان تجاوزت عمرها التصميمي، و25-40% من الأعطال بسبب التلاعب. الحل: غرف مدفونة بـ RFID وGPS وإنذار اهتزاز. القيادة: وزارة الري وهيئة الصرف مع الشركاء.",
            "01:40-03:30  أربعة تدخلات WEFE: توفير 150 مليون م³، 50 MW شمسية، زيادة الإنتاجية 15-25%، و20 أرضاً رطبة. التقييم المركب 92/100. الجدوى البنكية: تشغيل ذاتي التمويل وتكلفة متجنبة وتنوع الإيرادات وعدم التهجير.",
            "03:30-05:00  تمويل مختلط 420 مليون دولار (15/20/40/25). ملاءمة GCF 14/14 وصندوق التكيف 13/14. الجاهزية: حقل 10 أفدنة وTRL 8-9 وطلب PPF. التوسع: 10 أفدنة ثم 100 ألف فدان ثم 4.3 مليون فدان بالدلتا ثم العراق والأردن وباكستان.",
        ],
    },
}


# ----------------------------------------------------------------------------
# Builder
# ----------------------------------------------------------------------------
def build_deck(output_path, lang="en"):
    is_ar = lang == "ar"
    T = CONTENT[lang]
    warnings = []

    prs = Presentation()
    prs.slide_width = Inches(SW)
    prs.slide_height = Inches(SH)
    blank = prs.slide_layouts[6]

    # ---- geometry / text helpers -------------------------------------------
    def mx(left, width):
        """Mirror an LTR x position for Arabic (RTL) layouts."""
        return SW - left - width if is_ar else left

    def style_run(run, size, color, bold=False):
        f = run.font
        f.name = FONT
        f.size = Pt(size)
        f.bold = bold
        f.color.rgb = color
        rpr = run._r.get_or_add_rPr()
        rpr.set("lang", "ar-SA" if is_ar else "en-US")
        for tag in ("a:ea", "a:cs"):
            el = rpr.makeelement(qn(tag), {"typeface": FONT})
            rpr.append(el)

    def para(tf, first, runs, size, align=None, sa=0.0, ls=1.12, bullet=None):
        p = tf.paragraphs[0] if first else tf.add_paragraph()
        ppr = p._p.get_or_add_pPr()
        ppr.set("rtl", "1" if is_ar else "0")
        if align is None:
            align = PP_ALIGN.RIGHT if is_ar else PP_ALIGN.LEFT
        p.alignment = align
        p.line_spacing = ls
        p.space_after = Pt(sa)
        if bullet is not None:
            ppr.set("marL", str(int(Inches(0.22))))
            ppr.set("indent", str(-int(Inches(0.22))))
            ppr.append(parse_xml(
                f'<a:buClr {nsdecls("a")}><a:srgbClr val="{bullet}"/></a:buClr>'))
            ppr.append(parse_xml(f'<a:buFont {nsdecls("a")} typeface="Arial"/>'))
            ppr.append(parse_xml(f'<a:buChar {nsdecls("a")} char="&#8226;"/>'))
        for text, bold, color in runs:
            r = p.add_run()
            r.text = text
            style_run(r, size, color, bold)
        return p

    def fit_size(items, width, height, max_pt, min_pt, bullet=True, label=""):
        indent = 0.22 if bullet else 0.0
        pt = max_pt
        while pt >= min_pt - 1e-6:
            total = 0.0
            cpl = max(1, (width - indent) * 72 / (pt * 0.53))
            for lead, text in items:
                n = len(lead) + (1 if lead else 0) + len(text)
                lines = math.ceil(n / cpl)
                total += lines * pt * 1.2 * 1.12 / 72 + pt * 0.35 / 72
            if total <= height:
                return pt
            pt -= 0.5
        warnings.append(f"[{lang}] text may overflow: {label} (min {min_pt}pt)")
        return min_pt

    def textbox(g, l, t, w, h, anchor=MSO_ANCHOR.TOP):
        tb = g.shapes.add_textbox(Inches(mx(l, w)), Inches(t), Inches(w), Inches(h))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
        tf.vertical_anchor = anchor
        return tf

    def box(g, l, t, w, h, fill, line=None, lw=1.0, radius=0.15, alpha=None,
            shape=MSO_SHAPE.ROUNDED_RECTANGLE):
        s = g.shapes.add_shape(shape, Inches(mx(l, w)), Inches(t), Inches(w), Inches(h))
        if shape == MSO_SHAPE.ROUNDED_RECTANGLE:
            s.adjustments[0] = min(0.5, radius / min(w, h))
        s.fill.solid()
        s.fill.fore_color.rgb = fill
        if line is None:
            s.line.fill.background()
        else:
            s.line.color.rgb = line
            s.line.width = Pt(lw)
        s.shadow.inherit = False
        if alpha is not None:
            clr = s._element.spPr.find(qn("a:solidFill")).find(qn("a:srgbClr"))
            clr.append(parse_xml(f'<a:alpha {nsdecls("a")} val="{int(alpha * 1000)}"/>'))
        return s

    def pill(g, l, t, w, h, text, fill, color, size, line=None, bold=True):
        s = box(g, l, t, w, h, fill, line, 1.0, radius=h / 2)
        tf = s.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = 0
        tf.margin_left = tf.margin_right = Inches(0.06)
        tf.vertical_anchor = MSO_ANCHOR.MIDDLE
        para(tf, True, [(text, bold, color)], size, align=PP_ALIGN.CENTER)
        return s

    def card(slide, l, t, w, h, title, accent):
        g = slide.shapes.add_group_shape()
        box(g, l, t, w, h, BG_CARD, accent, 1.25, 0.2)
        box(g, l + 0.22, t + 0.17, 0.07, 0.28, accent, None, radius=0.035)
        tf = textbox(g, l + 0.42, t + 0.12, w - 0.64, 0.36, MSO_ANCHOR.MIDDLE)
        para(tf, True, [(title, True, accent)], 15)
        return g

    def bullets(g, l, t, w, h, items, accent, max_pt=13.0, min_pt=11.0, label=""):
        pt = fit_size(items, w, h, max_pt, min_pt, True, label)
        tf = textbox(g, l, t, w, h)
        for i, (lead, text) in enumerate(items):
            para(tf, i == 0, [(lead + " ", True, GOLD), (text, False, SOFT)], pt,
                 sa=pt * 0.35, bullet=str(accent))

    def image_card(slide, path, l, t, w, h, caption, accent):
        g = slide.shapes.add_group_shape()
        box(g, l, t, w, h, BG_CARD, accent, 1.25, 0.2)
        iw, ih = Image.open(path).size
        avail_w, avail_h = w - 0.3, h - 0.3 - 0.3
        scale = min(avail_w / iw, avail_h / ih)
        pw, ph = iw * scale, ih * scale
        pl = l + (w - pw) / 2
        pt = t + 0.15 + (avail_h - ph) / 2
        pic = g.shapes.add_picture(path, Inches(mx(pl, pw)), Inches(pt), Inches(pw), Inches(ph))
        pic.auto_shape_type = MSO_SHAPE.ROUNDED_RECTANGLE
        prst = pic._element.spPr.find(qn("a:prstGeom"))
        for old in prst.findall(qn("a:avLst")):
            prst.remove(old)
        adj = int(min(50000, 0.14 / min(pw, ph) * 100000))
        prst.append(parse_xml(
            f'<a:avLst {nsdecls("a")}><a:gd name="adj" fmla="val {adj}"/></a:avLst>'))
        tf = textbox(g, l + 0.15, t + h - 0.36, w - 0.3, 0.26, MSO_ANCHOR.MIDDLE)
        para(tf, True, [(caption, False, GRAY)], 11, align=PP_ALIGN.CENTER)
        return g

    # ---- slide scaffolding ---------------------------------------------------
    def new_slide(n, title, sub):
        s = prs.slides.add_slide(blank)
        bg = s.background.fill
        bg.gradient()
        bg.gradient_angle = 90.0
        bg.gradient_stops[0].color.rgb = BG_TOP
        bg.gradient_stops[1].color.rgb = BG_BOTTOM
        # decorative translucent circles
        box(s, 10.3, -1.7, 4.2, 4.2, TEAL, None, alpha=9, shape=MSO_SHAPE.OVAL)
        box(s, -1.3, 5.7, 3.6, 3.6, GOLD_DEEP, None, alpha=7, shape=MSO_SHAPE.OVAL)
        box(s, MARGIN, 1.6, SW - 2 * MARGIN, 0.03, TEAL, None, radius=0.015, alpha=45)

        hg = s.shapes.add_group_shape()
        pill(hg, MARGIN, 0.3, 4.6, 0.34, T["pill"].format(n=n), BG_CARD_ALT, TEAL, 11, line=TEAL)
        tf = textbox(hg, MARGIN, 0.72, SW - 2 * MARGIN, 0.5, MSO_ANCHOR.MIDDLE)
        para(tf, True, [(title, True, WHITE)], 26)
        tf = textbox(hg, MARGIN, 1.2, SW - 2 * MARGIN, 0.34, MSO_ANCHOR.MIDDLE)
        para(tf, True, [(sub, False, GOLD)], 14)

        fg = s.shapes.add_group_shape()
        tf = textbox(fg, MARGIN, 7.14, 9.5, 0.24, MSO_ANCHOR.MIDDLE)
        para(tf, True, [(T["footer"], False, GRAY)], 10)
        tf = textbox(fg, SW - MARGIN - 1.0, 7.14, 1.0, 0.24, MSO_ANCHOR.MIDDLE)
        para(tf, True, [(f"{n} / 3", True, TEAL)], 10,
             align=PP_ALIGN.LEFT if is_ar else PP_ALIGN.RIGHT)
        return s, [(hg, "fade")]

    def finish(slide, effects, note, n):
        animate(slide, effects)
        slide.notes_slide.notes_text_frame.text = note

    def animate(slide, effects, step=300, dur=500):
        cid = [4]

        def nid():
            cid[0] += 1
            return cid[0]

        parts = []
        for i, (shape, kind) in enumerate(effects):
            spid = shape.shape_id
            if kind == "wipe":
                preset, sub, flt = (22, 2, "wipe(right)") if is_ar else (22, 8, "wipe(left)")
            else:
                preset, sub, flt = 10, 0, "fade"
            a, b, c = nid(), nid(), nid()
            parts.append(
                f'<p:par><p:cTn id="{a}" presetID="{preset}" presetClass="entr" '
                f'presetSubtype="{sub}" fill="hold" nodeType="withEffect">'
                f'<p:stCondLst><p:cond delay="{i * step}"/></p:stCondLst><p:childTnLst>'
                f'<p:set><p:cBhvr><p:cTn id="{b}" dur="1" fill="hold"><p:stCondLst>'
                f'<p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="{spid}"/>'
                f'</p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName>'
                f'</p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>'
                f'<p:animEffect transition="in" filter="{flt}"><p:cBhvr>'
                f'<p:cTn id="{c}" dur="{dur}"/><p:tgtEl><p:spTgt spid="{spid}"/></p:tgtEl>'
                f'</p:cBhvr></p:animEffect></p:childTnLst></p:cTn></p:par>')
        timing = (
            f'<p:timing {nsdecls("p")}><p:tnLst><p:par><p:cTn id="1" dur="indefinite" '
            f'restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek">'
            f'<p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par>'
            f'<p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/>'
            f'<p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst>'
            f'<p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst>'
            f'<p:childTnLst>{"".join(parts)}</p:childTnLst></p:cTn></p:par>'
            f'</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn>'
            f'<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond>'
            f'</p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/>'
            f'</p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par>'
            f'</p:tnLst></p:timing>')
        transition = parse_xml(f'<p:transition {nsdecls("p")} spd="med"><p:fade/></p:transition>')
        slide._element.append(transition)
        slide._element.append(parse_xml(timing))

    # ======================================================================
    # SLIDE 1
    # ======================================================================
    d = T["s1"]
    s, fx = new_slide(1, d["title"], d["sub"])

    g = image_card(s, os.path.join(IMG, d["img"]), MARGIN, BODY_TOP, MEDIA_W, 2.3, d["cap"], TEAL)
    fx.append((g, "wipe"))
    accents = [RED, GOLD, TEAL]
    for i, (num, label) in enumerate(d["tiles"]):
        ty = BODY_TOP + 2.3 + 0.1 + i * 0.97
        tg = s.shapes.add_group_shape()
        box(tg, MARGIN, ty, MEDIA_W, 0.9, BG_CARD, accents[i], 1.1, 0.16)
        tf = textbox(tg, MARGIN + 0.1, ty, 1.5, 0.9, MSO_ANCHOR.MIDDLE)
        para(tf, True, [(num, True, accents[i])], 22, align=PP_ALIGN.CENTER)
        lw_ = MEDIA_W - 1.8
        tf = textbox(tg, MARGIN + 1.65, ty + 0.05, lw_, 0.8, MSO_ANCHOR.MIDDLE)
        para(tf, True, [(label, False, SOFT)], 11.5)
        fx.append((tg, "fade"))

    ca = card(s, CONTENT_L, BODY_TOP, CONTENT_W, 2.75, d["a_title"], RED)
    bullets(ca, CONTENT_L + 0.25, BODY_TOP + 0.6, CONTENT_W - 0.5, 2.75 - 0.68, d["a"], RED,
            label="s1 card A")
    fx.append((ca, "fade"))
    cb = card(s, CONTENT_L, BODY_TOP + 2.85, CONTENT_W, 2.5, d["b_title"], TEAL)
    bullets(cb, CONTENT_L + 0.25, BODY_TOP + 2.85 + 0.6, CONTENT_W - 0.5, 2.5 - 0.68, d["b"], TEAL,
            label="s1 card B")
    fx.append((cb, "fade"))
    finish(s, fx, T["notes"][0], 1)

    # ======================================================================
    # SLIDE 2
    # ======================================================================
    d = T["s2"]
    s, fx = new_slide(2, d["title"], d["sub"])

    g = image_card(s, os.path.join(IMG, d["img"]), MARGIN, BODY_TOP, MEDIA_W, 3.4, d["cap"], BLUE)
    fx.append((g, "wipe"))

    sy = BODY_TOP + 3.5
    sh = 7.05 - sy
    sg = card(s, MARGIN, sy, MEDIA_W, sh, d["score_title"], GOLD)
    tf = textbox(sg, MARGIN + 0.25, sy + 0.58, 1.75, 0.6, MSO_ANCHOR.MIDDLE)
    para(tf, True, [(d["score_big"], True, GOLD)], 32, align=PP_ALIGN.CENTER)
    pill(sg, MARGIN + 2.05, sy + 0.7, MEDIA_W - 2.3, 0.36, d["score_rating"], GREEN, NAVY_TEXT, 12)
    tf = textbox(sg, MARGIN + 0.25, sy + 1.25, MEDIA_W - 0.5, sh - 1.3)
    pt = fit_size([("", d["score_dims"])], MEDIA_W - 0.5, sh - 1.3, 12, 10.5, False, "s2 dims")
    para(tf, True, [(d["score_dims"], False, SOFT)], pt, align=PP_ALIGN.CENTER)
    fx.append((sg, "fade"))

    cw = (CONTENT_W - GAP) / 2
    ch = 1.65
    for i, (title, metric, desc, accent) in enumerate(d["cards"]):
        col, row = i % 2, i // 2
        cl = CONTENT_L + col * (cw + GAP)
        ct = BODY_TOP + row * (ch + 0.1)
        g = card(s, cl, ct, cw, ch, title, accent)
        tf = textbox(g, cl + 0.25, ct + 0.52, cw - 0.5, 0.3, MSO_ANCHOR.MIDDLE)
        para(tf, True, [(metric, True, GOLD)], 13.5)
        h_ = ch - 0.88 - 0.06
        pt = fit_size([("", desc)], cw - 0.5, h_, 12, 10.5, False, f"s2 {title}")
        tf = textbox(g, cl + 0.25, ct + 0.86, cw - 0.5, h_)
        para(tf, True, [(desc, False, SOFT)], pt)
        fx.append((g, "fade"))

    by = BODY_TOP + 2 * (ch + 0.1)
    bh = 7.05 - by
    bg_ = card(s, CONTENT_L, by, CONTENT_W, bh, d["bank_title"], GREEN)
    bullets(bg_, CONTENT_L + 0.25, by + 0.6, CONTENT_W - 0.5, bh - 0.68, d["bank"], GREEN,
            label="s2 bankability")
    fx.append((bg_, "fade"))
    finish(s, fx, T["notes"][1], 2)

    # ======================================================================
    # SLIDE 3
    # ======================================================================
    d = T["s3"]
    s, fx = new_slide(3, d["title"], d["sub"])

    g = image_card(s, os.path.join(IMG, d["img"]), MARGIN, BODY_TOP, MEDIA_W, 2.55, d["cap"], GOLD)
    fx.append((g, "wipe"))

    fy = BODY_TOP + 2.65
    fh = 7.05 - fy
    fg = card(s, MARGIN, fy, MEDIA_W, fh, d["fit_title"], TEAL)
    for i, (badge, name, desc) in enumerate(d["fit"]):
        ry = fy + 0.6 + i * 0.68
        pill(fg, MARGIN + 0.22, ry + 0.04, 0.95, 0.34, badge, GOLD, NAVY_TEXT, 12)
        tf = textbox(fg, MARGIN + 1.3, ry, MEDIA_W - 1.5, 0.66)
        para(tf, True, [(name, True, WHITE)], 12)
        para(tf, False, [(desc, False, GRAY)], 10.5)
    fx.append((fg, "fade"))

    bud_h, rdy_h, res_h = 1.8, 1.5, 1.75
    y1 = BODY_TOP
    y2 = y1 + bud_h + 0.1
    y3 = y2 + rdy_h + 0.1

    g = card(s, CONTENT_L, y1, CONTENT_W, bud_h, d["bud_title"], GOLD)
    bullets(g, CONTENT_L + 0.25, y1 + 0.58, CONTENT_W - 0.5, bud_h - 0.66, d["bud"], GOLD,
            max_pt=12.5, label="s3 budget")
    fx.append((g, "fade"))

    g = card(s, CONTENT_L, y2, CONTENT_W, rdy_h, d["rdy_title"], GREEN)
    bullets(g, CONTENT_L + 0.25, y2 + 0.58, CONTENT_W - 0.5, rdy_h - 0.66, d["rdy"], GREEN,
            max_pt=12.5, label="s3 readiness")
    fx.append((g, "fade"))

    g = card(s, CONTENT_L, y3, CONTENT_W, res_h, d["res_title"], WHITE)
    inner_w = CONTENT_W - 0.5
    chip_w = (inner_w - 3 * 0.1) / 4
    for i, (num, label) in enumerate(d["kpis"]):
        cx = CONTENT_L + 0.25 + i * (chip_w + 0.1)
        box(g, cx, y3 + 0.58, chip_w, 0.62, BG_CARD_ALT, TEAL, 1.0, 0.14)
        tf = textbox(g, cx, y3 + 0.6, chip_w, 0.58, MSO_ANCHOR.MIDDLE)
        para(tf, True, [(num, True, GOLD)], 16, align=PP_ALIGN.CENTER)
        para(tf, False, [(label, False, SOFT)], 10.5, align=PP_ALIGN.CENTER)
    ph_w = (inner_w - 2 * 0.1) / 3
    ph_colors = [TEAL, GREEN, GOLD]
    for i, text in enumerate(d["phases"]):
        px = CONTENT_L + 0.25 + i * (ph_w + 0.1)
        pill(g, px, y3 + 1.3, ph_w, 0.34, text, ph_colors[i], NAVY_TEXT, 11)
    fx.append((g, "fade"))
    finish(s, fx, T["notes"][2], 3)

    prs.save(output_path)
    print(f"[{lang.upper()}] saved: {output_path}")
    for w in warnings:
        print("  WARNING:", w)


def safe_copy(src, dst):
    try:
        shutil.copyfile(src, dst)
    except OSError as exc:  # file locked by a viewer / dev server
        print(f"  could not copy to {dst}: {exc}")


if __name__ == "__main__":
    ar = os.path.join(BASE, "Eco-Drain_WEFE_Nexus_5Min_Pitch_AR.pptx")
    en = os.path.join(BASE, "Eco-Drain_WEFE_Nexus_5Min_Pitch_EN.pptx")
    default = os.path.join(BASE, "Eco-Drain_WEFE_Nexus_5Min_Pitch.pptx")
    build_deck(ar, "ar")
    build_deck(en, "en")
    safe_copy(ar, default)
    report = os.path.join(BASE, "report")
    for src in (ar, en, default):
        safe_copy(src, os.path.join(report, os.path.basename(src)))
    print("All decks generated (AR / EN / default) and mirrored to report/.")
