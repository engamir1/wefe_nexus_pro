import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Global Color Palette
BG_DARK = RGBColor(10, 22, 40)        # #0A1628 Deep Navy
BG_CARD = RGBColor(15, 34, 64)        # #0F2240 Card Navy
BG_CARD_ALT = RGBColor(22, 45, 82)    # #162D52 Elevated Navy
CARD_BORDER = RGBColor(21, 181, 164)  # #15B5A4 Teal
TEAL_LIGHT = RGBColor(21, 181, 164)   # #15B5A4
TEAL_DARK = RGBColor(14, 140, 127)    # #0E8C7F
GOLD_LIGHT = RGBColor(240, 200, 96)   # #F0C860 Gold Light
GOLD = RGBColor(212, 168, 67)         # #D4A843
WHITE = RGBColor(248, 250, 252)       # #F8FAFC
GRAY = RGBColor(148, 163, 184)        # #94A3B8
GREEN = RGBColor(16, 185, 129)        # #10B981 Emerald
RED = RGBColor(239, 68, 68)           # #EF4444 Crimson
BLUE_LIGHT = RGBColor(59, 130, 246)   # #3B82F6 Sky Blue

FONT_NAME = "Segoe UI"

def build_deck(output_path, lang='en'):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)
    blank_layout = prs.slide_layouts[6]
    is_ar = (lang == 'ar')
    align = PP_ALIGN.RIGHT if is_ar else PP_ALIGN.LEFT

    def add_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_DARK
        bg.line.fill.background()
        
        # Subtle top accent bar with rounded edges
        line = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.16), Inches(11.733), Inches(0.04))
        line.adjustments[0] = 0.5
        line.fill.solid()
        line.fill.fore_color.rgb = TEAL_LIGHT
        line.line.fill.background()
        return bg

    def add_header(slide, slide_num, title, subtitle):
        # Pill badge background with smooth rounded corners
        badge_w = Inches(4.6) if is_ar else Inches(5.3)
        badge_left = Inches(13.333 - 0.8 - (4.6 if is_ar else 5.3)) if is_ar else Inches(0.8)
        
        pill = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, badge_left, Inches(0.32), badge_w, Inches(0.32))
        pill.adjustments[0] = 0.5  # Fully rounded pill
        pill.fill.solid()
        pill.fill.fore_color.rgb = BG_CARD_ALT
        pill.line.color.rgb = TEAL_LIGHT
        pill.line.width = Pt(1.2)
        
        tf_pill = pill.text_frame
        tf_pill.word_wrap = False
        tf_pill.margin_top = tf_pill.margin_bottom = tf_pill.margin_left = tf_pill.margin_right = 0
        p_pill = tf_pill.paragraphs[0]
        p_pill.alignment = PP_ALIGN.CENTER
        p_pill.text = (f"الشريحة {slide_num} من 3 | عرض تنفيذي في 5 دقائق | محور WEFE المقاوم للمناخ" 
                       if is_ar else 
                       f"SLIDE {slide_num} OF 3 | 5-MIN EXECUTIVE PITCH | CLIMATE RESILIENT WEFE")
        p_pill.font.name = FONT_NAME
        p_pill.font.size = Pt(10.5)
        p_pill.font.bold = True
        p_pill.font.color.rgb = TEAL_LIGHT

        # Header titles textbox
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.72), Inches(11.733), Inches(0.88))
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
        
        # Main Title (Large & Bold: 22 pt)
        p_title = tf.paragraphs[0]
        p_title.alignment = align
        p_title.line_spacing = 1.15
        p_title.text = title
        p_title.font.name = FONT_NAME
        p_title.font.size = Pt(22)
        p_title.font.bold = True
        p_title.font.color.rgb = WHITE
        
        # Subtitle (12 pt)
        p_sub = tf.add_paragraph()
        p_sub.alignment = align
        p_sub.line_spacing = 1.2
        p_sub.text = subtitle
        p_sub.font.name = FONT_NAME
        p_sub.font.size = Pt(12)
        p_sub.font.color.rgb = GOLD_LIGHT

    def create_card(slide, left, top, width, height, title="", title_color=TEAL_LIGHT, border_color=CARD_BORDER, border_width=1.3, corner_radius=0.08, pill_text="", pill_color=GOLD_LIGHT):
        # Outer card with smooth rounded corners
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.adjustments[0] = corner_radius
        card.fill.solid()
        card.fill.fore_color.rgb = BG_CARD
        card.line.color.rgb = border_color
        card.line.width = Pt(border_width)

        top_offset = Inches(0.12)
        if title:
            # Card header text
            tb_title = slide.shapes.add_textbox(left + Inches(0.22), top + top_offset, width - Inches(0.44), Inches(0.38))
            tf_title = tb_title.text_frame
            tf_title.word_wrap = True
            tf_title.margin_top = tf_title.margin_bottom = tf_title.margin_left = tf_title.margin_right = 0
            p_t = tf_title.paragraphs[0]
            p_t.alignment = align
            p_t.text = title
            p_t.font.name = FONT_NAME
            p_t.font.size = Pt(14.5)
            p_t.font.bold = True
            p_t.font.color.rgb = title_color
            top_offset += Inches(0.36)

        if pill_text:
            # Highlight metric pill with rounded corners
            pill_shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left + Inches(0.20), top + top_offset, width - Inches(0.40), Inches(0.34))
            pill_shape.adjustments[0] = 0.35  # Smooth rounded pill
            pill_shape.fill.solid()
            pill_shape.fill.fore_color.rgb = BG_CARD_ALT
            pill_shape.line.color.rgb = border_color
            pill_shape.line.width = Pt(1)
            
            tf_p = pill_shape.text_frame
            tf_p.word_wrap = True
            tf_p.margin_top = tf_p.margin_bottom = tf_p.margin_left = tf_p.margin_right = 0
            p_pill_in = tf_p.paragraphs[0]
            p_pill_in.alignment = align
            p_pill_in.text = f" {pill_text} "
            p_pill_in.font.name = FONT_NAME
            p_pill_in.font.size = Pt(11.5)
            p_pill_in.font.bold = True
            p_pill_in.font.color.rgb = pill_color
            top_offset += Inches(0.40)

        # Content text box with generous interior margins and clear typography
        tb_content = slide.shapes.add_textbox(left + Inches(0.22), top + top_offset, width - Inches(0.44), height - top_offset - Inches(0.10))
        tf_content = tb_content.text_frame
        tf_content.word_wrap = True
        tf_content.margin_top = tf_content.margin_bottom = tf_content.margin_left = tf_content.margin_right = 0
        return tf_content

    def add_bullet(tf, title, desc, title_color=GOLD_LIGHT, text_color=WHITE, font_size=11.5, space_after=6.5):
        p = tf.paragraphs[0] if (len(tf.paragraphs) == 1 and not tf.paragraphs[0].text) else tf.add_paragraph()
        p.alignment = align
        p.line_spacing = 1.25
        p.space_after = Pt(space_after)

        if title:
            r_title = p.add_run()
            r_title.text = f"{title} "
            r_title.font.name = FONT_NAME
            r_title.font.size = Pt(font_size)
            r_title.font.bold = True
            r_title.font.color.rgb = title_color

        r_text = p.add_run()
        r_text.text = desc
        r_text.font.name = FONT_NAME
        r_text.font.size = Pt(font_size)
        r_text.font.color.rgb = text_color

    # =========================================================================
    # SLIDE 1: Title, Problem Statement, Innovation & Lead Governance
    # =========================================================================
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1)

    if is_ar:
        add_header(s1, 1,
                   "إيكو-درين (Eco-Drain): إحلال وتجديد شبكات الصرف المغطى الذكية ومقاومة التلاعب",
                   "1. بيان المشكلة المائية وتحدي التلاعب | 3. الجهة القائدة، الشركاء وهيكل الحوكمة المؤسسية")

        # Column 1 (Right): Problem & Tampering Solution
        c1 = create_card(s1, Inches(6.8), Inches(1.68), Inches(5.733), Inches(2.6), 
                         "🚨 1. بيان المشكلة المائية وتدهور الشبكات بالدلتا", RED, RED, 1.3, 0.08,
                         pill_text="⚠️ 560 م³/فرد حصة حرجة | 2.3 مليون فدان متقادمة", pill_color=GOLD_LIGHT)
        add_bullet(c1, "• الفقر المائي الحرج:", "حصة مصر 55.5 مليار م³ لـ 106+ مليون نسمة (560 م³/فرد، أقل بكثير من حد الفقر العالمي 1,000 م³).", GOLD_LIGHT, WHITE, 11.5, 6)
        add_bullet(c1, "• تقادم 2.3 مليون فدان:", "شبكات الصرف تجاوزت عمرها (20-25 سنة)، مسببة تغدق الجذور وتملح التربة وخسارة 10-30% من المحاصيل.", GOLD_LIGHT, WHITE, 11.5, 6)
        add_bullet(c1, "• أزمة البنية غير المرئية:", "شبكات مدفونة تحت الأرض يصعب فحصها بصرياً وتتأخر صيانتها لسنوات حتى يتلف المحصول.", GOLD_LIGHT, WHITE, 11.5, 0)

        c2 = create_card(s1, Inches(6.8), Inches(4.45), Inches(5.733), Inches(2.65), 
                         "🛡️ 2. تحدي التلاعب الزراعي وابتكار الغرف غير المرئية", GOLD_LIGHT, GOLD, 1.3, 0.08,
                         pill_text="🔒 غرف غاطسة مدفونة 30-50 سم | رصد اهتزاز >2.5g وإنذار NB-IoT", pill_color=TEAL_LIGHT)
        add_bullet(c2, "• التعديات العشوائية:", "سجلات هيئة الصرف توثق أن 25-40% من الأعطال ناجمة عن سد المزارعين للمناهل لحبس المياه لزراعة الأرز.", GOLD_LIGHT, WHITE, 11.5, 6)
        add_bullet(c2, "• ابتكار الغرف الغاطسة:", "استبدال الغرف السطحية بغرف مدفونة 30-50 سم مع شرائح RFID وإحداثيات GPS تمنع التلاعب نهائياً.", GOLD_LIGHT, WHITE, 11.5, 6)
        add_bullet(c2, "• الردع الذكي الفوري:", "حساسات MPU6050 ترصد أي اهتزاز أو حفر (>2.5g) وتبث إنذار طوارئ مشفر (AES-128) عبر NB-IoT إلى مركز التحكم.", GOLD_LIGHT, WHITE, 11.5, 0)

        # Column 2 (Left): Lead Institution, Partners & Governance
        c3 = create_card(s1, Inches(0.8), Inches(1.68), Inches(5.7), Inches(5.42), 
                         "🏛️ 3. الجهة القائدة، الشركاء وهيكل الحوكمة المؤسسية", TEAL_LIGHT, TEAL_LIGHT, 1.3, 0.06,
                         pill_text="👑 قيادة تنفيذية: وزارة الموارد المائية والري (MWRI) / هيئة الصرف (EPADP)", pill_color=GOLD_LIGHT)
        add_bullet(c3, "• الجهة القائدة التنفيذية:", "الهيئة المصرية العامة لمشروعات الصرف (EPADP - المكتب الفني للوجه البحري) بخبرة 50 عاماً في إدارة الشبكات.", GOLD_LIGHT, WHITE, 11.5, 7)
        add_bullet(c3, "• الشركاء الفنيون والزراعيون:", "معهد بحوث الصرف (DRI) للمواصفات والتحقق، ووزارة الزراعة (MALR) لتحسين إنتاجية المحاصيل.", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c3, "• شريك الطاقة والربط الشبكي:", "هيئة الطاقة المتجددة (NREA) وجهاز تنظيم الكهرباء (EgyptERA) لتطبيق صافي القياس (Net Metering).", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c3, "• الحوكمة المجتمعية والميدانية:", "روابط مستخدمي المياه (WUAs) للمشاركة الفعالة عبر تطبيق الهاتف (Smart Farmer Drainage).", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c3, "• هيكل الحوكمة متعدد المستويات:", "لجنة وزارية عليا توجيهية + وحدة إدارة مخصصة (PIU) بهيئة الصرف + لجان محلية بالمحافظات.", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c3, "• التوافق الاستراتيجي الوطني:", "متوافق تماماً مع مبادرة 'منظومة المياه والري 2.0'، ورؤية مصر 2030، واستراتيجية المناخ 2050.", TEAL_LIGHT, WHITE, 11.5, 0)

    else:
        add_header(s1, 1,
                   "Eco-Drain: Invisible Tamper-Proof Subsurface Drainage Retrofitting",
                   "1. Problem Statement & Tampering Challenge | 3. Lead Institution, Partners & Governance Arrangements")

        # Column 1 (Left): Problem & Tampering Solution
        c1 = create_card(s1, Inches(0.8), Inches(1.68), Inches(5.7), Inches(2.6), 
                         "🚨 1. National Problem Statement & Delta Water Crisis", RED, RED, 1.3, 0.08,
                         pill_text="⚠️ 560 m³/capita Scarcity | 2.3 Million Feddans Expired Life", pill_color=GOLD_LIGHT)
        add_bullet(c1, "• Critical Water Scarcity:", "Egypt's quota is 55.5 BCM for 106M+ people -> 560 m³/capita (below 1,000 m³ global water poverty line).", GOLD_LIGHT, WHITE, 11.5, 6)
        add_bullet(c1, "• 2.3 Million Feddans Expired:", "Nile Delta drainage networks exceed design life (>20 yr), causing root waterlogging & slashing yields by 10-30%.", GOLD_LIGHT, WHITE, 11.5, 6)
        add_bullet(c1, "• Invisible Infrastructure Gap:", "Buried networks cannot be visually monitored, resulting in silent failures undetected until crop loss occurs.", GOLD_LIGHT, WHITE, 11.5, 0)

        c2 = create_card(s1, Inches(0.8), Inches(4.45), Inches(5.7), Inches(2.65), 
                         "🛡️ 2. The Tampering Challenge & Invisible Solution", GOLD_LIGHT, GOLD, 1.3, 0.08,
                         pill_text="🔒 Buried Chambers 30-50 cm | >2.5g Shock Sensors & NB-IoT Telemetry", pill_color=TEAL_LIGHT)
        add_bullet(c2, "• Unauthorized Interference:", "EPADP records show 25-40% of failures stem from farmers blocking manholes to pool water for rice cultivation.", GOLD_LIGHT, WHITE, 11.5, 6)
        add_bullet(c2, "• Eco-Drain Core Innovation:", "Replacing surface structures with invisible chambers buried 30-50 cm deep, secured via RFID tags & GPS coordinates.", GOLD_LIGHT, WHITE, 11.5, 6)
        add_bullet(c2, "• Active Tamper Deterrence:", "MPU6050 sensors detect digging/shock (>2.5g), triggering instant AES-128 encrypted alerts via NB-IoT to control centers.", GOLD_LIGHT, WHITE, 11.5, 0)

        # Column 2 (Right): Lead Institution, Partners & Governance
        c3 = create_card(s1, Inches(6.8), Inches(1.68), Inches(5.733), Inches(5.42), 
                         "🏛️ 3. Lead Institution, Key Partners & Governance Structure", TEAL_LIGHT, TEAL_LIGHT, 1.3, 0.06,
                         pill_text="👑 Executive Lead: Ministry of Water Resources & Irrigation (MWRI) / EPADP", pill_color=GOLD_LIGHT)
        add_bullet(c3, "• Lead Executive Institution:", "Ministry of Water Resources and Irrigation (MWRI) / Egyptian Public Authority for Drainage Projects (EPADP).", GOLD_LIGHT, WHITE, 11.5, 7)
        add_bullet(c3, "• Technical & Agricultural Partners:", "Drainage Research Institute (DRI - NWRC) for independent MRV, and Ministry of Agriculture (MALR) for crop advisory.", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c3, "• Energy & Regulatory Partner:", "New & Renewable Energy Authority (NREA) & EgyptERA for 50 MW net-metering grid interconnection.", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c3, "• Community & Field Governance:", "Farmers' Water User Associations (WUAs) co-managing water tables via the 'Smart Farmer Drainage' mobile app.", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c3, "• Multi-Tier Governance Structure:", "Inter-Ministerial Steering Committee + dedicated PIU inside EPADP + Multi-Stakeholder Platforms in Kafr El-Sheikh.", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c3, "• National Strategic Alignment:", "Fully aligned with Presidential Initiative 'Irrigation 2.0', Egypt Vision 2030, and National Climate Strategy 2050.", TEAL_LIGHT, WHITE, 11.5, 0)

    # =========================================================================
    # SLIDE 2: WEFE Interventions, Screening Matrix & Bankability
    # =========================================================================
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2)

    if is_ar:
        add_header(s2, 2,
                   "تدخلات WEFE المتكاملة، مصفوفة الفرز الرسمية، وتعزيز الجدوى البنكية",
                   "2. التدخلات المغلقة الأربعة | تصحيحات التقرير: مصفوفة الفرز، النظم البيئية، وتعزيز الجدوى البنكية")

        # Top 4 Columns: WEFE Pillars
        p_w = Inches(2.77)
        
        # 1. Water
        c_w = create_card(s2, Inches(0.8), Inches(1.68), p_w, Inches(2.62), "💧 المياه (Water)", BLUE_LIGHT, BLUE_LIGHT, 1.2, 0.09,
                          pill_text="وفر 150M م³/سنة (25-30%)", pill_color=GOLD_LIGHT)
        add_bullet(c_w, "• التقنية:", "إحلال 100k فدان بمواسير HDPE وفلاتر جيوتكستيل (عمر 40+ سنة).", BLUE_LIGHT, WHITE, 11, 5)
        add_bullet(c_w, "• التحكم:", "صمامات ذكية تستفيد من الصعود الشعري لتغذية جذور المحاصيل ذاتياً.", BLUE_LIGHT, WHITE, 11, 0)

        # 2. Energy
        c_e = create_card(s2, Inches(0.8 + 2.98), Inches(1.68), p_w, Inches(2.62), "☀️ الطاقة (Energy)", GOLD_LIGHT, GOLD_LIGHT, 1.2, 0.09,
                          pill_text="50 MW شمسية | 90 GWh/سنة", pill_color=GOLD_LIGHT)
        add_bullet(c_e, "• المحطات:", "طاقة شمسية موزعة على 25 محطة رفع لتشغيل طلمبات الصرف.", GOLD_LIGHT, WHITE, 11, 5)
        add_bullet(c_e, "• الوفر:", "خفض 40% من تكلفة الضخ وتصدير الفائض بنظام صافي القياس.", GOLD_LIGHT, WHITE, 11, 0)

        # 3. Food
        c_f = create_card(s2, Inches(0.8 + 2.98*2), Inches(1.68), p_w, Inches(2.62), "🌾 الغذاء (Food)", GREEN, GREEN, 1.2, 0.09,
                          pill_text="+15-25% زيادة المحاصيل", pill_color=GOLD_LIGHT)
        add_bullet(c_f, "• المحاصيل:", "زيادة غلة القمح والذرة والأرز لأكثر من 200,000 مزارع.", GREEN, WHITE, 11, 5)
        add_bullet(c_f, "• غسيل ذكي:", "صمام آلي يفتح للغسيل عند EC > 4 ويغلق عند EC < 2 dS/m.", GREEN, WHITE, 11, 0)

        # 4. Ecosystems
        c_eco = create_card(s2, Inches(0.8 + 2.98*3), Inches(1.68), p_w, Inches(2.62), "🌿 النظم البيئية", TEAL_LIGHT, TEAL_LIGHT, 1.2, 0.09,
                          pill_text="100k طن كربون | 4,760 فدان", pill_color=GOLD_LIGHT)
        add_bullet(c_eco, "• أراضٍ رطبة:", "20 محطة معالجة طبيعية تنقي 200M م³ مياه صرف حيوياً.", TEAL_LIGHT, WHITE, 11, 5)
        add_bullet(c_eco, "• الأثر البيئي:", "حماية بحيرات الدلتا الشمالية والتنوع البيولوجي للأسماك.", TEAL_LIGHT, WHITE, 11, 0)

        # Bottom 2 Cards: Screening Matrix & Bankability
        c_sc = create_card(s2, Inches(6.8), Inches(4.45), Inches(5.733), Inches(2.65), 
                           "📊 مصفوفة فرز WEFE الرسمية وبطاقة قياس المرونة", TEAL_LIGHT, TEAL_LIGHT, 1.3, 0.08,
                           pill_text="التقييم المركب: 92 / 100 | تصنيف المرونة: HIGH", pill_color=GOLD_LIGHT)
        add_bullet(c_sc, "• درجات الأبعاد الخمسة:", "المياه: 10/10 | الطاقة: 8/10 | الغذاء: 10/10 | النظم البيئية: 8/10 | الرقمنة: 9/10.", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c_sc, "• الممكنات الرقمية:", "5,000 عقدة IoT مع توأم رقمي هيدروليكي للتحكم عن بُعد والتنبؤ بالأعطال لحظياً.", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c_sc, "• استيعاب الصدمات المناخية:", "مرونة عالية في مواجهة موجات الجفاف المتكررة وحماية التربة من نوبات التملح المفاجئ.", TEAL_LIGHT, WHITE, 11.5, 0)

        c_bk = create_card(s2, Inches(0.8), Inches(4.45), Inches(5.7), Inches(2.65), 
                           "💎 كيف يعزز تكامل WEFE الجدوى البنكية للمشروع", GREEN, GREEN, 1.3, 0.08,
                           pill_text="DSCR = 1.35x | فترة استرداد 0.48 سنة للري الذكي", pill_color=GOLD_LIGHT)
        add_bullet(c_bk, "• الاكتفاء الذاتي التشغيلي:", "الطاقة الشمسية تسدد تكاليف كهرباء الطلمبات، مما يحمي نسبة خدمة الدين (DSCR = 1.35x).", GOLD_LIGHT, WHITE, 11.5, 5)
        add_bullet(c_bk, "• قيمة التكلفة المتجنبة:", "توفير 150M م³ يعادل 52.5 مليون دولار سنوياً، محققاً استرداداً خلال 0.48 سنة للري الذكي.", GREEN, WHITE, 11.5, 5)
        add_bullet(c_bk, "• تنوع التدفقات النقدية:", "وفر طاقة ($8M) + تصدير شبكة ($3M) + مياه ($5M) + أرصدة كربون ($0.5M).", GREEN, WHITE, 11.5, 5)
        add_bullet(c_bk, "• انعدام النزاعات الاجتماعية:", "تطوير موضعي (In-situ) في أراضي الدلتا القديمة دون تهجير أو نزاع على المياه.", GREEN, WHITE, 11.5, 0)

    else:
        add_header(s2, 2,
                   "Integrated WEFE Interventions, Nexus Value & Resilience Scorecard",
                   "2. The 4 Closed-Loop Interventions | Feedback Corrections: Screening Matrix, Blue Carbon & Bankability")

        # Top 4 Columns: WEFE Pillars
        p_w = Inches(2.77)

        # 1. Water
        c_w = create_card(s2, Inches(0.8), Inches(1.68), p_w, Inches(2.62), "💧 WATER", BLUE_LIGHT, BLUE_LIGHT, 1.2, 0.09,
                          pill_text="150M m³/yr Saved (25-30%)", pill_color=GOLD_LIGHT)
        add_bullet(c_w, "• Tech:", "Retrofit 100k feddans with HDPE & geotextile filters (40+ yr design life).", BLUE_LIGHT, WHITE, 11, 5)
        add_bullet(c_w, "• Control:", "Smart gate valves regulate water table, leveraging natural capillary rise.", BLUE_LIGHT, WHITE, 11, 0)

        # 2. Energy
        c_e = create_card(s2, Inches(0.8 + 2.98), Inches(1.68), p_w, Inches(2.62), "☀️ ENERGY", GOLD_LIGHT, GOLD_LIGHT, 1.2, 0.09,
                          pill_text="50 MW Solar PV | 90 GWh/yr", pill_color=GOLD_LIGHT)
        add_bullet(c_e, "• Integration:", "Distributed solar across 25 pump stations powering lift pumps.", GOLD_LIGHT, WHITE, 11, 5)
        add_bullet(c_e, "• Grid Export:", "Cuts 40% electricity bill; surplus exported via Net Metering to grid.", GOLD_LIGHT, WHITE, 11, 0)

        # 3. Food
        c_f = create_card(s2, Inches(0.8 + 2.98*2), Inches(1.68), p_w, Inches(2.62), "🌾 FOOD", GREEN, GREEN, 1.2, 0.09,
                          pill_text="+15-25% Yield Boost", pill_color=GOLD_LIGHT)
        add_bullet(c_f, "• Crops:", "Restores yields for wheat, maize, and rice by mitigating root waterlogging.", GREEN, WHITE, 11, 5)
        add_bullet(c_f, "• Auto-Flushing:", "EC probes trigger automated flushing when EC > 4 dS/m; close at < 2 dS/m.", GREEN, WHITE, 11, 0)

        # 4. Ecosystems
        c_eco = create_card(s2, Inches(0.8 + 2.98*3), Inches(1.68), p_w, Inches(2.62), "🌿 ECOSYSTEMS", TEAL_LIGHT, TEAL_LIGHT, 1.2, 0.09,
                          pill_text="100k tCO₂/yr | 4,760 Feddans", pill_color=GOLD_LIGHT)
        add_bullet(c_eco, "• Wetlands:", "20 constructed wetlands (4,760 feddans) bio-treating 200M m³/yr drainage.", TEAL_LIGHT, WHITE, 11, 5)
        add_bullet(c_eco, "• Biodiversity:", "Protects northern Delta coastal lakes and commercial inland fisheries.", TEAL_LIGHT, WHITE, 11, 0)

        # Bottom 2 Cards: Screening Matrix & Bankability
        c_sc = create_card(s2, Inches(0.8), Inches(4.45), Inches(5.7), Inches(2.65), 
                           "📊 Formal WEFE Screening Matrix & Resilience Scoring", TEAL_LIGHT, TEAL_LIGHT, 1.3, 0.08,
                           pill_text="Composite Score: 92/100 | Resilience Rating: HIGH", pill_color=GOLD_LIGHT)
        add_bullet(c_sc, "• Dimension Scores:", "Water: 10/10 | Energy: 8/10 | Food: 10/10 | Ecosystems: 8/10 | Enablers: 9/10.", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c_sc, "• Digital Enablers:", "5,000 IoT nodes + hydraulic Digital Twin enabling real-time remote telemetry.", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c_sc, "• Shock Absorption:", "Robust buffer against prolonged droughts and rapid soil salinization waves.", TEAL_LIGHT, WHITE, 11.5, 0)

        c_bk = create_card(s2, Inches(6.8), Inches(4.45), Inches(5.733), Inches(2.65), 
                           "💎 How WEFE Integration Drives Project Bankability", GREEN, GREEN, 1.3, 0.08,
                           pill_text="DSCR = 1.35x | Smart Valve Payback = 0.48 Years", pill_color=GOLD_LIGHT)
        add_bullet(c_bk, "• OPEX Self-Sufficiency:", "Solar net-metering monetized electricity pays pump operating costs, protecting DSCR (1.35x).", GOLD_LIGHT, WHITE, 11.5, 5)
        add_bullet(c_bk, "• Avoided Cost Valuation:", "150M m³/yr saved equals $52.5M/yr, achieving an investment payback of 0.48 yr on smart gates.", GREEN, WHITE, 11.5, 5)
        add_bullet(c_bk, "• Diversified Revenues:", "Energy savings ($8M) + Grid export ($3M) + Avoided water ($5M) + Carbon ($0.5M).", GREEN, WHITE, 11.5, 5)
        add_bullet(c_bk, "• In-Situ Development:", "Zero displacement or community conflict, removing ESG social safeguard risks completely.", GREEN, WHITE, 11.5, 0)

    # =========================================================================
    # SLIDE 3: Budget & Financiers, Readiness & Scaling Pathway
    # =========================================================================
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3)

    if is_ar:
        add_header(s3, 3,
                   "هيكل رأس المال المختلط، الجاهزية، ومسار التوسع الإقليمي",
                   "4. الميزانية والجهات الممولة | 5. الجاهزية ووحدة الاختبار | 6. مؤشرات الأداء ومسار التوسع الإقليمي")

        # Top Row: Budget (Right) & Readiness (Left)
        c_b = create_card(s3, Inches(6.8), Inches(1.68), Inches(5.733), Inches(2.65), 
                          "💰 1. إجمالي التكلفة (420M$) وهيكل التمويل المختلط", GOLD_LIGHT, GOLD, 1.3, 0.08,
                          pill_text="CAPEX: 420M$ (4,200 $/فدان) | EIRR = 21.4% | استرداد 8 سنوات", pill_color=GOLD_LIGHT)
        add_bullet(c_b, "• 15% منح ومعونات (63M$):", "من GCF و AfDB و EU للحقل التجريبي والتوأم الرقمي وبناء القدرات.", GOLD_LIGHT, WHITE, 11.5, 5)
        add_bullet(c_b, "• 20% مساهمة حكومية (84M$):", "أراضٍ ومحطات طلمبات قائمة وكوادر هيئة الصرف (EPADP).", WHITE, WHITE, 11.5, 5)
        add_bullet(c_b, "• 40% قروض ميسرة (168M$):", "البنك الدولي (NDP V) وبنك التنمية الأفريقي (سداد 20-25 سنة).", WHITE, WHITE, 11.5, 5)
        add_bullet(c_b, "• 25% شراكة PPP (105M$):", "مستثمرون لمكون الطاقة الشمسية 50 MW مدعوماً بـ PPA | خدمة دين DSCR = 1.35x.", GOLD_LIGHT, WHITE, 11.5, 0)

        c_r = create_card(s3, Inches(0.8), Inches(1.68), Inches(5.7), Inches(2.65), 
                          "🚀 2. حالة الجاهزية ووحدة الاختبار (10 أفدنة)", GREEN, GREEN, 1.3, 0.08,
                          pill_text="خبرة 50 عاماً في 6M فدان | وحدة 10 أفدنة جاهزة فوراً", pill_color=GOLD_LIGHT)
        add_bullet(c_r, "• الخبرة المؤسسية المتراكمة:", "50 عاماً لهيئة الصرف (EPADP) في تنفيذ 6 ملايين فدان تقضي تماماً على مخاطر التنفيذ.", GOLD_LIGHT, WHITE, 11.5, 5)
        add_bullet(c_r, "• وحدة الاختبار الحقلية (10 أفدنة):", "تصميم هندسي متكامل جاهز للتنفيذ الفوري لاختبار الصعود الشعري والصمامات الذكية.", GREEN, WHITE, 11.5, 5)
        add_bullet(c_r, "• الجاهزية التكنولوجية (TRL 8-9):", "نضج تجاري كامل لمواسير HDPE والمحطات الشمسية وحساسات NB-IoT.", GREEN, WHITE, 11.5, 5)
        add_bullet(c_r, "• الخطوة التنفيذية الفورية:", "طلب منحة تحضيرية (PPF Grant بقيمة 1.5M$) من GCF/AfDB لدراسات كفر الشيخ.", GOLD_LIGHT, WHITE, 11.5, 0)

        # Bottom Row: Funding Call Fit (Right) & KPIs / Scaling (Left)
        c_fit = create_card(s3, Inches(6.8), Inches(4.45), Inches(5.733), Inches(2.65), 
                            "🎯 3. تحليل ملاءمة جهات التمويل (Funding Fit)", TEAL_LIGHT, TEAL_LIGHT, 1.3, 0.08,
                            pill_text="صندوق المناخ GCF: 14/14 (GO) | صندوق التكيف: 13/14", pill_color=GOLD_LIGHT)
        add_bullet(c_fit, "• صندوق المناخ الأخضر (GCF):", "درجة 14/14 (قرار: GO) — تطابق كامل مع نافذة التكيف وتخفيف الانبعاثات والتحول الرقمي.", GOLD_LIGHT, WHITE, 11.5, 6)
        add_bullet(c_fit, "• صندوق التكيف (Adaptation Fund):", "درجة 13/14 (قرار: GO) — مثالي لتمويل الحقل التجريبي وبناء القدرات المجتمعية.", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c_fit, "• برنامج PRIMA وبوابة الاتحاد الأوروبي:", "ملاءمة مرتفعة جداً للابتكار الزراعي المائي والتعاون الإقليمي الأورومتوسطي.", TEAL_LIGHT, WHITE, 11.5, 0)

        c_kpi = create_card(s3, Inches(0.8), Inches(4.45), Inches(5.7), Inches(2.65), 
                            "📈 4. مؤشرات الأداء الرئيسية ومسار التوسع الإقليمي", WHITE, WHITE, 1.3, 0.08,
                            pill_text="وفر 150M م³ | +25% غلة | 50 MW طاقة | 100k طن كربون", pill_color=GOLD_LIGHT)
        add_bullet(c_kpi, "• مؤشرات الأداء (100,000 فدان):", "150M م³/سنة وفر مائي • +15-25% غلة المحاصيل • 50 MW طاقة نظيفة.", GOLD_LIGHT, WHITE, 11.5, 6)
        add_bullet(c_kpi, "• الردع والاستدامة:", "خفض التعديات 80% • كشف الأعطال في ساعة واحدة • احتجاز 100,000 طن CO₂ سنوياً.", WHITE, WHITE, 11.5, 6)
        add_bullet(c_kpi, "• مسار التوسع الإقليمي (3 مراحل):", "حقل تجريبي 10 أفدنة (سنة 1) ⬅️ مشروع ريادي 100 ألف فدان (سنوات 2-5) ⬅️ تعميم عبر 4.3M فدان بالدلتا ونقل التجربة للعراق والأردن وباكستان.", TEAL_LIGHT, WHITE, 11, 0)

    else:
        add_header(s3, 3,
                   "Blended Finance Capital Stack, Readiness, KPIs & Scaling Strategy",
                   "4. Budget & Financiers | 5. Readiness & 10-Feddan Pilot | 6. Expected Results & Regional Scaling Pathway")

        # Top Row: Budget (Left) & Readiness (Right)
        c_b = create_card(s3, Inches(0.8), Inches(1.68), Inches(5.7), Inches(2.65), 
                          "💰 1. Total Investment ($420M) & Capital Stack", GOLD_LIGHT, GOLD, 1.3, 0.08,
                          pill_text="CAPEX: $420M ($4,200/feddan) | EIRR: 21.4% | Payback: 8 Yrs", pill_color=GOLD_LIGHT)
        add_bullet(c_b, "• 15% Grants / TA ($63M):", "GCF Readiness, AfDB, EU for 10-feddan pilot, Digital Twin & capacity.", GOLD_LIGHT, WHITE, 11.5, 5)
        add_bullet(c_b, "• 20% Government Equity ($84M):", "In-kind land, existing pump stations, EPADP engineering workforce.", WHITE, WHITE, 11.5, 5)
        add_bullet(c_b, "• 40% Concessional Debt ($168M):", "World Bank (NDP V), AfDB (20-25 yr tenor, long grace period).", WHITE, WHITE, 11.5, 5)
        add_bullet(c_b, "• 25% Commercial Debt / PPP ($105M):", "Private solar developers secured by PPA | DSCR = 1.35x.", GOLD_LIGHT, WHITE, 11.5, 0)

        c_r = create_card(s3, Inches(6.8), Inches(1.68), Inches(5.733), Inches(2.65), 
                          "🚀 2. Implementation Readiness & 10-Feddan Pilot", GREEN, GREEN, 1.3, 0.08,
                          pill_text="50-Yr Proven Track Record | 10-Feddan Unit Shovel-Ready", pill_color=GOLD_LIGHT)
        add_bullet(c_r, "• Institutional Delivery Track Record:", "EPADP's 50-year proven experience across 6M feddans eliminates construction risk.", GOLD_LIGHT, WHITE, 11.5, 5)
        add_bullet(c_r, "• 10-Feddan Pilot Unit Ready:", "Detailed engineering ready for immediate deployment on an isolated collector line.", GREEN, WHITE, 11.5, 5)
        add_bullet(c_r, "• Technology Maturity (TRL 8-9):", "High commercial maturity for HDPE pipes, solar systems, and NB-IoT telemetry.", GREEN, WHITE, 11.5, 5)
        add_bullet(c_r, "• Immediate Action Step:", "Request Project Preparation Facility (PPF $1.5M) from GCF/AfDB for Kafr El-Sheikh.", GOLD_LIGHT, WHITE, 11.5, 0)

        # Bottom Row: Funding Call Fit (Left) & KPIs / Scaling (Right)
        c_fit = create_card(s3, Inches(0.8), Inches(4.45), Inches(5.7), Inches(2.65), 
                            "🎯 3. Funding Call Fit Score Assessment", TEAL_LIGHT, TEAL_LIGHT, 1.3, 0.08,
                            pill_text="GCF Adaptation: 14/14 (GO) | Adaptation Fund: 13/14", pill_color=GOLD_LIGHT)
        add_bullet(c_fit, "• Green Climate Fund (GCF Adaptation):", "Score: 14/14 (Decision: GO) — Perfect match for transformational adaptation & digital MRV.", GOLD_LIGHT, WHITE, 11.5, 6)
        add_bullet(c_fit, "• Adaptation Fund:", "Score: 13/14 (Decision: GO) — Ideal for financing the 10-feddan pilot component.", TEAL_LIGHT, WHITE, 11.5, 6)
        add_bullet(c_fit, "• PRIMA & EU Global Gateway:", "High fit for Euro-Mediterranean water innovation and circular climate agriculture.", TEAL_LIGHT, WHITE, 11.5, 0)

        c_kpi = create_card(s3, Inches(6.8), Inches(4.45), Inches(5.733), Inches(2.65), 
                            "📈 4. Main Expected Results & Regional Scaling Pathway", WHITE, WHITE, 1.3, 0.08,
                            pill_text="150M m³ Water Saved | +25% Yield | 50 MW Solar | 100k tCO₂", pill_color=GOLD_LIGHT)
        add_bullet(c_kpi, "• Core Target KPIs (100,000 Feddans):", "150M m³/yr water saved • 15-25% crop yield increase • 50 MW clean power.", GOLD_LIGHT, WHITE, 11.5, 6)
        add_bullet(c_kpi, "• Active Security & Carbon:", "80% tampering cut • incident response in 1 hour • 100,000 tCO₂e/yr emissions offset.", WHITE, WHITE, 11.5, 6)
        add_bullet(c_kpi, "• 3-Stage Scaling Pathway:", "Phase 1: 10-feddan pilot (Yr 1) ➡️ Phase 2: 100k feddans (Yrs 2-5) ➡️ Phase 3: Scaling to 4.3M feddans in Nile Delta and exporting to Iraq, Jordan, and Pakistan.", TEAL_LIGHT, WHITE, 11, 0)

    prs.save(output_path)
    print(f"Presentation ({lang.upper()}) created at: {output_path}")

if __name__ == "__main__":
    # Generate Arabic Presentations
    build_deck("d:/dev/wefe_nexus/wefe_nexus_pro/Eco-Drain_WEFE_Nexus_5Min_Pitch_AR.pptx", lang='ar')
    build_deck("d:/dev/wefe_nexus/wefe_nexus_pro/report/Eco-Drain_WEFE_Nexus_5Min_Pitch_AR.pptx", lang='ar')

    # Generate English Presentations
    build_deck("d:/dev/wefe_nexus/wefe_nexus_pro/Eco-Drain_WEFE_Nexus_5Min_Pitch_EN.pptx", lang='en')
    build_deck("d:/dev/wefe_nexus/wefe_nexus_pro/report/Eco-Drain_WEFE_Nexus_5Min_Pitch_EN.pptx", lang='en')

    # Default fallback
    build_deck("d:/dev/wefe_nexus/wefe_nexus_pro/Eco-Drain_WEFE_Nexus_5Min_Pitch.pptx", lang='ar')
    build_deck("d:/dev/wefe_nexus/wefe_nexus_pro/report/Eco-Drain_WEFE_Nexus_5Min_Pitch.pptx", lang='ar')
    print("All presentations regenerated successfully with rounded corners and optimized typography!")
