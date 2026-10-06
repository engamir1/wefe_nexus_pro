import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

# Global Color Palette
BG_DARK = RGBColor(10, 22, 40)       # #0A1628
BG_CARD = RGBColor(15, 34, 64)       # #0F2240
CARD_BORDER = RGBColor(21, 181, 164) # #15B5A4
TEAL_LIGHT = RGBColor(21, 181, 164)  # #15B5A4
GOLD_LIGHT = RGBColor(240, 200, 96)  # #F0C860
GOLD = RGBColor(212, 168, 67)        # #D4A843
WHITE = RGBColor(248, 250, 252)      # #F8FAFC
GRAY = RGBColor(148, 163, 184)       # #94A3B8
GREEN = RGBColor(16, 185, 129)       # #10B981
RED = RGBColor(239, 68, 68)          # #EF4444

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
        return bg

    def add_header(slide, slide_num, title, subtitle):
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.15))
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
        
        # Pill badge
        p_badge = tf.paragraphs[0]
        p_badge.alignment = align
        if is_ar:
            p_badge.text = f"الشريحة {slide_num} من 3 | عرض تنفيذي في 5 دقائق | برنامج تمويل المرونة ومحور WEFE"
        else:
            p_badge.text = f"SLIDE {slide_num} OF 3 | 5-MIN EXECUTIVE PITCH | WEFE NEXUS PROFESSIONAL SERIES"
        p_badge.font.size = Pt(9.5)
        p_badge.font.bold = True
        p_badge.font.color.rgb = TEAL_LIGHT
        
        # Main Title
        p_title = tf.add_paragraph()
        p_title.alignment = align
        p_title.text = title
        p_title.font.size = Pt(19)
        p_title.font.bold = True
        p_title.font.color.rgb = WHITE
        
        # Subtitle
        p_sub = tf.add_paragraph()
        p_sub.alignment = align
        p_sub.text = subtitle
        p_sub.font.size = Pt(11)
        p_sub.font.color.rgb = GOLD_LIGHT

    def add_card(slide, left, top, width, height, title, title_color=TEAL_LIGHT):
        card = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
        card.fill.solid()
        card.fill.fore_color.rgb = BG_CARD
        card.line.color.rgb = CARD_BORDER
        card.line.width = Pt(1)
        
        # Title box
        tb = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.15), width - Inches(0.4), Inches(0.4))
        tf = tb.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
        p = tf.paragraphs[0]
        p.alignment = align
        p.text = title
        p.font.size = Pt(12)
        p.font.bold = True
        p.font.color.rgb = title_color
        
        # Content box
        content_box = slide.shapes.add_textbox(left + Inches(0.2), top + Inches(0.55), width - Inches(0.4), height - Inches(0.65))
        ctf = content_box.text_frame
        ctf.word_wrap = True
        ctf.margin_top = ctf.margin_bottom = ctf.margin_left = ctf.margin_right = 0
        return ctf

    # ==========================================
    # SLIDE 1
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1)
    if is_ar:
        add_header(s1, 1, 
                   "إيكو-درين (Eco-Drain): إحلال وتجديد شبكات الصرف المغطى الذكية ومقاومة التلاعب", 
                   "1. عنوان المشروع والأزمة الهيدروليكية | 3. الجهة القائدة، الشركاء وهيكل الحوكمة")
        
        c1 = add_card(s1, Inches(6.8), Inches(1.6), Inches(5.733), Inches(2.6), "🚨 1. بيان المشكلة المائية وتدهور الشبكات بالدلتا", RED)
        bullets1 = [
            "• الفقر المائي الحرج: حصة مصر الثابتة 55.5 مليار م³ لـ 106+ مليون نسمة تعني نصيب فرد 560 م³/سنة (تحت خط الفقر المائي العالمي 1,000 م³).",
            "• 2.3 مليون فدان متقادمة: شبكات الصرف المغطى تجاوزت عمرها التصميمي (20-25 سنة) مسببة تغدق الجذور وتملح التربة وخسارة 10-30% من المحاصيل.",
            "• أزمة البنية غير المرئية: الشبكات مدفونة تحت الأرض؛ لا يمكن فحصها بصرياً وتتأخر الصيانة لسنوات حتى تظهر أعراض تلف المحصول."
        ]
        for i, b in enumerate(bullets1):
            p = c1.paragraphs[0] if i == 0 else c1.add_paragraph()
            p.alignment = align
            p.text = b
            p.font.size = Pt(9.5)
            p.font.color.rgb = WHITE

        c2 = add_card(s1, Inches(6.8), Inches(4.35), Inches(5.733), Inches(2.65), "🛡️ 2. تحدي التلاعب الزراعي وابتكار الغرف غير المرئية", GOLD_LIGHT)
        bullets2 = [
            "• التعديات الزراعية العشوائية: سجلات هيئة الصرف (EPADP) توثق أن 25-40% من أعطال الصرف ناجمة عن سد المزارعين للمجمعات لحبس المياه للأرز.",
            "• ابتكار الغرف الغاطسة: استبدال الغرف السطحية بغرف مدفونة 30-50 سم تحت الأرض مع شرائح RFID ومواقع GPS للصيانة الرسمية فقط.",
            "• إنذار لحظي مشفر: حساسات MPU6050 ترصد أي اهتزاز أو حفر (>2.5g) وتبث إنذار طوارئ GPS مشفر بـ AES-128 عبر NB-IoT إلى مركز التحكم."
        ]
        for i, b in enumerate(bullets2):
            p = c2.paragraphs[0] if i == 0 else c2.add_paragraph()
            p.alignment = align
            p.text = b
            p.font.size = Pt(9.5)
            p.font.color.rgb = WHITE

        c3 = add_card(s1, Inches(0.8), Inches(1.6), Inches(5.7), Inches(5.4), "🏛️ 3. الجهة القائدة، الشركاء وهيكل الحوكمة المؤسسية", TEAL_LIGHT)
        p_lead = c3.paragraphs[0]
        p_lead.alignment = align
        p_lead.text = "• الجهة القائدة التنفيذية: وزارة الموارد المائية والري (MWRI) / الهيئة العامة لمشروعات الصرف (EPADP - المكتب الفني للوجه البحري)."
        p_lead.font.size = Pt(10)
        p_lead.font.bold = True
        p_lead.font.color.rgb = GOLD_LIGHT

        gov_items = [
            ("الشركاء الفنيون والزراعيون:", "معهد بحوث الصرف (DRI) للمواصفات والتحقق المستقل، ووزارة الزراعة (MALR) لتحسين المحاصيل والإرشاد."),
            ("شريك الطاقة والربط الشبكي:", "هيئة الطاقة المتجددة (NREA) وجهاز تنظيم الكهرباء (EgyptERA) لتطبيق صافي القياس (Net Metering)."),
            ("الحوكمة المجتمعية والميدانية:", "روابط مستخدمي المياه (WUAs) للمشاركة في إدارة المياه عبر تطبيق المحمول (Smart Farmer Drainage)."),
            ("هيكل الحوكمة متعدد المستويات:", "• لجنة وزارية توجيهية عليا (الري، الزراعة، البيئة، المالية).\n• وحدة مخصصة لإدارة المشروع (PIU) داخل هيئة الصرف.\n• لجان محلية بالمحافظات لروابط المزارعين في كفر الشيخ والدقهلية."),
            ("التوافق الاستراتيجي الوطني:", "متوافق تماماً مع مبادرة 'منظومة المياه والري 2.0'، ورؤية مصر 2030، واستراتيجية تغير المناخ 2050.")
        ]
        for h, desc in gov_items:
            p_item = c3.add_paragraph()
            p_item.alignment = align
            p_item.text = f"• {h} {desc}"
            p_item.font.size = Pt(9)
            p_item.font.color.rgb = WHITE

    else:
        add_header(s1, 1, 
                   "Eco-Drain: Invisible Tamper-Proof Subsurface Drainage Retrofitting", 
                   "1. Project Title & Problem Addressed | 3. Lead Institution & Key Governance Arrangements")
        
        c1 = add_card(s1, Inches(0.8), Inches(1.6), Inches(5.7), Inches(2.6), "🚨 1. National Problem Statement & Delta Water Crisis", RED)
        bullets1 = [
            "• Severe Water Scarcity: Egypt's quota is 55.5 BCM for 106M+ people -> 560 m³/capita (below 1,000 m³ water poverty line).",
            "• 2.3 Million Feddans Expired: Aging Nile Delta subsurface drainage (>20 yr design life) causes waterlogging & root-zone salinity, slashing crop yields by 10–30%.",
            "• The Invisible Infrastructure Gap: Buried networks cannot be visually monitored, resulting in silent failures undetected until topsoil salinization is severe."
        ]
        for i, b in enumerate(bullets1):
            p = c1.paragraphs[0] if i == 0 else c1.add_paragraph()
            p.alignment = align
            p.text = b
            p.font.size = Pt(9.5)
            p.font.color.rgb = WHITE

        c2 = add_card(s1, Inches(0.8), Inches(4.35), Inches(5.7), Inches(2.65), "🛡️ 2. The Tampering Challenge & Invisible Solution", GOLD_LIGHT)
        bullets2 = [
            "• The Vulnerability: EPADP records show 25–40% of drainage failures stem from unauthorized farmer modifications (blocking manholes for rice or extracting drainage).",
            "• Eco-Drain Core Innovation: Replacing surface manholes with invisible chambers buried 30–50 cm underground, accessible only via RFID tags & GPS.",
            "• Active Deterrence: Accelerometers detect digging (>2.5g), triggering instant AES-128 GPS alerts via NB-IoT to control centers, stopping sabotage immediately."
        ]
        for i, b in enumerate(bullets2):
            p = c2.paragraphs[0] if i == 0 else c2.add_paragraph()
            p.alignment = align
            p.text = b
            p.font.size = Pt(9.5)
            p.font.color.rgb = WHITE

        c3 = add_card(s1, Inches(6.8), Inches(1.6), Inches(5.733), Inches(5.4), "🏛️ 3. Lead Institution, Key Partners & Governance Structure", TEAL_LIGHT)
        p_lead = c3.paragraphs[0]
        p_lead.alignment = align
        p_lead.text = "• Lead Executive Institution: Ministry of Water Resources and Irrigation (MWRI) / Egyptian Public Authority for Drainage Projects (EPADP - Technical Office Lower Egypt)."
        p_lead.font.size = Pt(10)
        p_lead.font.bold = True
        p_lead.font.color.rgb = GOLD_LIGHT

        gov_items = [
            ("Key Technical Partners:", "Drainage Research Institute (DRI - NWRC) for soil telemetry, hydraulic criteria & independent MRV verification."),
            ("Energy & Regulatory Partner:", "New & Renewable Energy Authority (NREA) & EgyptERA for 50 MW net-metering grid interconnection."),
            ("Agricultural Extension:", "Ministry of Agriculture & Land Reclamation (MALR) for crop yield monitoring & farmer advisory."),
            ("Local Governance & Ownership:", "Farmers' Water User Associations (WUAs) co-managing water tables via the 'Smart Farmer' mobile app."),
            ("Multi-Tier Governance Structure:", "• Steering Committee: Inter-ministerial oversight (MWRI, MALR, MoE, MoF).\n• Dedicated PIU: Project Implementation Unit inside EPADP.\n• Multi-Stakeholder Platform: Local WUA committees in Kafr El-Sheikh & Dakahlia."),
            ("Strategic Alignment:", "Fully aligned with MWRI Water & Irrigation System 2.0 (Presidential Initiative), Egypt Vision 2030, and National Climate Change Strategy 2050.")
        ]
        for h, desc in gov_items:
            p_item = c3.add_paragraph()
            p_item.alignment = align
            p_item.text = f"• {h} {desc}"
            p_item.font.size = Pt(9)
            p_item.font.color.rgb = WHITE

    # ==========================================
    # SLIDE 2
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2)
    if is_ar:
        add_header(s2, 2, 
                   "تدخلات WEFE المتكاملة، مصفوفة الفرز الرسمية، وتعزيز الجدوى البنكية", 
                   "2. التدخلات الرئيسية ومنافع WEFE المغلقة | تصحيحات التقرير: مصفوفة الفرز، النظم البيئية وتعزيز التمويل")

        c_wefe = add_card(s2, Inches(0.8), Inches(1.6), Inches(11.733), Inches(2.7), "⚡ 1. التدخلات الأربعة المغلقة لمحور WEFE (المشروع التجريبي 100,000 فدان)", GOLD_LIGHT)
        wefe_ar = [
            ("💧 المياه (Water):", "إحلال 100 ألف فدان بمواسير HDPE وفلاتر جيوتكستيل (40+ سنة عمر). صمامات كهربائية تتحكم بالصعود الشعري (h=2γcosθ/ρgr) لتوفير 150M م³/سنة مياه ري (25-30%)."),
            ("☀️ الطاقة (Energy):", "محطات شمسية 50 MW موزعة على 25 محطة رفع تولد 90 GWh/سنة وتخفض فاتورة الكهرباء 40% مع تصدير الفائض للشبكة القومية بنظام صافي القياس."),
            ("🌾 الغذاء (Food):", "زيادة إنتاجية القمح والأرز والذرة 15-25%، مع غسيل أملاح ذكي تلقائي يفتح الصمام عند EC > 4 dS/m ويغلقه عند EC < 2 dS/m."),
            ("🌿 النظم البيئية (Ecosystems):", "20 محطة أراضٍ رطبة اصطناعية (4,760 فدان) تعالج 200M م³/سنة حيوياً وتحمي بحيرات الدلتا الشمالية، وتحتجز 100k طن CO₂ (شهادات Verra VM0033).")
        ]
        for i, (head, body) in enumerate(wefe_ar):
            p_h = c_wefe.paragraphs[0] if i == 0 else c_wefe.add_paragraph()
            p_h.alignment = align
            p_h.text = f"{head} {body}"
            p_h.font.size = Pt(9.5)
            p_h.font.color.rgb = WHITE

        c_screen = add_card(s2, Inches(6.8), Inches(4.45), Inches(5.733), Inches(2.55), "📊 2. مصفوفة فرز WEFE الرسمية وبطاقة قياس المرونة", TEAL_LIGHT)
        screen_ar = [
            ("بعد المياه (الدرجة: 10/10):", "وفر 150M م³/سنة؛ تثبيت منسوب الماء الأرضي عند 0.8-1.2م."),
            ("بعد الطاقة (الدرجة: 8/10):", "50 MW شمسية؛ 90 GWh/سنة؛ خفض تكاليف الضخ 40%."),
            ("بعد الغذاء (الدرجة: 10/10):", "زيادة المحاصيل 15-25%؛ دعم أمن الغذاء لـ 200,000 مزارع."),
            ("بعد النظم البيئية (الدرجة: 8/10):", "4,760 فدان رطبة؛ كربون أزرق؛ حماية الثروة السمكية."),
            ("الممكنات الرقمية (الدرجة: 9/10):", "5,000 عقدة IoT؛ توأم رقمي؛ رصد تلاعب لحظي مشفر."),
            ("التقييم المركب لمشروع WEFE:", "92 / 100 | بطاقة قياس المرونة: HIGH عبر الأبعاد الـ 5.")
        ]
        for i, (k, v) in enumerate(screen_ar):
            p_s = c_screen.paragraphs[0] if i == 0 else c_screen.add_paragraph()
            p_s.alignment = align
            p_s.text = f"• {k} {v}"
            p_s.font.size = Pt(8.5)
            p_s.font.color.rgb = GOLD_LIGHT if "التقييم المركب" in k else WHITE

        c_bank = add_card(s2, Inches(0.8), Inches(4.45), Inches(5.7), Inches(2.55), "💎 3. كيف يعزز تكامل WEFE الجدوى البنكية للمشروع", GREEN)
        bank_ar = [
            ("الاكتفاء الذاتي من تكاليف التشغيل:", "الطاقة الشمسية المصدرة تسدد فواتير كهرباء الطلمبات، مما يحمي نسبة خدمة الدين (DSCR = 1.35x)."),
            ("قيمة التكلفة المتجنبة (Avoided Costs):", "توفير 150M م³ مياه يعادل 52.5M دولار/سنة بتكلفة البديل (0.35$/م³)، محققاً فترة استرداد 0.48 سنة للري الذكي."),
            ("تنوع التدفقات النقدية الإيرادية:", "وفر طاقة ($8M) + تصدير شبكة ($3M) + مياه متجنبة ($5M) + أرصدة كربون ($0.5M) ورسوم صيانة."),
            ("انعدام النزاعات الاجتماعية:", "تطوير موضعي (In-situ) في أراضي الدلتا القديمة دون تحويل المياه بعيداً عن صغار المزارعين (بخلاف مشاريع بحر البقر والمحسمة).")
        ]
        for i, (k, v) in enumerate(bank_ar):
            p_b = c_bank.paragraphs[0] if i == 0 else c_bank.add_paragraph()
            p_b.alignment = align
            p_b.text = f"• {k} {v}"
            p_b.font.size = Pt(8.5)
            p_b.font.color.rgb = WHITE

    else:
        add_header(s2, 2, 
                   "Integrated WEFE Interventions, Nexus Value & Resilience Scorecard", 
                   "2. Main WEFE Interventions & Closed-Loop Benefits | Feedback Report Corrections: Screening Matrix, Ecosystem & Bankability")

        c_wefe = add_card(s2, Inches(0.8), Inches(1.6), Inches(11.733), Inches(2.7), "⚡ 1. The 4 Closed-Loop WEFE Interventions (100,000 Feddans Pilot)", GOLD_LIGHT)
        wefe_en = [
            ("💧 WATER:", "Retrofit 100k feddans with HDPE & geotextile filters (40+ yr life). Motorized gate valves control water table, leveraging Capillary Rise (h=2γcosθ/ρgr) to save 150M m³/yr (25–30% irrigation water)."),
            ("☀️ ENERGY:", "50 MW distributed solar PV across 25 pump stations producing 90 GWh/yr. Offsets 40% pumping electricity bill; surplus exported via EgyptERA Net Metering to national grid."),
            ("🌾 FOOD:", "Yield restored by 15–25% (wheat, rice, maize) across 100,000 feddans. Soil EC probe automated flushing (opens valve when EC > 4 dS/m; closes when EC < 2 dS/m)."),
            ("🌿 ECOSYSTEMS:", "20 constructed wetlands (4,760 feddans) treating 200M m³/yr of drainage water biologically (Phragmites/Typha). Protects Delta lakes fisheries + 100k tCO₂/yr (Verra VM0033).")
        ]
        for i, (head, body) in enumerate(wefe_en):
            p_h = c_wefe.paragraphs[0] if i == 0 else c_wefe.add_paragraph()
            p_h.alignment = align
            p_h.text = f"{head} {body}"
            p_h.font.size = Pt(9.5)
            p_h.font.color.rgb = WHITE

        c_screen = add_card(s2, Inches(0.8), Inches(4.45), Inches(5.7), Inches(2.55), "📊 2. Formal WEFE Screening Matrix & Resilience Scoring", TEAL_LIGHT)
        screen_en = [
            ("Water Dimension (Score: 10/10):", "150M m³/yr saved; water table stabilized at 0.8–1.2m."),
            ("Energy Dimension (Score: 8/10):", "50 MW solar; 90 GWh/yr clean power; 40% OPEX cut."),
            ("Food Dimension (Score: 10/10):", "15–25% yield boost; import substitution; 200k farmers."),
            ("Ecosystem Dimension (Score: 8/10):", "4,760 feddans wetlands; blue carbon; lake biodiversity."),
            ("Enabler Dimension (Score: 9/10):", "5,000 IoT nodes (ESP32); AI digital twin; tamper alert."),
            ("COMPOSITE WEFE SCORE: 92/100:", "RESILIENCE SCORECARD: HIGH across all 5 dimensions.")
        ]
        for i, (k, v) in enumerate(screen_en):
            p_s = c_screen.paragraphs[0] if i == 0 else c_screen.add_paragraph()
            p_s.alignment = align
            p_s.text = f"• {k} {v}"
            p_s.font.size = Pt(8.5)
            p_s.font.color.rgb = GOLD_LIGHT if "COMPOSITE" in k else WHITE

        c_bank = add_card(s2, Inches(6.8), Inches(4.45), Inches(5.733), Inches(2.55), "💎 3. How WEFE Integration Directly Drives Bankability", GREEN)
        bank_en = [
            ("OPEX Self-Sufficiency:", "Solar generation monetized via Net Metering pays pump electricity, slashing default risk and protecting DSCR (1.35x)."),
            ("Quantified Avoided-Cost Value:", "Saving 150M m³/yr is valued at $52.5M/yr (at $0.35/m³ alternative supply), yielding an investment payback of only 0.48 years for smart irrigation."),
            ("Diversified Revenue Streams:", "Energy savings ($8M/yr) + Grid export ($3M/yr) + Avoided water costs ($5M/yr) + Carbon offsets ($0.5M/yr) + Enhanced farmer tariffs."),
            ("Zero Downstream Displacement:", "Unlike mega-diversion plants (Bahr El-Baqar), Eco-Drain operates in-situ, eliminating community conflict and environmental safeguard rejection.")
        ]
        for i, (k, v) in enumerate(bank_en):
            p_b = c_bank.paragraphs[0] if i == 0 else c_bank.add_paragraph()
            p_b.alignment = align
            p_b.text = f"• {k} {v}"
            p_b.font.size = Pt(8.5)
            p_b.font.color.rgb = WHITE

    # ==========================================
    # SLIDE 3
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3)
    if is_ar:
        add_header(s3, 3, 
                   "هيكل رأس المال المختلط، الجاهزية، ومسار التوسع الإقليمي", 
                   "4. الميزانية والجهات الممولة | 5. الجاهزية والحقل التجريبي | 6. مؤشرات الأداء والتوسع الإقليمي")

        c_budg = add_card(s3, Inches(6.8), Inches(1.6), Inches(5.733), Inches(2.8), "💰 1. إجمالي التكلفة (420M$) وهيكل رأس المال المختلط", GOLD_LIGHT)
        b_ar = [
            ("إجمالي الاستثمار: 420 مليون دولار", "(4,200 دولار/فدان لكافة العناصر الهندسية والتقنية والبيئية)."),
            ("15% منح ومساعدات فنية (63M$):", "من GCF و AfDB و EU للحقل التجريبي والمنصة ودراسات الجدوى."),
            ("20% مساهمة حكومية عينية (84M$):", "أراضٍ ومحطات طلمبات قائمة وكوادر هيئة الصرف."),
            ("40% قروض ميسرة (168M$):", "البنك الدولي (NDP V)، وبنك التنمية الأفريقي (سداد 20-25 سنة)."),
            ("25% قروض تجارية وشراكة PPP (105M$):", "لمكون الطاقة الشمسية 50 MW مدعوماً باتفاقيات PPA."),
            ("مؤشرات الربحية الاقتصادية:", "EIRR = 21.4% | DSCR = 1.35x (تحت الضغط) | استرداد 8 سنوات.")
        ]
        for i, (k, v) in enumerate(b_ar):
            p_bu = c_budg.paragraphs[0] if i == 0 else c_budg.add_paragraph()
            p_bu.alignment = align
            p_bu.text = f"• {k} {v}"
            p_bu.font.size = Pt(8.8)
            p_bu.font.color.rgb = WHITE

        c_fit = add_card(s3, Inches(6.8), Inches(4.55), Inches(5.733), Inches(2.45), "🎯 2. تحليل ملاءمة جهات التمويل (Funding Call Fit)", TEAL_LIGHT)
        fit_ar = [
            ("صندوق المناخ الأخضر (GCF Adaptation):", "درجة 14/14 (قرار: GO) - تطابق كامل مع نافذة التكيف والتخفيف والتحول الرقمي."),
            ("صندوق التكيف (Adaptation Fund):", "درجة 13/14 (قرار: GO لمقترح مجزأ أو الحقل التجريبي 10 أفدنة)."),
            ("برنامج PRIMA وبوابة الاتحاد الأوروبي:", "ملاءمة مرتفعة جداً للتحول الرقمي الزراعي والتعاون الأورومتوسطي.")
        ]
        for i, (k, v) in enumerate(fit_ar):
            p_f = c_fit.paragraphs[0] if i == 0 else c_fit.add_paragraph()
            p_f.alignment = align
            p_f.text = f"• {k} {v}"
            p_f.font.size = Pt(8.8)
            p_f.font.color.rgb = WHITE

        c_read = add_card(s3, Inches(0.8), Inches(1.6), Inches(5.7), Inches(2.45), "🚀 3. حالة الجاهزية ووحدة الاختبار (10 أفدنة)", GREEN)
        r_ar = [
            ("النضج المؤسسي والتنفيذي:", "خبرة 50 عاماً لهيئة الصرف (EPADP) في شبكات 6 ملايين فدان تقضي على المخاطر التنفيذية."),
            ("وحدة الاختبار الحقلية (10 أفدنة):", "تصميم جاهز للتنفيذ الفوري على مجمع صرف مستقل لاختبار الصعود الشعري قبل التعميم."),
            ("نضج التقنيات (TRL):", "مستوى TRL 8-9 لكافة مكونات HDPE والطاقة الشمسية وحساسات ESP32 و NB-IoT."),
            ("الخطوة التنفيذية الفورية:", "طلب منحة تحضيرية (PPF Grant بقيمة 1.5 مليون دولار) من GCF/AfDB لدراسات كفر الشيخ.")
        ]
        for i, (k, v) in enumerate(r_ar):
            p_r = c_read.paragraphs[0] if i == 0 else c_read.add_paragraph()
            p_r.alignment = align
            p_r.text = f"• {k} {v}"
            p_r.font.size = Pt(8.8)
            p_r.font.color.rgb = WHITE

        c_kpi = add_card(s3, Inches(0.8), Inches(4.2), Inches(5.7), Inches(2.8), "📈 4. مؤشرات الأداء الرئيسية ومسار التوسع الإقليمي", WHITE)
        kpi_ar = [
            ("مؤشرات الأداء الكمية المستهدفة (100,000 فدان):", ""),
            ("• 150 مليون م³/سنة", "وفر مائي عبر الصرف المحكوم والصعود الشعري."),
            ("• 15-25% زيادة إنتاجية المحاصيل", "لأكثر من 200,000 مزارع مستفيد."),
            ("• 50 MW طاقة شمسية و 90 GWh/سنة", "كهرباء نظيفة + خفض 40% من تكلفة الضخ."),
            ("• خفض التعديات 80%", "وكشف الأعطال في ساعة واحدة بدلاً من أسابيع."),
            ("• خفض 100,000 طن CO₂/سنة", "وتوفير 5,000 فرصة عمل مباشرة."),
            ("مسار التوسع (3 مراحل):", "حقل تجريبي 10 أفدنة (سنة 1) -> تجربة ريادية 100 ألف فدان (سنوات 2-5) -> تعميم عبر 4.3M فدان بالدلتا ونقل التجربة إقليمياً للعراق والأردن وباكستان.")
        ]
        for i, (k, v) in enumerate(kpi_ar):
            p_k = c_kpi.paragraphs[0] if i == 0 else c_kpi.add_paragraph()
            p_k.alignment = align
            p_k.text = f"{k} {v}"
            p_k.font.size = Pt(8.5)
            p_k.font.color.rgb = GOLD_LIGHT if ("مؤشرات الأداء" in k or "مسار التوسع" in k) else WHITE

    else:
        add_header(s3, 3, 
                   "Blended Finance Capital Stack, Readiness, KPIs & Scaling Strategy", 
                   "4. Budget & Financiers | 5. Readiness Status | 6. Expected Results & Scaling Pathway")

        c_budg = add_card(s3, Inches(0.8), Inches(1.6), Inches(5.7), Inches(2.8), "💰 1. Total Investment ($420M) & Blended Capital Stack", GOLD_LIGHT)
        b_en = [
            ("Total CAPEX: USD 420 Million", "($4,200/feddan all-inclusive vs. $1,800/feddan baseline drainage)."),
            ("15% Grants / TA ($63M):", "GCF Readiness, AfDB, GEF for Feasibility, 10-feddan pilot, Digital Twin."),
            ("20% Government Equity ($84M):", "In-kind land for wetlands, existing pump stations, EPADP engineering staff."),
            ("40% Concessional Debt ($168M):", "World Bank (NDP V), AfDB, IsDB (20–25 yr tenor, long grace period)."),
            ("25% Commercial Debt / PPP ($105M):", "Private solar developers secured by solar PPA and energy savings."),
            ("Financial Returns:", "Project EIRR: 21.4% | DSCR: 1.35x (stress tested) | Payback: 8 years.")
        ]
        for i, (k, v) in enumerate(b_en):
            p_bu = c_budg.paragraphs[0] if i == 0 else c_budg.add_paragraph()
            p_bu.alignment = align
            p_bu.text = f"• {k} {v}"
            p_bu.font.size = Pt(8.8)
            p_bu.font.color.rgb = WHITE

        c_fit = add_card(s3, Inches(0.8), Inches(4.55), Inches(5.7), Inches(2.45), "🎯 2. Funding Call Fit Score Assessment", TEAL_LIGHT)
        fit_en = [
            ("Green Climate Fund (GCF) Adaptation Window:", "14/14 Score (Decision: GO) — Perfect match for transformational adaptation, mitigation co-benefits, and MRV automation."),
            ("Adaptation Fund:", "13/14 Score (Decision: GO for Component / 10-feddan Pilot Submission)."),
            ("PRIMA & EU Global Gateway:", "High Fit for Euro-Mediterranean digital water management, circular agriculture, and climate resilience.")
        ]
        for i, (k, v) in enumerate(fit_en):
            p_f = c_fit.paragraphs[0] if i == 0 else c_fit.add_paragraph()
            p_f.alignment = align
            p_f.text = f"• {k} {v}"
            p_f.font.size = Pt(8.8)
            p_f.font.color.rgb = WHITE

        c_read = add_card(s3, Inches(6.8), Inches(1.6), Inches(5.733), Inches(2.45), "🚀 3. Implementation Readiness & 10-Feddan Pilot Unit", GREEN)
        r_en = [
            ("Institutional Maturity:", "Built on EPADP's 50-year delivery platform (6M feddans installed); eliminates construction and execution risks."),
            ("10-Feddan Pilot Validation Unit:", "Engineered for immediate launch on a single isolated collector network to test dynamic capillary rise and IoT valve loops before scaling."),
            ("Technology Readiness Level:", "High TRL (TRL 8-9) for HDPE pipes, solar PV, and LoRaWAN/NB-IoT ESP32 hardware."),
            ("Immediate Next Step:", "Request Project Preparation Facility (PPF) Grant ($1.5M) from GCF/AfDB for site-specific feasibility in Kafr El-Sheikh.")
        ]
        for i, (k, v) in enumerate(r_en):
            p_r = c_read.paragraphs[0] if i == 0 else c_read.add_paragraph()
            p_r.alignment = align
            p_r.text = f"• {k} {v}"
            p_r.font.size = Pt(8.8)
            p_r.font.color.rgb = WHITE

        c_kpi = add_card(s3, Inches(6.8), Inches(4.2), Inches(5.733), Inches(2.8), "📈 4. Main Expected Results & Regional Scaling Pathway", WHITE)
        kpi_en = [
            ("Core Quantified KPIs (Target: 100,000 Feddans):", ""),
            ("• 150 Million m³/yr water saved", "(smart controlled capillary drainage)."),
            ("• 15–25% crop yield boost", "across 200,000+ direct beneficiary farmers."),
            ("• 50 MW solar installed & 90 GWh/yr", "clean energy produced."),
            ("• 80% reduction in tampering incidents", "+ failure detection from weeks to 1 hour."),
            ("• 100,000 tCO₂e/yr emissions reduced", "via solar energy and restored wetlands."),
            ("Replication & Scaling Strategy:", "Phase 1: 10-feddan pilot (Yr 1) -> Phase 2: 100k feddans core pilot (Yrs 2-5) -> Phase 3: Scaling across 4.3M feddans of Delta drainage and exporting model to Iraq, Jordan, and Pakistan.")
        ]
        for i, (k, v) in enumerate(kpi_en):
            p_k = c_kpi.paragraphs[0] if i == 0 else c_kpi.add_paragraph()
            p_k.alignment = align
            p_k.text = f"{k} {v}"
            p_k.font.size = Pt(8.5)
            p_k.font.color.rgb = GOLD_LIGHT if ("Core" in k or "Replication" in k) else WHITE

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
