import collections.abc
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

def create_deck(output_path):
    prs = Presentation()
    prs.slide_width = Inches(13.333)
    prs.slide_height = Inches(7.5)

    # Color Palette
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

    blank_layout = prs.slide_layouts[6]

    def add_bg(slide):
        bg = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, 0, 0, Inches(13.333), Inches(7.5))
        bg.fill.solid()
        bg.fill.fore_color.rgb = BG_DARK
        bg.line.fill.background()
        return bg

    def add_header(slide, slide_num, title, subtitle):
        # Header Box
        header_box = slide.shapes.add_textbox(Inches(0.8), Inches(0.4), Inches(11.733), Inches(1.1))
        tf = header_box.text_frame
        tf.word_wrap = True
        tf.margin_top = tf.margin_bottom = tf.margin_left = tf.margin_right = 0
        
        # Pill badge
        p_badge = tf.paragraphs[0]
        p_badge.text = f"SLIDE {slide_num} OF 3 | 5-MIN EXECUTIVE PITCH | WEFE NEXUS PROFESSIONAL SERIES"
        p_badge.font.size = Pt(9.5)
        p_badge.font.bold = True
        p_badge.font.color.rgb = TEAL_LIGHT
        
        # Main Title
        p_title = tf.add_paragraph()
        p_title.text = title
        p_title.font.size = Pt(20)
        p_title.font.bold = True
        p_title.font.color.rgb = WHITE
        
        # Subtitle
        p_sub = tf.add_paragraph()
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
    # SLIDE 1: Strategic Framing, Problem & Governance
    # ==========================================
    s1 = prs.slides.add_slide(blank_layout)
    add_bg(s1)
    add_header(s1, 1, 
               "Eco-Drain: Invisible Tamper-Proof Subsurface Drainage Retrofitting", 
               "1. Project Title & Problem Addressed | 3. Lead Institution & Key Governance Arrangements")

    # Card 1.1: Core Problem & National Water Urgency (Left Top)
    c1 = add_card(s1, Inches(0.8), Inches(1.6), Inches(5.7), Inches(2.6), "🚨 1. National Problem Statement & Delta Water Crisis", RED)
    p = c1.paragraphs[0]
    p.text = "• Severe Water Scarcity: Egypt's quota is 55.5 BCM for 106M+ people -> 560 m³/capita (below 1,000 m³ water poverty line)."
    p.font.size = Pt(10)
    p.font.color.rgb = WHITE
    
    p2 = c1.add_paragraph()
    p2.text = "• 2.3 Million Feddans Expired: Aging Nile Delta subsurface drainage (>20 yr design life) causes waterlogging & root-zone salinity, slashing crop yields by 10–30%."
    p2.font.size = Pt(10)
    p2.font.color.rgb = WHITE
    
    p3 = c1.add_paragraph()
    p3.text = "• The Invisible Infrastructure Gap: Buried networks cannot be visually monitored, resulting in silent failures undetected until topsoil salinization is severe."
    p3.font.size = Pt(10)
    p3.font.color.rgb = WHITE

    # Card 1.2: Tampering & Ground Vulnerability (Left Bottom)
    c2 = add_card(s1, Inches(0.8), Inches(4.35), Inches(5.7), Inches(2.65), "🛡️ 2. The Tampering Challenge & Invisible Solution", GOLD_LIGHT)
    p = c2.paragraphs[0]
    p.text = "• The Vulnerability: EPADP records show 25–40% of drainage failures stem from unauthorized farmer modifications (blocking manholes for rice or extracting drainage)."
    p.font.size = Pt(10)
    p.font.color.rgb = WHITE
    
    p2 = c2.add_paragraph()
    p2.text = "• Eco-Drain Core Innovation: Replacing surface manholes with invisible chambers buried 30–50 cm underground, accessible only via RFID tags & GPS."
    p2.font.size = Pt(10)
    p2.font.color.rgb = WHITE
    
    p3 = c2.add_paragraph()
    p3.text = "• Active Deterrence: Accelerometers detect digging (>2.5g), triggering instant AES-128 GPS alerts via NB-IoT to control centers, stopping sabotage immediately."
    p3.font.size = Pt(10)
    p3.font.color.rgb = WHITE

    # Card 1.3: Institutional Leadership & Multi-Tier Governance (Right)
    c3 = add_card(s1, Inches(6.8), Inches(1.6), Inches(5.733), Inches(5.4), "🏛️ 3. Lead Institution, Key Partners & Governance Structure", TEAL_LIGHT)
    
    p = c3.paragraphs[0]
    p.text = "• Lead Executive Institution: Ministry of Water Resources and Irrigation (MWRI) / Egyptian Public Authority for Drainage Projects (EPADP - Technical Office Lower Egypt)."
    p.font.size = Pt(10.5)
    p.font.bold = True
    p.font.color.rgb = GOLD_LIGHT
    
    items = [
        ("Key Technical Partners:", "Drainage Research Institute (DRI - NWRC) for soil telemetry, hydraulic criteria & independent MRV verification."),
        ("Energy & Regulatory Partner:", "New & Renewable Energy Authority (NREA) & EgyptERA for 50 MW net-metering grid interconnection."),
        ("Agricultural Extension:", "Ministry of Agriculture & Land Reclamation (MALR) for crop yield monitoring & farmer advisory."),
        ("Local Governance & Ownership:", "Farmers' Water User Associations (WUAs) co-managing water tables via the 'Smart Farmer' mobile app."),
        ("Multi-Tier Governance Structure:", "• Steering Committee: Inter-ministerial oversight (MWRI, MALR, MoE, MoF).\n• Dedicated PIU: Project Implementation Unit inside EPADP.\n• Multi-Stakeholder Platform: Local WUA committees in Kafr El-Sheikh & Dakahlia."),
        ("Strategic Alignment:", "Fully aligned with MWRI Water & Irrigation System 2.0 (Presidential Initiative), Egypt Vision 2030, and National Climate Change Strategy 2050.")
    ]
    for h, desc in items:
        p_item = c3.add_paragraph()
        p_item.text = f"• {h} {desc}"
        p_item.font.size = Pt(9.5)
        p_item.font.color.rgb = WHITE

    # ==========================================
    # SLIDE 2: WEFE Interventions, Nexus Value & Resilience
    # ==========================================
    s2 = prs.slides.add_slide(blank_layout)
    add_bg(s2)
    add_header(s2, 2, 
               "Integrated WEFE Interventions, Nexus Value & Resilience Scorecard", 
               "2. Main WEFE Interventions & Closed-Loop Benefits | Feedback Report Corrections: Screening Matrix, Ecosystem & Bankability")

    # Card 2.1: The 4 WEFE Interventions (Top Half)
    c_wefe = add_card(s2, Inches(0.8), Inches(1.6), Inches(11.733), Inches(2.7), "⚡ 1. The 4 Closed-Loop WEFE Interventions (100,000 Feddans Pilot)", GOLD_LIGHT)
    
    col_w = Inches(2.75)
    # We will format this into 4 distinct quadrants inside or paragraphs
    wefe_data = [
        ("💧 WATER (الـمـيـاه):", "Retrofit 100k feddans with HDPE & geotextile filters (40+ yr life). Motorized gate valves control water table, leveraging Capillary Rise (h=2γcosθ/ρgr) to save 150M m³/yr (25–30% irrigation water)."),
        ("☀️ ENERGY (الـطـاقـة):", "50 MW distributed solar PV across 25 pump stations producing 90 GWh/yr. Offsets 40% pumping electricity bill; surplus exported via EgyptERA Net Metering to national grid."),
        ("🌾 FOOD (الـغـذاء):", "Yield restored by 15–25% (wheat, rice, maize) across 100,000 feddans. Soil EC probe automated flushing (opens valve when EC > 4 dS/m; closes when EC < 2 dS/m)."),
        ("🌿 ECOSYSTEMS (الـبـيـئـة):", "20 constructed wetlands (4,760 feddans) treating 200M m³/yr of drainage water biologically (Phragmites/Typha). Protects Delta lakes fisheries + 100k tCO₂/yr (Verra VM0033).")
    ]
    
    for i, (head, body) in enumerate(wefe_data):
        p_h = c_wefe.paragraphs[0] if i == 0 else c_wefe.add_paragraph()
        p_h.text = f"{head} {body}"
        p_h.font.size = Pt(9.5)
        p_h.font.color.rgb = WHITE

    # Card 2.2: Formal WEFE Screening Matrix & Scoring (Bottom Left)
    c_screen = add_card(s2, Inches(0.8), Inches(4.45), Inches(5.7), Inches(2.55), "📊 2. Formal WEFE Screening Matrix & Resilience Scoring", TEAL_LIGHT)
    screen_items = [
        ("Water Dimension (Score: 10/10):", "150M m³/yr saved; water table stabilized at 0.8–1.2m."),
        ("Energy Dimension (Score: 8/10):", "50 MW solar; 90 GWh/yr clean power; 40% OPEX cut."),
        ("Food Dimension (Score: 10/10):", "15–25% yield boost; import substitution; 200k farmers."),
        ("Ecosystem Dimension (Score: 8/10):", "4,760 feddans wetlands; blue carbon; lake biodiversity."),
        ("Enabler Dimension (Score: 9/10):", "5,000 IoT nodes (ESP32); AI digital twin; tamper alert."),
        ("COMPOSITE WEFE SCORE: 92/100:", "RESILIENCE SCORECARD: HIGH across all 5 dimensions.")
    ]
    for i, (k, v) in enumerate(screen_items):
        p_s = c_screen.paragraphs[0] if i == 0 else c_screen.add_paragraph()
        p_s.text = f"• {k} {v}"
        p_s.font.size = Pt(8.5)
        p_s.font.color.rgb = GOLD_LIGHT if "COMPOSITE" in k else WHITE

    # Card 2.3: How WEFE Integration Improves Bankability (Bottom Right)
    c_bank = add_card(s2, Inches(6.8), Inches(4.45), Inches(5.733), Inches(2.55), "💎 3. How WEFE Integration Directly Drives Bankability", GREEN)
    bank_items = [
        ("OPEX Self-Sufficiency:", "Solar generation monetized via Net Metering pays pump electricity, slashing default risk and protecting DSCR (1.35x)."),
        ("Quantified Avoided-Cost Value:", "Saving 150M m³/yr is valued at $52.5M/yr (at $0.35/m³ alternative supply), yielding an investment payback of only 0.48 years for smart irrigation."),
        ("Diversified Revenue Streams:", "Energy savings ($8M/yr) + Grid export ($3M/yr) + Avoided water costs ($5M/yr) + Carbon offsets ($0.5M/yr) + Enhanced farmer tariffs."),
        ("Zero Downstream Displacement:", "Unlike mega-diversion plants (Bahr El-Baqar), Eco-Drain operates in-situ, eliminating community conflict and environmental safeguard rejection.")
    ]
    for i, (k, v) in enumerate(bank_items):
        p_b = c_bank.paragraphs[0] if i == 0 else c_bank.add_paragraph()
        p_b.text = f"• {k} {v}"
        p_b.font.size = Pt(8.5)
        p_b.font.color.rgb = WHITE

    # ==========================================
    # SLIDE 3: Budget, Capital Stack, Readiness & Scaling
    # ==========================================
    s3 = prs.slides.add_slide(blank_layout)
    add_bg(s3)
    add_header(s3, 3, 
               "Blended Finance Capital Stack, Readiness, KPIs & Scaling Strategy", 
               "4. Budget & Financiers | 5. Readiness Status | 6. Expected Results & Scaling Pathway")

    # Card 3.1: Budget & Blended Capital Stack (Left Top)
    c_budg = add_card(s3, Inches(0.8), Inches(1.6), Inches(5.7), Inches(2.8), "💰 1. Total Investment ($420M) & Blended Capital Stack", GOLD_LIGHT)
    b_items = [
        ("Total CAPEX: USD 420 Million", "($4,200/feddan all-inclusive vs. $1,800/feddan baseline drainage)."),
        ("15% Grants / TA ($63M):", "GCF Readiness, AfDB, GEF for Feasibility, 10-feddan pilot, Digital Twin."),
        ("20% Government Equity ($84M):", "In-kind land for wetlands, existing pump stations, EPADP engineering staff."),
        ("40% Concessional Debt ($168M):", "World Bank (NDP V), AfDB, IsDB (20–25 yr tenor, long grace period)."),
        ("25% Commercial Debt / PPP ($105M):", "Private solar developers secured by solar PPA and energy savings."),
        ("Financial Returns:", "Project EIRR: 21.4% | DSCR: 1.35x (stress tested) | Payback: 8 years.")
    ]
    for i, (k, v) in enumerate(b_items):
        p_bu = c_budg.paragraphs[0] if i == 0 else c_budg.add_paragraph()
        p_bu.text = f"• {k} {v}"
        p_bu.font.size = Pt(8.8)
        p_bu.font.color.rgb = WHITE

    # Card 3.2: Funding Call Fit Analysis (Left Bottom)
    c_fit = add_card(s3, Inches(0.8), Inches(4.55), Inches(5.7), Inches(2.45), "🎯 2. Funding Call Fit Score Assessment", TEAL_LIGHT)
    fit_items = [
        ("Green Climate Fund (GCF) Adaptation Window:", "14/14 Score (Decision: GO) — Perfect match for transformational adaptation, mitigation co-benefits, and MRV automation."),
        ("Adaptation Fund:", "13/14 Score (Decision: GO for Component / 10-feddan Pilot Submission)."),
        ("PRIMA & EU Global Gateway:", "High Fit for Euro-Mediterranean digital water management, circular agriculture, and climate resilience.")
    ]
    for i, (k, v) in enumerate(fit_items):
        p_f = c_fit.paragraphs[0] if i == 0 else c_fit.add_paragraph()
        p_f.text = f"• {k} {v}"
        p_f.font.size = Pt(8.8)
        p_f.font.color.rgb = WHITE

    # Card 3.3: Implementation Readiness & Pilot (Right Top)
    c_read = add_card(s3, Inches(6.8), Inches(1.6), Inches(5.733), Inches(2.45), "🚀 3. Implementation Readiness & 10-Feddan Pilot Unit", GREEN)
    r_items = [
        ("Institutional Maturity:", "Built on EPADP's 50-year delivery platform (6M feddans installed); eliminates construction and execution risks."),
        ("10-Feddan Pilot Validation Unit:", "Engineered for immediate launch on a single isolated collector network to test dynamic capillary rise and IoT valve loops before scaling."),
        ("Technology Readiness Level:", "High TRL (TRL 8-9) for HDPE pipes, solar PV, and LoRaWAN/NB-IoT ESP32 hardware."),
        ("Immediate Next Step:", "Request Project Preparation Facility (PPF) Grant ($1.5M) from GCF/AfDB for site-specific feasibility in Kafr El-Sheikh.")
    ]
    for i, (k, v) in enumerate(r_items):
        p_r = c_read.paragraphs[0] if i == 0 else c_read.add_paragraph()
        p_r.text = f"• {k} {v}"
        p_r.font.size = Pt(8.8)
        p_r.font.color.rgb = WHITE

    # Card 3.4: Main KPIs & 3-Stage Scaling Pathway (Right Bottom)
    c_kpi = add_card(s3, Inches(6.8), Inches(4.2), Inches(5.733), Inches(2.8), "📈 4. Main Expected Results & Regional Scaling Pathway", WHITE)
    kpi_items = [
        ("Core Quantified KPIs (Target: 100,000 Feddans):", ""),
        ("• 150 Million m³/yr water saved", "(smart controlled capillary drainage)."),
        ("• 15–25% crop yield boost", "across 200,000+ direct beneficiary farmers."),
        ("• 50 MW solar installed & 90 GWh/yr", "clean energy produced."),
        ("• 80% reduction in tampering incidents", "+ failure detection from weeks to 1 hour."),
        ("• 100,000 tCO₂e/yr emissions reduced", "via solar energy and restored wetlands."),
        ("Replication & Scaling Strategy:", "Phase 1: 10-feddan pilot (Yr 1) -> Phase 2: 100k feddans core pilot (Yrs 2-5) -> Phase 3: Scaling across 4.3M feddans of Delta drainage and exporting model to Iraq, Jordan, and Pakistan.")
    ]
    for i, (k, v) in enumerate(kpi_items):
        p_k = c_kpi.paragraphs[0] if i == 0 else c_kpi.add_paragraph()
        p_k.text = f"{k} {v}"
        p_k.font.size = Pt(8.5)
        p_k.font.color.rgb = GOLD_LIGHT if ("Core" in k or "Replication" in k) else WHITE

    prs.save(output_path)
    print(f"Presentation successfully created at: {output_path}")

if __name__ == "__main__":
    create_deck("d:/dev/wefe_nexus/wefe_nexus_pro/Eco-Drain_WEFE_Nexus_5Min_Pitch.pptx")
    # Also save in report/ for convenience
    create_deck("d:/dev/wefe_nexus/wefe_nexus_pro/report/Eco-Drain_WEFE_Nexus_5Min_Pitch.pptx")
