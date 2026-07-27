# Eco-Drain — WEFE Nexus Subsurface Drainage & IoT Anti-Tamper Project

> **Invisible Tamper-Proof Subsurface Drainage Retrofitting for Climate-Resilient Agriculture in Egypt's Nile Delta**

---

### 👨‍💼 Project Author & Supervision
- **Prepared by:** Eng. Mohamed Amir Abdelrahman (م / محمد أمير عبدالرحمن)
- **Affiliation:** Technical Office of the Vice Chairman of EPADP for Lower Egypt (المكتب الفني لنائب رئيس هيئة الصرف للوجه البحري — الهيئة المصرية العامة لمشروعات الصرف)
- **Group:** Group 10 (المجموعة العاشرة)
- **Lecture Series:** *Unlocking Finance for a Resilient WEFE (Water Energy Food Ecosystems) Nexus in Egypt*
- **Supervised by:** Dr. Hassan Tolba Aboelnga (د. حسن طلبة أبو النجا)

---

## 🌐 Live GitHub Pages Deployment

When hosted on **GitHub Pages**, this repository renders as an interactive, multi-language presentation pitch deck and scientific report.

### 🔗 Project Live Pages (After enabling GitHub Pages):
- **Interactive Pitch Deck (Main Landing Page):** `https://<your-username>.github.io/wefe_nexus_pro/`
- **IoT Anti-Tamper Simulation Demo:** `https://<your-username>.github.io/wefe_nexus_pro/iot_tamper_simulation.html`
- **Full Technical WEFE Report:** `https://<your-username>.github.io/wefe_nexus_pro/Eco-Drain_WEFE_Report_Group10.html`

---

## 🚀 How to Enable GitHub Pages (خطوات تفعيل GitHub Pages)

1. Push this repository to GitHub:
   ```bash
   git add .
   git commit -m "Configure project for GitHub Pages deployment"
   git push origin main
   ```
2. Navigate to your GitHub Repository -> **Settings** -> **Pages**.
3. Under **Build and deployment**:
   - **Source**: Select `Deploy from a branch`
   - **Branch**: Select `main` / `root (/)`
   - Click **Save**.
4. GitHub Pages will build automatically within 1-2 minutes. Your live site will be ready at:
   `https://<your-username>.github.io/<repository-name>/`

---

## 📑 Project Structure

```
wefe_nexus/
├── index.html                           # Main Interactive Investor Pitch Deck (GitHub Pages Entry)
├── iot_tamper_simulation.html           # Standalone Interactive 60fps IoT & Anti-Tamper Physics Simulation
├── Eco-Drain_WEFE_Report_Group10.html   # Comprehensive Technical Report (Bilingual)
├── Eco-Drain_WEFE_Report_Group10_AR.html# Arabic Technical Report
├── Eco-Drain_WEFE_Report_Group10_EN.html# English Technical Report
├── images/                              # High-resolution Technical Diagrams & Maps (fig1 to fig8)
├── report/                              # Sub-folder report mirror
│   ├── index.html
│   ├── pitch.html
│   ├── iot_tamper_simulation.html
│   └── Eco-Drain_WEFE_Report_Group10.html
└── README.md                            # Repository Documentation & GitHub Pages Guide
```

---

## 💡 Key Features

1. **Interactive Investor Pitch (`index.html`):**
   - Single-page application with smooth scroll animations.
   - Live animated metric counters (100,000 Feddans, $420M CAPEX, 92/100 WEFE Score).
   - Before/After interactive HTML5 Canvas diagrams comparing surface exposed manholes vs. invisible buried tamper-proof chambers.
   - Embedded real-time IoT simulation.

2. **IoT Anti-Tamper Physics Engine (`iot_tamper_simulation.html`):**
   - Real-time 60 FPS HTML5 Canvas cross-section animation of concrete inspection manholes, HDPE pipes, and geotextile envelopes.
   - Sensor physics for MPU6050 Accelerometer, Ultrasonic Water Level, Electromagnetic Flowmeter, and Soil EC probes.
   - Physical Tamper Attack Trigger simulating lid opening/shock (>2.5g) with encrypted AES-128 GPS timestamp telemetry via NB-IoT fallback.

3. **WEFE Nexus Integration:**
   - Closed-loop integration across Water, Energy (50 MW solar), Food (15-25% crop yield boost), and Ecosystems (constructed wetlands).
