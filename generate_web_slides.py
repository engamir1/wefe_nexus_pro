"""
Regenerates the interactive 5-minute pitch deck inside index.html from the SAME
content used for the PPTX files (generate_deck.CONTENT), so web and PowerPoint
always match. Each slide has a real RTL (Arabic) and LTR (English) variant,
a media column (figure + KPI tiles) and a content column, with staged animation.

Run:  python generate_web_slides.py
"""
import html
import os
import re
import shutil

from generate_deck import (BLUE, CONTENT, GOLD, GREEN, RED, TEAL, WHITE)

BASE = os.path.dirname(os.path.abspath(__file__))
INDEX = os.path.join(BASE, "index.html")
START, END = "<!-- DECK:START -->", "<!-- DECK:END -->"

TIMES = ["00:00 - 01:40", "01:40 - 03:20", "03:20 - 05:00"]


def hx(c):
    return "#" + str(c)


def e(s):
    return html.escape(s, quote=True)


class Counter:
    def __init__(self):
        self.i = 0

    def style(self):
        self.i += 1
        return f'style="--i:{self.i}"'


def card(c, color, title, body, cnt, extra_cls=""):
    return (f'<div class="deck-card anim {extra_cls}" {cnt.style()[:-1]};--c:{hx(color)}">'
            f'<h4>{e(title)}</h4>{body}</div>')


def ul(items):
    return "<ul>" + "".join(
        f"<li><strong>{e(a)}</strong> {e(b)}</li>" for a, b in items) + "</ul>"


def fig(path, cap, cnt, color=TEAL):
    return (f'<figure class="deck-fig anim wipe" {cnt.style()[:-1]};--c:{hx(color)}">'
            f'<img src="images/{path}" alt="{e(cap)}" loading="lazy">'
            f'<figcaption>{e(cap)}</figcaption></figure>')


def header(n, d, lang, cnt):
    badge = (f'الشريحة {n} من 3 • الوقت <bdi dir="ltr">{TIMES[n - 1]}</bdi>' if lang == "ar"
             else f"SLIDE {n} OF 3 • TIME: {TIMES[n - 1]}")
    return (f'<span class="slide-header-badge anim" {cnt.style()}>{badge}</span>'
            f'<h3 class="slide-main-title anim" {cnt.style()}>{e(d["title"])}</h3>'
            f'<div class="slide-meta-subtitle anim" {cnt.style()}>{e(d["sub"])}</div>')


def slide1(lang):
    d, cnt = CONTENT[lang]["s1"], Counter()
    media = fig(d["img"], d["cap"], cnt)
    for (num, label), col in zip(d["tiles"], [RED, GOLD, TEAL]):
        media += (f'<div class="deck-tile anim" {cnt.style()[:-1]};--c:{hx(col)}">'
                  f'<b>{e(num)}</b><span>{e(label)}</span></div>')
    content = (card(0, RED, d["a_title"], ul(d["a"]), cnt) +
               card(0, TEAL, d["b_title"], ul(d["b"]), cnt))
    return header(1, d, lang, cnt) + layout(media, content)


def slide2(lang):
    d, cnt = CONTENT[lang]["s2"], Counter()
    media = fig(d["img"], d["cap"], cnt, BLUE)
    score = (f'<div class="score-row"><span class="score-big">{e(d["score_big"])}</span>'
             f'<span class="pill" style="--c:{hx(GREEN)}">{e(d["score_rating"])}</span></div>'
             f'<p class="score-dims">{e(d["score_dims"])}</p>')
    media += card(0, GOLD, d["score_title"], score, cnt)
    grid = ""
    for title, metric, desc, col in d["cards"]:
        grid += card(0, col, title,
                     f'<div class="metric">{e(metric)}</div><p>{e(desc)}</p>', cnt)
    content = (f'<div class="deck-grid2">{grid}</div>' +
               card(0, GREEN, d["bank_title"], ul(d["bank"]), cnt))
    return header(2, d, lang, cnt) + layout(media, content)


def slide3(lang):
    d, cnt = CONTENT[lang]["s3"], Counter()
    media = fig(d["img"], d["cap"], cnt, GOLD)
    rows = "".join(
        f'<div class="fit-row"><span class="pill" style="--c:{hx(GOLD)}">{e(b)}</span>'
        f'<div><strong>{e(n)}</strong><small>{e(t)}</small></div></div>'
        for b, n, t in d["fit"])
    media += card(0, TEAL, d["fit_title"], rows, cnt)
    chips = "".join(
        f'<div class="kpi-chip"><b>{e(a)}</b><span>{e(b)}</span></div>' for a, b in d["kpis"])
    cols = [TEAL, GREEN, GOLD]
    phases = "".join(
        f'<span class="pill phase" style="--c:{hx(c)}">{e(p)}</span>'
        for p, c in zip(d["phases"], cols))
    res = f'<div class="kpi-grid">{chips}</div><div class="phase-row">{phases}</div>'
    content = (card(0, GOLD, d["bud_title"], ul(d["bud"]), cnt) +
               card(0, GREEN, d["rdy_title"], ul(d["rdy"]), cnt) +
               card(0, WHITE, d["res_title"], res, cnt))
    return header(3, d, lang, cnt) + layout(media, content)


def layout(media, content):
    return (f'<div class="deck-layout"><aside class="deck-media">{media}</aside>'
            f'<div class="deck-content">{content}</div></div>')


def build_html():
    out = []
    for n, fn in enumerate((slide1, slide2, slide3), start=1):
        active = " active" if n == 1 else ""
        out.append(f'<div class="slide-panel{active}" id="slidePanel{n}">')
        out.append(f'<div class="slide-lang" data-lang="ar" dir="rtl">{fn("ar")}</div>')
        out.append(f'<div class="slide-lang" data-lang="en" dir="ltr" style="display:none">{fn("en")}</div>')
        out.append("</div>")
    return "\n".join(out)


CSS = """
/* === DECK (generated layout, mirrors the PPTX design) === */
.slide-lang { text-align: start; }
.deck-layout { display: grid; grid-template-columns: minmax(250px, 0.85fr) 2.2fr; gap: 1.2rem; align-items: start; margin-top: 0.4rem; }
@media (max-width: 960px) { .deck-layout { grid-template-columns: 1fr; } }
.deck-media, .deck-content { display: flex; flex-direction: column; gap: 0.9rem; min-width: 0; }
.deck-fig { margin: 0; background: rgba(15,34,64,0.92); border: 1.5px solid var(--c, #15B5A4); border-radius: 20px; padding: 0.8rem; text-align: center; }
.deck-fig img { width: 100%; max-height: 300px; object-fit: contain; border-radius: 14px; display: block; background: #060E1A; }
.deck-fig figcaption { font-size: 0.8rem; color: var(--gray); margin-top: 0.55rem; }
.deck-tile { display: flex; align-items: center; gap: 0.8rem; background: rgba(15,34,64,0.92); border: 1.2px solid var(--c); border-radius: 16px; padding: 0.7rem 0.9rem; }
.deck-tile b { font-size: 1.45rem; font-weight: 900; color: var(--c); min-width: 5.6rem; text-align: center; line-height: 1.2; }
.deck-tile span { font-size: 0.88rem; color: #CBD5E1; line-height: 1.55; }
.deck-card { background: rgba(15,34,64,0.92); border: 1.3px solid var(--c); border-radius: 20px; padding: 1.1rem 1.3rem; }
.deck-card h4 { color: var(--c); font-size: 1.1rem; font-weight: 800; margin-bottom: 0.75rem; display: flex; align-items: center; gap: 0.6rem; line-height: 1.4; }
.deck-card h4::before { content: ""; width: 6px; height: 22px; border-radius: 4px; background: var(--c); flex: none; }
.deck-card ul { list-style: none; padding: 0; }
.deck-card li { font-size: 0.95rem; color: #CBD5E1; line-height: 1.7; margin-bottom: 0.5rem; position: relative; padding-inline-start: 1.2rem; }
.deck-card li::before { content: "•"; color: var(--c); font-weight: 900; position: absolute; inset-inline-start: 0; }
.deck-card li strong { color: var(--gold-light); font-weight: 700; }
.deck-card p { font-size: 0.92rem; line-height: 1.65; color: #CBD5E1; }
.deck-grid2 { display: grid; grid-template-columns: 1fr 1fr; gap: 0.9rem; }
@media (max-width: 700px) { .deck-grid2 { grid-template-columns: 1fr; } }
.metric { color: var(--gold-light); font-weight: 800; font-size: 1rem; margin-bottom: 0.4rem; }
.score-row { display: flex; align-items: center; justify-content: space-between; gap: 0.8rem; margin-bottom: 0.6rem; }
.score-big { font-size: 2.3rem; font-weight: 900; color: var(--gold-light); line-height: 1; }
.pill { display: inline-block; border-radius: 999px; padding: 0.28rem 0.95rem; font-weight: 800; font-size: 0.85rem; background: var(--c); color: #0A1628; text-align: center; white-space: nowrap; }
.score-dims { text-align: center; }
.fit-row { display: flex; gap: 0.8rem; align-items: flex-start; margin-bottom: 0.75rem; }
.fit-row .pill { min-width: 4.4rem; }
.fit-row strong { display: block; color: var(--white); font-size: 0.92rem; }
.fit-row small { display: block; color: var(--gray); font-size: 0.82rem; line-height: 1.5; }
.kpi-grid { display: grid; grid-template-columns: repeat(4, 1fr); gap: 0.6rem; }
@media (max-width: 700px) { .kpi-grid { grid-template-columns: repeat(2, 1fr); } }
.kpi-chip { background: rgba(24,49,88,0.95); border: 1px solid var(--teal-light); border-radius: 14px; text-align: center; padding: 0.65rem 0.3rem; }
.kpi-chip b { display: block; color: var(--gold-light); font-size: 1.2rem; font-weight: 900; }
.kpi-chip span { font-size: 0.8rem; color: #CBD5E1; }
.phase-row { display: grid; grid-template-columns: repeat(3, 1fr); gap: 0.5rem; margin-top: 0.8rem; }
@media (max-width: 700px) { .phase-row { grid-template-columns: 1fr; } }
.pill.phase { white-space: normal; font-size: 0.8rem; }
/* staged entrance animation (re-plays whenever a slide becomes active) */
.slide-panel.active .anim { opacity: 0; animation: deckIn 0.6s ease forwards; animation-delay: calc(var(--i, 0) * 0.14s); }
.slide-panel.active .anim.wipe { animation-name: deckWipe; }
.slide-panel.active .slide-lang[dir="rtl"] .anim.wipe { animation-name: deckWipeRtl; }
@keyframes deckIn { from { opacity: 0; transform: translateY(18px); } to { opacity: 1; transform: none; } }
@keyframes deckWipe { from { opacity: 0; clip-path: inset(0 100% 0 0 round 20px); } to { opacity: 1; clip-path: inset(0 0 0 0 round 20px); } }
@keyframes deckWipeRtl { from { opacity: 0; clip-path: inset(0 0 0 100% round 20px); } to { opacity: 1; clip-path: inset(0 0 0 0 round 20px); } }
@media (prefers-reduced-motion: reduce) { .slide-panel.active .anim { animation: none; opacity: 1; } }
"""


def patch():
    with open(INDEX, "r", encoding="utf-8", newline="") as f:
        src = f.read()
    block = f"{START}\n{build_html()}\n{END}"
    if START in src:
        src = re.sub(re.escape(START) + r".*?" + re.escape(END), lambda m: block, src, flags=re.S)
    else:
        pat = r'(<div class="slide-card-deck">)(.*?)(</div>\s*</div>\s*</section>)'
        src, n = re.subn(pat, lambda m: f"{m.group(1)}\n{block}\n        {m.group(3)}", src, count=1, flags=re.S)
        assert n == 1, "deck container not found"
    css_marker = "/* === DECK (generated layout"
    if css_marker in src:
        src = re.sub(r"/\* === DECK \(generated layout.*?(?=/\* === ANIMATIONS === \*/)", lambda m: CSS.lstrip() + "\n", src, flags=re.S)
    else:
        src = src.replace("/* === ANIMATIONS === */", CSS.lstrip() + "\n/* === ANIMATIONS === */", 1)
    with open(INDEX, "w", encoding="utf-8", newline="") as f:
        f.write(src)
    for name in ("index.html", "pitch.html"):
        shutil.copyfile(INDEX, os.path.join(BASE, "report", name))
    print("index.html deck regenerated; mirrored to report/index.html and report/pitch.html")


if __name__ == "__main__":
    patch()
