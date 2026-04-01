#!/usr/bin/env python3
"""Generate a PowerPoint presentation on Vertical Approach in SaaS."""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# ── Color palette ──
DARK_BG = RGBColor(0x1B, 0x1B, 0x2F)
ACCENT_BLUE = RGBColor(0x00, 0x9E, 0xFF)
ACCENT_GREEN = RGBColor(0x00, 0xC9, 0x8D)
ACCENT_PURPLE = RGBColor(0x7C, 0x5C, 0xFC)
ACCENT_ORANGE = RGBColor(0xFF, 0x8C, 0x42)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT_GRAY = RGBColor(0xBB, 0xBB, 0xCC)
MEDIUM_GRAY = RGBColor(0x88, 0x88, 0x99)
CARD_BG = RGBColor(0x26, 0x26, 0x40)


def set_slide_bg(slide, color):
    bg = slide.background
    fill = bg.fill
    fill.solid()
    fill.fore_color.rgb = color


def add_shape(slide, left, top, width, height, fill_color, border_color=None, radius=None):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = fill_color
    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = Pt(1.5)
    else:
        shape.line.fill.background()
    return shape


def add_text_box(slide, left, top, width, height, text, font_size=18, color=WHITE,
                 bold=False, alignment=PP_ALIGN.LEFT, font_name="Calibri"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(font_size)
    p.font.color.rgb = color
    p.font.bold = bold
    p.font.name = font_name
    p.alignment = alignment
    return txBox


def add_bullet_slide_content(slide, bullets, left, top, width, height,
                              font_size=17, color=WHITE):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, bullet in enumerate(bullets):
        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()
        p.text = bullet
        p.font.size = Pt(font_size)
        p.font.color.rgb = color
        p.font.name = "Calibri"
        p.space_after = Pt(10)
        p.level = 0
    return txBox


# ═══════════════════════════════════════════════════════════════════
# SLIDE 1 – Title Slide
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])  # blank
set_slide_bg(slide, DARK_BG)

# Accent bar top
add_shape(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.08), ACCENT_BLUE)

# Title
add_text_box(slide, Inches(1), Inches(1.8), Inches(11), Inches(1.5),
             "The Vertical Approach in SaaS",
             font_size=44, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)

# Subtitle
add_text_box(slide, Inches(2), Inches(3.4), Inches(9), Inches(1),
             "A Strategic Guide to Building Industry-Specific Software Solutions",
             font_size=24, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

# Divider line
shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Inches(5), Inches(4.6), Inches(3.333), Pt(3))
shape.fill.solid()
shape.fill.fore_color.rgb = ACCENT_BLUE
shape.line.fill.background()

add_text_box(slide, Inches(2), Inches(5.2), Inches(9), Inches(0.8),
             "What every founder, product leader, and GTM team needs to know",
             font_size=18, color=MEDIUM_GRAY, alignment=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 2 – What Is a Vertical SaaS Approach?
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.9),
             "What Is a Vertical SaaS Approach?",
             font_size=36, color=ACCENT_BLUE, bold=True)

# Left column – definition
add_shape(slide, Inches(0.8), Inches(1.7), Inches(5.6), Inches(5), CARD_BG, ACCENT_BLUE)
add_text_box(slide, Inches(1.2), Inches(1.9), Inches(4.8), Inches(0.7),
             "Definition", font_size=22, color=ACCENT_BLUE, bold=True)
add_bullet_slide_content(slide, [
    "Software built for ONE specific industry or niche",
    "Deeply tailored workflows, compliance & terminology",
    "End-to-end solution replacing horizontal point tools",
    "Examples: Veeva (pharma), Procore (construction), Toast (restaurants)",
], Inches(1.2), Inches(2.7), Inches(4.8), Inches(3.8), font_size=16)

# Right column – vs horizontal
add_shape(slide, Inches(6.9), Inches(1.7), Inches(5.6), Inches(5), CARD_BG, ACCENT_PURPLE)
add_text_box(slide, Inches(7.3), Inches(1.9), Inches(4.8), Inches(0.7),
             "Vertical vs. Horizontal SaaS", font_size=22, color=ACCENT_PURPLE, bold=True)

comparisons = [
    "Vertical: Deep domain expertise  |  Horizontal: Broad feature set",
    "Vertical: Smaller TAM, higher win rate  |  Horizontal: Large TAM, more competition",
    "Vertical: Industry-specific compliance  |  Horizontal: General-purpose",
    "Vertical: Higher NRR & stickiness  |  Horizontal: Easier initial adoption",
]
add_bullet_slide_content(slide, comparisons,
                         Inches(7.3), Inches(2.7), Inches(4.8), Inches(3.8), font_size=15)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 3 – Why Go Vertical?
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.9),
             "Why Go Vertical? The Strategic Advantages",
             font_size=36, color=ACCENT_GREEN, bold=True)

advantages = [
    ("Higher Win Rates", "Prospects see you as the expert. You speak their language,\nunderstand their pain, and demo workflows they recognize.", ACCENT_BLUE),
    ("Stronger Retention", "Deep integration into daily operations = high switching costs.\nVertical SaaS often sees 120-140%+ NRR.", ACCENT_GREEN),
    ("Efficient GTM", "Concentrated buyer personas, focused conferences,\nindustry publications, and word-of-mouth referrals.", ACCENT_PURPLE),
    ("Pricing Power", "Industry-specific value justifies premium pricing.\nCustomers pay for outcomes, not generic features.", ACCENT_ORANGE),
]

for i, (title, desc, color) in enumerate(advantages):
    col = i % 4
    x = Inches(0.6 + col * 3.1)
    y = Inches(1.8)
    add_shape(slide, x, y, Inches(2.9), Inches(4.8), CARD_BG, color)
    add_text_box(slide, x + Inches(0.3), y + Inches(0.3), Inches(2.3), Inches(0.7),
                 title, font_size=20, color=color, bold=True)
    add_text_box(slide, x + Inches(0.3), y + Inches(1.2), Inches(2.3), Inches(3.2),
                 desc, font_size=15, color=LIGHT_GRAY)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 4 – Choosing Your Vertical
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.9),
             "Choosing the Right Vertical: Selection Criteria",
             font_size=36, color=ACCENT_ORANGE, bold=True)

criteria = [
    "Market Size & Density  --  Is the TAM large enough? Are buyers concentrated or fragmented?",
    "Underserved by Tech  --  Are incumbents still using spreadsheets, paper, or legacy on-prem systems?",
    "Regulatory Complexity  --  Industries with compliance needs (HIPAA, SOX, FDA) create moats for vertical players",
    "Willingness to Pay  --  Does the industry have healthy margins? Do they already budget for software?",
    "Domain Access  --  Do you have founder-market fit? Can you access early design partners?",
    "Workflow Standardization  --  Are core workflows consistent enough to build a scalable product?",
    "Expansion Potential  --  Can you layer on payments, financing, marketplace, or data products?",
]

for i, item in enumerate(criteria):
    y = Inches(1.7 + i * 0.75)
    # Number circle
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL, Inches(0.9), y + Inches(0.05), Inches(0.45), Inches(0.45))
    circle.fill.solid()
    circle.fill.fore_color.rgb = ACCENT_ORANGE
    circle.line.fill.background()
    tf = circle.text_frame
    tf.word_wrap = False
    p = tf.paragraphs[0]
    p.text = str(i + 1)
    p.font.size = Pt(16)
    p.font.bold = True
    p.font.color.rgb = DARK_BG
    p.alignment = PP_ALIGN.CENTER
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE

    add_text_box(slide, Inches(1.6), y, Inches(10.5), Inches(0.6),
                 item, font_size=17, color=WHITE)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 5 – Product Strategy
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.9),
             "Product Strategy: Building for the Vertical",
             font_size=36, color=ACCENT_BLUE, bold=True)

# Three pillars
pillars = [
    ("1. Deep Domain Modeling", [
        "Map the industry value chain end to end",
        "Build domain-specific data models & objects",
        "Use industry terminology in UX (not generic labels)",
        "Embed compliance rules into the product itself",
    ], ACCENT_BLUE),
    ("2. Workflow-First Design", [
        "Shadow real users in their environment",
        "Digitize existing paper / manual processes first",
        "Automate the boring, error-prone steps",
        "Design for the least technical user persona",
    ], ACCENT_GREEN),
    ("3. Platform & Ecosystem", [
        "Integrate with industry-specific tools & ERPs",
        "Build APIs for partner / channel integrations",
        "Create a data layer for analytics & benchmarking",
        "Plan for embedded fintech (payments, lending, insurance)",
    ], ACCENT_PURPLE),
]

for i, (title, items, color) in enumerate(pillars):
    x = Inches(0.6 + i * 4.15)
    add_shape(slide, x, Inches(1.7), Inches(3.9), Inches(5.2), CARD_BG, color)
    add_text_box(slide, x + Inches(0.3), Inches(1.9), Inches(3.3), Inches(0.7),
                 title, font_size=20, color=color, bold=True)
    add_bullet_slide_content(slide, items,
                             x + Inches(0.3), Inches(2.7), Inches(3.3), Inches(3.8), font_size=15)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 6 – Go-to-Market Playbook
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.9),
             "Go-to-Market Playbook for Vertical SaaS",
             font_size=36, color=ACCENT_GREEN, bold=True)

gtm_items = [
    ("Land with a Wedge", "Start with ONE critical pain point. Win a single workflow before expanding to a full suite. Early adopters become your evangelists.", ACCENT_BLUE),
    ("Industry-Native Sales", "Hire reps from the industry, not just SaaS. Attend vertical trade shows. Publish thought leadership in trade publications.", ACCENT_GREEN),
    ("Customer-Led Growth", "Case studies, ROI calculators, peer referrals. In tight verticals, reputation spreads fast -- both good and bad.", ACCENT_PURPLE),
    ("Expand Revenue Per Account", "Layer on modules, payments, data products, and professional services. Vertical SaaS can often 3-5x initial ACV over time.", ACCENT_ORANGE),
]

for i, (title, desc, color) in enumerate(gtm_items):
    y = Inches(1.7 + i * 1.4)
    add_shape(slide, Inches(0.8), y, Inches(11.5), Inches(1.2), CARD_BG, color)
    add_text_box(slide, Inches(1.2), y + Inches(0.1), Inches(3), Inches(0.6),
                 title, font_size=20, color=color, bold=True)
    add_text_box(slide, Inches(1.2), y + Inches(0.6), Inches(10.5), Inches(0.5),
                 desc, font_size=15, color=LIGHT_GRAY)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 7 – Key Metrics & Benchmarks
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.9),
             "Key Metrics & Benchmarks for Vertical SaaS",
             font_size=36, color=ACCENT_PURPLE, bold=True)

metrics = [
    ("Net Revenue\nRetention", "120-140%+", "Best-in-class verticals expand\nwithin accounts aggressively", ACCENT_BLUE),
    ("Gross\nMargin", "70-80%+", "Can be lower if services-heavy;\naim for software-like margins", ACCENT_GREEN),
    ("CAC Payback", "12-18 mo", "Efficient GTM due to\nconcentrated buyer base", ACCENT_PURPLE),
    ("Logo\nRetention", "90-95%+", "High switching costs\nkeep churn low", ACCENT_ORANGE),
    ("Market\nPenetration", "10-30%", "Realistic ceiling within a\nsingle vertical segment", ACCENT_BLUE),
]

for i, (label, value, desc, color) in enumerate(metrics):
    x = Inches(0.4 + i * 2.55)
    add_shape(slide, x, Inches(1.8), Inches(2.35), Inches(4.8), CARD_BG, color)
    add_text_box(slide, x + Inches(0.15), Inches(2.0), Inches(2.05), Inches(0.9),
                 label, font_size=16, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, x + Inches(0.15), Inches(3.0), Inches(2.05), Inches(0.9),
                 value, font_size=36, color=color, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, x + Inches(0.15), Inches(4.1), Inches(2.05), Inches(1.2),
                 desc, font_size=14, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 8 – Common Pitfalls
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.9),
             "Common Pitfalls to Avoid",
             font_size=36, color=ACCENT_ORANGE, bold=True)

pitfalls = [
    ("Going Too Broad Too Soon", "Resist the temptation to serve adjacent verticals before dominating your first one. Depth beats breadth early on."),
    ("Ignoring Services Revenue", "Many verticals expect onboarding, training, and customization. Build a services motion -- don't treat it as a distraction."),
    ("Underestimating Compliance", "Regulatory requirements aren't optional. Bake compliance (HIPAA, PCI, SOC 2, industry regs) into the product from day one."),
    ("Building for Power Users Only", "Your ICP may include non-technical users. If the UX requires training, you'll lose deals to simpler (even inferior) tools."),
    ("Neglecting Data & Analytics", "Your aggregated data is a strategic asset. Industry benchmarking and insights can become a standalone revenue stream."),
    ("Over-Customizing Per Client", "Custom work that only one customer needs erodes margins. Build configurable, not custom. Say no to one-off requests."),
]

for i, (title, desc) in enumerate(pitfalls):
    col = i % 3
    row = i // 3
    x = Inches(0.6 + col * 4.1)
    y = Inches(1.7 + row * 2.8)
    add_shape(slide, x, y, Inches(3.85), Inches(2.5), CARD_BG, ACCENT_ORANGE)
    add_text_box(slide, x + Inches(0.25), y + Inches(0.2), Inches(3.35), Inches(0.6),
                 title, font_size=18, color=ACCENT_ORANGE, bold=True)
    add_text_box(slide, x + Inches(0.25), y + Inches(0.85), Inches(3.35), Inches(1.4),
                 desc, font_size=14, color=LIGHT_GRAY)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 9 – Scaling & Expansion Framework
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)

add_text_box(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.9),
             "Scaling Your Vertical SaaS: The Expansion Framework",
             font_size=36, color=ACCENT_BLUE, bold=True)

phases = [
    ("Phase 1\nWedge", "Single killer workflow\n5-20 design partners\nProve ROI in one use case\nAchieve product-market fit", ACCENT_BLUE),
    ("Phase 2\nSuite", "Add adjacent workflows\nBecome system of record\nLaunch self-serve onboarding\nHit $1-5M ARR", ACCENT_GREEN),
    ("Phase 3\nPlatform", "APIs & integrations layer\nEmbedded fintech\nMarketplace / app store\nHit $10-30M ARR", ACCENT_PURPLE),
    ("Phase 4\nEcosystem", "Data & benchmarking products\nAdjacent vertical expansion\nM&A bolt-ons\nScale to $50M+ ARR", ACCENT_ORANGE),
]

for i, (phase, items, color) in enumerate(phases):
    x = Inches(0.5 + i * 3.15)
    add_shape(slide, x, Inches(1.8), Inches(2.9), Inches(5), CARD_BG, color)
    add_text_box(slide, x + Inches(0.2), Inches(2.0), Inches(2.5), Inches(1),
                 phase, font_size=22, color=color, bold=True, alignment=PP_ALIGN.CENTER)
    add_text_box(slide, x + Inches(0.2), Inches(3.2), Inches(2.5), Inches(3.2),
                 items, font_size=15, color=LIGHT_GRAY, alignment=PP_ALIGN.CENTER)

    # Arrow between phases
    if i < 3:
        arrow = slide.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                        x + Inches(3.0), Inches(4.0), Inches(0.25), Inches(0.35))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = MEDIUM_GRAY
        arrow.line.fill.background()


# ═══════════════════════════════════════════════════════════════════
# SLIDE 10 – Key Takeaways
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg(slide, DARK_BG)

add_shape(slide, Inches(0), Inches(0), Inches(13.333), Inches(0.08), ACCENT_GREEN)

add_text_box(slide, Inches(1), Inches(0.8), Inches(11), Inches(1),
             "Key Takeaways",
             font_size=40, color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)

takeaways = [
    "Go deep before going wide -- dominate one vertical before expanding",
    "Build with domain experts, not just engineers -- hire from the industry",
    "Compliance and workflows ARE the product -- not afterthoughts",
    "Your GTM is your moat -- industry relationships compound over time",
    "Layer revenue streams -- software, payments, data, services",
    "Vertical SaaS businesses command premium valuations (10-20x+ ARR) due to stickiness and expansion",
]

for i, item in enumerate(takeaways):
    y = Inches(2.0 + i * 0.82)
    # Checkmark box
    check = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(1.5), y + Inches(0.05), Inches(0.4), Inches(0.4))
    check.fill.solid()
    check.fill.fore_color.rgb = ACCENT_GREEN
    check.line.fill.background()
    tf = check.text_frame
    p = tf.paragraphs[0]
    p.text = "\u2713"
    p.font.size = Pt(18)
    p.font.bold = True
    p.font.color.rgb = DARK_BG
    p.alignment = PP_ALIGN.CENTER

    add_text_box(slide, Inches(2.2), y, Inches(9.5), Inches(0.6),
                 item, font_size=19, color=WHITE)


# ── Save ──
output_path = "/home/user/Claude-Code-April/Vertical_Approach_SaaS_Guide.pptx"
prs.save(output_path)
print(f"Presentation saved to: {output_path}")
