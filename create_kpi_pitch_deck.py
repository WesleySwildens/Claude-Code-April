#!/usr/bin/env python3
"""
Generate KPI-Driven Operations Pitch Deck using SD Worx template.
Enhanced visual design while maintaining SD Worx look & feel.
"""

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
import copy

TEMPLATE = "/home/user/Claude-Code-April/2602 - SD Worx PowerPoint GOED .pptx"
OUTPUT = "/home/user/Claude-Code-April/KPI_Driven_Operations_Pitch_Deck.pptx"

# SD Worx theme colors
DK1 = RGBColor(0x30, 0x36, 0x42)       # Dark navy
ACCENT2 = RGBColor(0x43, 0x8A, 0xB5)   # Blue
ACCENT3 = RGBColor(0x75, 0x82, 0x9B)   # Slate gray
ACCENT4 = RGBColor(0xE4, 0xE6, 0xEC)   # Light gray
ACCENT5 = RGBColor(0xFF, 0x4E, 0x0F)   # SD Worx Orange
ACCENT6 = RGBColor(0x7A, 0x00, 0x51)   # SD Worx Purple/Magenta
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
BLACK = RGBColor(0x00, 0x00, 0x00)
LIGHT_BLUE = RGBColor(0xDB, 0xE9, 0xF1)  # Tinted blue background
DARK_BLUE = RGBColor(0x2A, 0x5C, 0x7A)   # Darker accent blue

# Slide dimensions (standard widescreen)
SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# Load template
prs = Presentation(TEMPLATE)

layout_map = {}
for layout in prs.slide_layouts:
    layout_map[layout.name] = layout

# Remove all existing slides
while len(prs.slides) > 0:
    rId = prs.slides._sldIdLst[0].get(
        '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id'
    )
    prs.part.drop_rel(rId)
    prs.slides._sldIdLst.remove(prs.slides._sldIdLst[0])


# ── Helper functions ──

def get_layout(name):
    if name in layout_map:
        return layout_map[name]
    for k, v in layout_map.items():
        if name.lower() in k.lower():
            return v
    return prs.slide_layouts[0]


def set_placeholder_text(slide, idx, text, font_size=None, bold=None, color=None):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            ph.text = text
            if ph.text_frame.paragraphs:
                para = ph.text_frame.paragraphs[0]
                for run in para.runs:
                    if font_size:
                        run.font.size = Pt(font_size)
                    if bold is not None:
                        run.font.bold = bold
                    if color:
                        run.font.color.rgb = color
            return ph
    return None


def set_placeholder_bullets(slide, idx, items, font_size=14, color=None):
    for ph in slide.placeholders:
        if ph.placeholder_format.idx == idx:
            tf = ph.text_frame
            tf.clear()
            for i, item in enumerate(items):
                p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
                p.text = item
                p.font.size = Pt(font_size)
                if color:
                    p.font.color.rgb = color
                p.space_after = Pt(6)
            return ph
    return None


def add_shape(slide, shape_type, left, top, width, height, fill_color=None, line_color=None):
    shape = slide.shapes.add_shape(shape_type, left, top, width, height)
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    else:
        shape.fill.background()
    if line_color:
        shape.line.color.rgb = line_color
    else:
        shape.line.fill.background()
    return shape


def add_textbox(slide, left, top, width, height, text, font_size=14,
                bold=False, color=DK1, alignment=PP_ALIGN.LEFT, font_name="Roboto"):
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.alignment = alignment
    run = p.add_run()
    run.text = text
    run.font.size = Pt(font_size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font_name
    return txBox


def add_rich_textbox(slide, left, top, width, height, lines, alignment=PP_ALIGN.LEFT):
    """Add a textbox with multiple lines, each with its own formatting.
    lines: list of dicts with keys: text, font_size, bold, color, space_after, space_before
    """
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, line_info in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = alignment
        run = p.add_run()
        run.text = line_info.get("text", "")
        run.font.size = Pt(line_info.get("font_size", 14))
        run.font.bold = line_info.get("bold", False)
        run.font.color.rgb = line_info.get("color", DK1)
        run.font.name = line_info.get("font_name", "Roboto")
        if "space_after" in line_info:
            p.space_after = Pt(line_info["space_after"])
        if "space_before" in line_info:
            p.space_before = Pt(line_info["space_before"])
    return txBox


def add_bullet_list(slide, left, top, width, height, items, font_size=14,
                    color=DK1, bullet_color=ACCENT5, spacing=8):
    """Add a bulleted list with orange bullet markers."""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True
    for i, item in enumerate(items):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.alignment = PP_ALIGN.LEFT
        # Add bullet character
        bullet_run = p.add_run()
        bullet_run.text = "\u25CF  "  # filled circle
        bullet_run.font.size = Pt(font_size - 2)
        bullet_run.font.color.rgb = bullet_color
        bullet_run.font.name = "Roboto"
        # Add text
        text_run = p.add_run()
        text_run.text = item
        text_run.font.size = Pt(font_size)
        text_run.font.color.rgb = color
        text_run.font.name = "Roboto"
        p.space_after = Pt(spacing)
    return txBox


def add_card(slide, left, top, width, height, title, items, title_color=WHITE,
             bg_color=DK1, text_color=WHITE, title_size=16, text_size=13,
             bullet_color=ACCENT5, corner_radius=None):
    """Add a card-style box with title and bullet items."""
    # Background rounded rectangle
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg_color
    shape.line.fill.background()
    # Adjust corner rounding
    if corner_radius is not None:
        shape.adjustments[0] = corner_radius

    # Title
    add_textbox(slide, left + Inches(0.3), top + Inches(0.25), width - Inches(0.6), Inches(0.5),
                title, font_size=title_size, bold=True, color=title_color)

    # Separator line
    line = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  left + Inches(0.3), top + Inches(0.75),
                                  Inches(1.2), Pt(3))
    line.fill.solid()
    line.fill.fore_color.rgb = bullet_color
    line.line.fill.background()

    # Bullet items
    if items:
        add_bullet_list(slide, left + Inches(0.3), top + Inches(0.95),
                        width - Inches(0.6), height - Inches(1.2),
                        items, font_size=text_size, color=text_color,
                        bullet_color=bullet_color, spacing=6)


def add_blank_slide():
    """Add a blank slide using a minimal layout."""
    # Try to find a blank or minimal layout
    for name in ['Blank', 'blank', 'Leeg']:
        if name in layout_map:
            return prs.slides.add_slide(layout_map[name])
    # Fallback: use the first layout
    return prs.slides.add_slide(prs.slide_layouts[0])


# ═══════════════════════════════════════════════════════════════════
# SLIDE 1 – Titel (Cover)
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(get_layout('Cover text - option 1'))
set_placeholder_text(slide, 0, "Driving Value Through\nKPI-Driven Operations")
set_placeholder_text(slide, 10, "Van inzicht naar actie per afdeling")
set_placeholder_text(slide, 12, "")
set_placeholder_text(slide, 13, "2026")


# ═══════════════════════════════════════════════════════════════════
# SLIDE 2 – De kernboodschap
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(get_layout('Quote / statement - gradient option 1'))
set_placeholder_text(slide, 0,
    "Zonder duidelijke KPI's → geen focus\n"
    "Zonder focus → geen voorspelbare resultaten\n"
    "Zonder resultaten → geen schaalbare groei")
set_placeholder_text(slide, 10, "Elk team moet sturen op meetbare waarde")


# ═══════════════════════════════════════════════════════════════════
# SLIDE 3 – Ons uitgangspunt
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(get_layout('Content - 2 blocks'))
set_placeholder_text(slide, 0, "Eén manier van werken, meerdere teams")
set_placeholder_text(slide, 18, "Ons uitgangspunt voor KPI-gedreven operatie")

set_placeholder_text(slide, 19, "Elk team heeft:", bold=True)
set_placeholder_bullets(slide, 1, [
    "Een duidelijke doelstelling",
    "Een set KPI's",
    "Inzicht in bijdrage aan totaalresultaat",
], font_size=14)

set_placeholder_text(slide, 22, "KPI's zijn:", bold=True)
set_placeholder_bullets(slide, 23, [
    "Beïnvloedbaar door het team",
    "Meetbaar en concreet",
    "Gekoppeld aan waarde",
], font_size=14)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 4 – Top-down + Bottom-up
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(get_layout('Content - 2 blocks'))
set_placeholder_text(slide, 0, "Balans tussen centrale sturing en team ownership")
set_placeholder_text(slide, 18, "Top-down + Bottom-up = één KPI-structuur")

set_placeholder_text(slide, 19, "Top-down KPI's", bold=True)
set_placeholder_bullets(slide, 1, [
    "Strategische doelen organisatie",
    "Financiële performance",
    "Klantimpact",
], font_size=14)

set_placeholder_text(slide, 22, "Bottom-up KPI's", bold=True)
set_placeholder_bullets(slide, 23, [
    "Team-specifieke metrics",
    "Operationele drivers",
    "Dagelijkse sturing",
], font_size=14)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 5 – Van strategie naar operatie (visual flow)
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(get_layout('Chapter text - option 1'))
set_placeholder_text(slide, 11, "")
set_placeholder_text(slide, 0, "Van strategie\nnaar operatie")
set_placeholder_text(slide, 10, "Elke KPI moet leiden tot concreet gedrag")

# We'll add the visual flow on a separate content slide
slide2 = add_blank_slide()

# Title bar at top
add_textbox(slide2, Inches(0.8), Inches(0.5), Inches(11), Inches(0.7),
            "Van strategie naar operatie", font_size=28, bold=True, color=DK1)

# Subtitle
add_textbox(slide2, Inches(0.8), Inches(1.15), Inches(11), Inches(0.5),
            "Hoe strategische doelen vertalen naar dagelijkse acties", font_size=14, color=ACCENT3)

# Flow boxes - horizontal cascade
flow_items = [
    ("Strategie", ACCENT6, "Organisatie-\ndoelstellingen"),
    ("Centrale KPI's", DK1, "Financieel,\nklant, groei"),
    ("Team KPI's", ACCENT2, "Per afdeling\nmeetbaar"),
    ("Dagelijkse Acties", ACCENT5, "Concreet\ngedrag"),
]

start_x = Inches(0.9)
box_w = Inches(2.6)
box_h = Inches(2.8)
gap = Inches(0.55)
arrow_w = Inches(0.35)
y_pos = Inches(2.3)

for i, (title, bg, desc) in enumerate(flow_items):
    x = start_x + i * (box_w + gap + arrow_w)

    # Main box
    shape = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, y_pos, box_w, box_h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg
    shape.line.fill.background()
    shape.adjustments[0] = 0.05

    # Title in box
    add_textbox(slide2, x + Inches(0.2), y_pos + Inches(0.4), box_w - Inches(0.4), Inches(0.6),
                title, font_size=20, bold=True, color=WHITE, alignment=PP_ALIGN.CENTER)

    # Divider line
    div = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                  x + Inches(0.5), y_pos + Inches(1.15),
                                  box_w - Inches(1.0), Pt(2))
    div.fill.solid()
    div.fill.fore_color.rgb = WHITE
    div.line.fill.background()

    # Description
    add_textbox(slide2, x + Inches(0.2), y_pos + Inches(1.4), box_w - Inches(0.4), Inches(1.2),
                desc, font_size=15, color=WHITE, alignment=PP_ALIGN.CENTER)

    # Arrow between boxes
    if i < len(flow_items) - 1:
        arrow_x = x + box_w + Inches(0.08)
        arrow_y = y_pos + box_h / 2 - Inches(0.2)
        arrow = slide2.shapes.add_shape(MSO_SHAPE.RIGHT_ARROW,
                                        arrow_x, arrow_y, arrow_w, Inches(0.4))
        arrow.fill.solid()
        arrow.fill.fore_color.rgb = ACCENT3
        arrow.line.fill.background()

# Bottom callout
callout = slide2.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                  Inches(3.5), Inches(5.6), Inches(6.3), Inches(0.8))
callout.fill.solid()
callout.fill.fore_color.rgb = ACCENT4
callout.line.fill.background()
callout.adjustments[0] = 0.3

add_textbox(slide2, Inches(3.7), Inches(5.7), Inches(5.9), Inches(0.6),
            "\U0001F449 Elke KPI moet leiden tot concreet gedrag",
            font_size=16, bold=True, color=DK1, alignment=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 6 – Recurring Services
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(get_layout('Content - 3 blocks'))
set_placeholder_text(slide, 0, "Recurring Services – sturen op klantwaarde en efficiency")
set_placeholder_text(slide, 18, "KPI's en focus voor het Recurring Services team")

set_placeholder_text(slide, 19, "NPS", bold=True)
set_placeholder_bullets(slide, 1, [
    "Klanttevredenheid",
    "Klantbehoud versterken",
    "Signalen vroeg oppakken",
], font_size=13)

set_placeholder_text(slide, 28, "eNPS", bold=True)
set_placeholder_bullets(slide, 29, [
    "Medewerkerbetrokkenheid",
    "Teamvitaliteit meten",
    "Retentie waarborgen",
], font_size=13)

set_placeholder_text(slide, 25, "Contribution / Cost to Serve", bold=True)
set_placeholder_bullets(slide, 30, [
    "Efficiënte service delivery",
    "Schaalbaarheid",
    "Waarde per klant optimaliseren",
], font_size=13)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 7 – Consulting
# ═══════════════════════════════════════════════════════════════════
slide = add_blank_slide()

# Title
add_textbox(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.7),
            "Consulting – sturen op waardecreatie", font_size=28, bold=True, color=DK1)

add_textbox(slide, Inches(0.8), Inches(1.15), Inches(11), Inches(0.5),
            "Van uren schrijven naar aantoonbare klantwaarde", font_size=14, color=ACCENT3)

# KPI cards - 2x2 grid
card_data = [
    ("tNPS", ["Initieel: onboarding/implementatie", "Klanttevredenheid bij trajectory", "Kwaliteit van oplevering"], ACCENT2),
    ("eNPS", ["Medewerkerbetrokkenheid", "Teamdynamiek en groei", "Consultants als ambassadeurs"], DK1),
    ("Revenue per FTE / Billability", ["Productiviteit per consultant", "Effectieve inzet van capaciteit", "Financiële gezondheid team"], ACCENT6),
    ("Proven Value Time", ["(Toekomst) Tijd besteed met/voor klant", "Aantoonbare waardecreatie", "Meetbare klantimpact"], ACCENT5),
]

cards_start_x = Inches(0.8)
cards_start_y = Inches(1.9)
card_w = Inches(5.8)
card_h = Inches(2.4)
card_gap_x = Inches(0.5)
card_gap_y = Inches(0.35)

for i, (title, items, accent) in enumerate(card_data):
    col = i % 2
    row = i // 2
    x = cards_start_x + col * (card_w + card_gap_x)
    y = cards_start_y + row * (card_h + card_gap_y)

    add_card(slide, x, y, card_w, card_h, title, items,
             title_color=WHITE, bg_color=accent, text_color=WHITE,
             title_size=18, text_size=13, bullet_color=WHITE,
             corner_radius=0.04)

# Bottom banner
banner = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                Inches(2.5), Inches(6.85), Inches(8.3), Inches(0.55))
banner.fill.solid()
banner.fill.fore_color.rgb = ACCENT4
banner.line.fill.background()
banner.adjustments[0] = 0.4

add_textbox(slide, Inches(2.7), Inches(6.9), Inches(7.9), Inches(0.45),
            "\U0001F449 Belangrijke shift: van uren schrijven → naar aantoonbare klantwaarde",
            font_size=14, bold=True, color=DK1, alignment=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 8 – KC (Knowledge Center)
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(get_layout('Content - 3 blocks'))
set_placeholder_text(slide, 0, "KC – sturen op kwaliteit en adoptie")
set_placeholder_text(slide, 18, "KPI's en focus voor het Knowledge Center team")

set_placeholder_text(slide, 19, "Kennis & Kwaliteit", bold=True)
set_placeholder_bullets(slide, 1, [
    "Kennisborging",
    "Kwaliteitsstandaarden",
    "Kennisdeling bevorderen",
], font_size=13)

set_placeholder_text(slide, 28, "Compliance & Control", bold=True)
set_placeholder_bullets(slide, 29, [
    "Risicobeheersing",
    "Naleving van regelgeving",
    "Audit-readiness",
], font_size=13)

set_placeholder_text(slide, 25, "Training & Adoption", bold=True)
set_placeholder_bullets(slide, 30, [
    "Gebruik van oplossingen",
    "Adoptiegraad verhogen",
    "Continue ontwikkeling",
], font_size=13)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 9 – Van KPI naar actie
# ═══════════════════════════════════════════════════════════════════
slide = prs.slides.add_slide(get_layout('Content - 4 blocks option 1'))
set_placeholder_text(slide, 0, "Van KPI naar actie")
set_placeholder_text(slide, 18, "Geen KPI zonder actieplan")

set_placeholder_text(slide, 33, "1")
set_placeholder_text(slide, 19, "KPI's definiëren", bold=True)
set_placeholder_bullets(slide, 29, [
    "Kies de juiste metrics",
    "Koppel aan teamdoelstelling",
    "Maak ze beïnvloedbaar",
], font_size=11)

set_placeholder_text(slide, 34, "2")
set_placeholder_text(slide, 28, "Targets bepalen", bold=True)
set_placeholder_bullets(slide, 30, [
    "Realistisch maar ambitieus",
    "Gebaseerd op baseline",
    "Tijdsgebonden",
], font_size=11)

set_placeholder_text(slide, 35, "3")
set_placeholder_text(slide, 25, "Acties koppelen", bold=True)
set_placeholder_bullets(slide, 17, [
    "Elke KPI → concrete actie",
    "Verantwoordelijke toewijzen",
    "Middelen alloceren",
], font_size=11)

set_placeholder_text(slide, 36, "4")
set_placeholder_text(slide, 32, "Wekelijkse opvolging", bold=True)
set_placeholder_bullets(slide, 31, [
    "Voortgang monitoren",
    "Bijsturen waar nodig",
    "Successen vieren",
], font_size=11)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 10 – Ownership bij de teams
# ═══════════════════════════════════════════════════════════════════
slide = add_blank_slide()

# Title
add_textbox(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.7),
            "Ownership bij de teams", font_size=28, bold=True, color=DK1)

add_textbox(slide, Inches(0.8), Inches(1.15), Inches(11), Inches(0.5),
            "Wat verwachten we van elk team?", font_size=14, color=ACCENT3)

# Four ownership pillars as vertical cards
pillars = [
    ("Begrijp je KPI's", "Weet wat je meet\nen waarom het\nbelangrijk is", "01"),
    ("Weet hoe je ze\nbeïnvloedt", "Ken de hefbomen\ndie je als team\nkunt gebruiken", "02"),
    ("Stuur actief bij", "Wacht niet af maar\nneem initiatief\nom bij te sturen", "03"),
    ("Maak impact\nzichtbaar", "Laat resultaten\nzien en deel\nsuccessen", "04"),
]

pillar_w = Inches(2.7)
pillar_h = Inches(4.5)
pillar_gap = Inches(0.45)
pillar_start_x = Inches(0.8)
pillar_y = Inches(1.9)

colors = [ACCENT6, DK1, ACCENT2, ACCENT5]

for i, (title, desc, num) in enumerate(pillars):
    x = pillar_start_x + i * (pillar_w + pillar_gap)
    bg = colors[i]

    # Card background
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, x, pillar_y, pillar_w, pillar_h)
    shape.fill.solid()
    shape.fill.fore_color.rgb = bg
    shape.line.fill.background()
    shape.adjustments[0] = 0.04

    # Number circle
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                    x + Inches(0.9), pillar_y + Inches(0.4),
                                    Inches(0.9), Inches(0.9))
    circle.fill.solid()
    circle.fill.fore_color.rgb = WHITE
    circle.line.fill.background()

    add_textbox(slide, x + Inches(0.9), pillar_y + Inches(0.5),
                Inches(0.9), Inches(0.7),
                num, font_size=24, bold=True, color=bg, alignment=PP_ALIGN.CENTER)

    # Title
    add_textbox(slide, x + Inches(0.2), pillar_y + Inches(1.6),
                pillar_w - Inches(0.4), Inches(1.0),
                title, font_size=17, bold=True, color=WHITE, alignment=PP_ALIGN.CENTER)

    # Description
    add_textbox(slide, x + Inches(0.2), pillar_y + Inches(2.7),
                pillar_w - Inches(0.4), Inches(1.4),
                desc, font_size=14, color=WHITE, alignment=PP_ALIGN.CENTER)


# ═══════════════════════════════════════════════════════════════════
# SLIDE 11 – Next steps
# ═══════════════════════════════════════════════════════════════════
slide = add_blank_slide()

# Title
add_textbox(slide, Inches(0.8), Inches(0.5), Inches(11), Inches(0.7),
            "Van inzicht naar implementatie", font_size=28, bold=True, color=DK1)

add_textbox(slide, Inches(0.8), Inches(1.15), Inches(11), Inches(0.5),
            "Next steps: workshops per afdeling", font_size=14, color=ACCENT3)

# Left section: Workshop steps
left_x = Inches(0.8)
left_y = Inches(2.0)
section_w = Inches(5.8)

add_textbox(slide, left_x, left_y, section_w, Inches(0.5),
            "Workshopreeks per afdeling", font_size=18, bold=True, color=DK1)

# Step items with numbered circles
steps = [
    "Bepalen van teamwaarde",
    "Vertalen naar KPI's",
    "Koppelen aan concrete acties",
    "Inrichten van ritme (weekly/monthly)",
]

for i, step in enumerate(steps):
    step_y = left_y + Inches(0.7) + i * Inches(0.75)

    # Number circle
    circle = slide.shapes.add_shape(MSO_SHAPE.OVAL,
                                    left_x, step_y, Inches(0.5), Inches(0.5))
    circle.fill.solid()
    circle.fill.fore_color.rgb = ACCENT5
    circle.line.fill.background()

    add_textbox(slide, left_x, step_y + Inches(0.05),
                Inches(0.5), Inches(0.4),
                str(i + 1), font_size=16, bold=True, color=WHITE, alignment=PP_ALIGN.CENTER)

    # Step text
    add_textbox(slide, left_x + Inches(0.7), step_y + Inches(0.07),
                section_w - Inches(0.7), Inches(0.4),
                step, font_size=16, color=DK1)

# Right section: Output card
right_x = Inches(7.2)
right_y = Inches(2.0)
output_w = Inches(5.3)
output_h = Inches(4.2)

# Output card
shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE, right_x, right_y, output_w, output_h)
shape.fill.solid()
shape.fill.fore_color.rgb = DK1
shape.line.fill.background()
shape.adjustments[0] = 0.04

add_textbox(slide, right_x + Inches(0.4), right_y + Inches(0.35),
            output_w - Inches(0.8), Inches(0.5),
            "Output per team", font_size=20, bold=True, color=WHITE)

# Separator
sep = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                              right_x + Inches(0.4), right_y + Inches(0.95),
                              Inches(1.5), Pt(3))
sep.fill.solid()
sep.fill.fore_color.rgb = ACCENT5
sep.line.fill.background()

outputs = [
    "Heldere KPI-set",
    "Concrete targets",
    "Actieplan",
]

add_bullet_list(slide, right_x + Inches(0.4), right_y + Inches(1.3),
                output_w - Inches(0.8), Inches(2.5),
                outputs, font_size=18, color=WHITE, bullet_color=ACCENT5, spacing=14)

# Bottom bar with gradient feel
bottom_bar = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE,
                                    Inches(0.8), Inches(6.6), Inches(11.7), Inches(0.7))
bottom_bar.fill.solid()
bottom_bar.fill.fore_color.rgb = ACCENT4
bottom_bar.line.fill.background()
bottom_bar.adjustments[0] = 0.3

add_textbox(slide, Inches(1.0), Inches(6.68), Inches(11.3), Inches(0.5),
            "\U0001F449 Laten we samen beginnen – elk team, elke KPI, elke actie telt",
            font_size=15, bold=True, color=DK1, alignment=PP_ALIGN.CENTER)


# ── Save ──
prs.save(OUTPUT)
print(f"Presentation saved to: {OUTPUT}")
print(f"Total slides: {len(prs.slides)}")
