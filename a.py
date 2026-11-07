from pptx import Presentation
from pptx.util import Inches, Pt, Emu, Cm
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml.ns import qn
from lxml import etree
import copy

# ── Helper Functions ──────────────────────────────────────────────

def hex_to_rgb(hex_color):
    """Convert hex color to RGBColor"""
    hex_color = hex_color.lstrip('#')
    return RGBColor(int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16))

def add_rounded_rect(slide, left, top, width, height, fill_color=None, border_color=None,
                     border_width=None, corner_radius=None, opacity=None, shadow=False):
    """Add a rounded rectangle shape"""
    shape = slide.shapes.add_shape(
        MSO_SHAPE.ROUNDED_RECTANGLE, left, top, width, height
    )
    shape.line.fill.background()

    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
        if opacity is not None:
            # Set transparency
            solidFill = shape.fill._fill
            srgbClr = solidFill.find(qn('a:solidFill')).find(qn('a:srgbClr'))
            if srgbClr is None:
                solidFill_elem = solidFill.find(qn('a:solidFill'))
                for child in solidFill_elem:
                    srgbClr = child
                    break
            if srgbClr is not None:
                alpha = etree.SubElement(srgbClr, qn('a:alpha'))
                alpha.set('val', str(int(opacity * 1000)))

    if border_color:
        shape.line.color.rgb = border_color
        shape.line.width = border_width if border_width else Pt(1)
    else:
        shape.line.fill.background()

    if shadow:
        add_shadow(shape)

    return shape

def add_shadow(shape):
    """Add shadow effect to shape"""
    spPr = shape._element.spPr
    effectLst = etree.SubElement(spPr, qn('a:effectLst'))
    outerShdw = etree.SubElement(effectLst, qn('a:outerShdw'))
    outerShdw.set('blurRad', '40000')
    outerShdw.set('dist', '20000')
    outerShdw.set('dir', '5400000')
    outerShdw.set('rotWithShape', '0')
    srgbClr = etree.SubElement(outerShdw, qn('a:srgbClr'))
    srgbClr.set('val', '000000')
    alpha = etree.SubElement(srgbClr, qn('a:alpha'))
    alpha.set('val', '20000')

def add_text_box(slide, left, top, width, height, text, font_size=12,
                 font_color=None, bold=False, alignment=PP_ALIGN.LEFT,
                 font_name='Segoe UI', word_wrap=True, italic=False):
    """Add a text box"""
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = word_wrap
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = font_size
    p.font.bold = bold
    p.font.italic = italic
    p.alignment = alignment
    if font_color:
        p.font.color.rgb = font_color
    if font_name:
        p.font.name = font_name
    return txBox

def add_multiline_text(slide, left, top, width, height, lines, font_name='Segoe UI'):
    """Add text box with multiple formatted lines.
    lines = [(text, font_size, font_color, bold, alignment, italic, spacing_before, spacing_after), ...]
    """
    txBox = slide.shapes.add_textbox(left, top, width, height)
    tf = txBox.text_frame
    tf.word_wrap = True

    for i, line_data in enumerate(lines):
        text = line_data[0]
        font_size = line_data[1] if len(line_data) > 1 else Pt(12)
        font_color = line_data[2] if len(line_data) > 2 else None
        bold = line_data[3] if len(line_data) > 3 else False
        alignment = line_data[4] if len(line_data) > 4 else PP_ALIGN.LEFT
        italic = line_data[5] if len(line_data) > 5 else False

        if i == 0:
            p = tf.paragraphs[0]
        else:
            p = tf.add_paragraph()

        p.text = text
        p.font.size = font_size
        p.font.bold = bold
        p.font.italic = italic
        p.alignment = alignment
        if font_color:
            p.font.color.rgb = font_color
        p.font.name = font_name

    return txBox

def add_circle(slide, left, top, size, fill_color=None, text='', font_size=Pt(10), font_color=None):
    """Add a circle shape with optional text"""
    shape = slide.shapes.add_shape(MSO_SHAPE.OVAL, left, top, size, size)
    if fill_color:
        shape.fill.solid()
        shape.fill.fore_color.rgb = fill_color
    shape.line.fill.background()

    if text:
        tf = shape.text_frame
        tf.word_wrap = True
        p = tf.paragraphs[0]
        p.text = text
        p.font.size = font_size
        if font_color:
            p.font.color.rgb = font_color
        p.font.bold = True
        p.alignment = PP_ALIGN.CENTER
        tf.paragraphs[0].alignment = PP_ALIGN.CENTER

    return shape

def add_gradient_rect(slide, left, top, width, height, color1, color2, angle=0):
    """Add a rectangle with gradient fill"""
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, left, top, width, height)
    shape.line.fill.background()

    fill = shape.fill
    fill.gradient()
    fill.gradient_stops[0].color.rgb = color1
    fill.gradient_stops[0].position = 0
    fill.gradient_stops[1].color.rgb = color2
    fill.gradient_stops[1].position = 1.0

    return shape

def set_slide_bg_gradient(slide, color1, color2):
    """Set slide background to gradient"""
    background = slide.background
    fill = background.fill
    fill.gradient()
    fill.gradient_stops[0].color.rgb = color1
    fill.gradient_stops[0].position = 0
    fill.gradient_stops[1].color.rgb = color2
    fill.gradient_stops[1].position = 1.0


# ── Color Palette (from CSS) ─────────────────────────────────────
DARK_BG = hex_to_rgb('0f172a')
DARK_BG_2 = hex_to_rgb('1e293b')
CARD_BG = hex_to_rgb('1e293b')
CARD_BORDER = hex_to_rgb('334155')
EMERALD = hex_to_rgb('10b981')
EMERALD_DARK = hex_to_rgb('059669')
EMERALD_LIGHT = hex_to_rgb('34d399')
TEAL = hex_to_rgb('14b8a6')
CYAN = hex_to_rgb('06b6d4')
PURPLE = hex_to_rgb('8b5cf6')
PINK = hex_to_rgb('ec4899')
AMBER = hex_to_rgb('f59e0b')
RED = hex_to_rgb('ef4444')
BLUE = hex_to_rgb('3b82f6')
SLATE_300 = hex_to_rgb('cbd5e1')
SLATE_400 = hex_to_rgb('94a3b8')
SLATE_500 = hex_to_rgb('64748b')
SLATE_600 = hex_to_rgb('475569')
SLATE_700 = hex_to_rgb('334155')
SLATE_800 = hex_to_rgb('1e293b')
WHITE = hex_to_rgb('ffffff')
WHITE_90 = hex_to_rgb('e2e8f0')
GREEN_900_BG = hex_to_rgb('064e3b')

# ── Create Presentation ──────────────────────────────────────────
prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

SLIDE_W = Inches(13.333)
SLIDE_H = Inches(7.5)

# ══════════════════════════════════════════════════════════════════
# SLIDE 1: COVER / TITLE SLIDE
# ══════════════════════════════════════════════════════════════════
slide1 = prs.slides.add_slide(prs.slide_layouts[6])  # Blank layout

# Background
set_slide_bg_gradient(slide1, DARK_BG, DARK_BG_2)

# Top decorative gradient bar
add_gradient_rect(slide1, Inches(0), Inches(0), SLIDE_W, Inches(0.06), EMERALD, CYAN)

# Main content area - centered card
card_w = Inches(8)
card_h = Inches(5)
card_left = (SLIDE_W - card_w) // 2
card_top = Inches(1.25)

# Glow effect background
glow = add_rounded_rect(slide1, card_left - Inches(0.3), card_top - Inches(0.3),
                        card_w + Inches(0.6), card_h + Inches(0.6),
                        fill_color=EMERALD, corner_radius=Inches(0.3))
# Set very low opacity for glow
spPr = glow._element.spPr
solidFill = spPr.find(qn('a:solidFill'))
if solidFill is None:
    fill_elem = glow.fill._fill.find(qn('a:solidFill'))
    if fill_elem is not None:
        for child in fill_elem:
            alpha = etree.SubElement(child, qn('a:alpha'))
            alpha.set('val', '8000')

# Main card
main_card = add_rounded_rect(slide1, card_left, card_top, card_w, card_h,
                             fill_color=CARD_BG, border_color=SLATE_700, border_width=Pt(1),
                             shadow=True)

# App icon - circle with emoji placeholder
icon_size = Inches(1.0)
icon_left = (SLIDE_W - icon_size) // 2
icon_top = card_top + Inches(0.5)
icon_circle = add_circle(slide1, icon_left, icon_top, icon_size,
                         fill_color=EMERALD, text='✈', font_size=Pt(36), font_color=WHITE)

# Title: TripSplit
add_text_box(slide1, card_left, icon_top + Inches(1.2), card_w, Inches(0.8),
             'TripSplit', font_size=Pt(44), font_color=WHITE, bold=True,
             alignment=PP_ALIGN.CENTER, font_name='Segoe UI')

# Subtitle
add_text_box(slide1, card_left, icon_top + Inches(2.0), card_w, Inches(0.5),
             'Smart Group Expense Splitting', font_size=Pt(20),
             font_color=EMERALD_LIGHT, bold=False, alignment=PP_ALIGN.CENTER)

# Description
add_text_box(slide1, card_left + Inches(0.5), icon_top + Inches(2.7), card_w - Inches(1), Inches(0.8),
             'Split expenses fairly among friends, track balances in real-time,\nand settle debts with minimal transactions.',
             font_size=Pt(14), font_color=SLATE_400, alignment=PP_ALIGN.CENTER)

# Bottom badge
badge_w = Inches(3)
badge_h = Inches(0.45)
badge_left = (SLIDE_W - badge_w) // 2
badge_top = card_top + card_h - Inches(0.8)
badge = add_rounded_rect(slide1, badge_left, badge_top, badge_w, badge_h,
                         fill_color=hex_to_rgb('064e3b'), border_color=EMERALD_DARK, border_width=Pt(1))
add_text_box(slide1, badge_left, badge_top + Inches(0.05), badge_w, badge_h,
             '📱 Mobile-First Progressive Web App', font_size=Pt(11),
             font_color=EMERALD_LIGHT, alignment=PP_ALIGN.CENTER)

# Bottom gradient bar
add_gradient_rect(slide1, Inches(0), SLIDE_H - Inches(0.06), SLIDE_W, Inches(0.06), EMERALD, CYAN)


# ══════════════════════════════════════════════════════════════════
# SLIDE 2: PROBLEM STATEMENT
# ══════════════════════════════════════════════════════════════════
slide2 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg_gradient(slide2, DARK_BG, DARK_BG_2)
add_gradient_rect(slide2, Inches(0), Inches(0), SLIDE_W, Inches(0.06), EMERALD, CYAN)
add_gradient_rect(slide2, Inches(0), SLIDE_H - Inches(0.06), SLIDE_W, Inches(0.06), EMERALD, CYAN)

# Section badge
badge = add_rounded_rect(slide2, Inches(0.8), Inches(0.5), Inches(2.2), Inches(0.4),
                         fill_color=GREEN_900_BG, border_color=EMERALD_DARK, border_width=Pt(1))
add_text_box(slide2, Inches(0.8), Inches(0.52), Inches(2.2), Inches(0.4),
             '⚠️  The Problem', font_size=Pt(11), font_color=EMERALD_LIGHT,
             alignment=PP_ALIGN.CENTER, bold=True)

# Title
add_text_box(slide2, Inches(0.8), Inches(1.1), Inches(11), Inches(0.7),
             'Why TripSplit?', font_size=Pt(36), font_color=WHITE, bold=True)

# Subtitle
add_text_box(slide2, Inches(0.8), Inches(1.8), Inches(11), Inches(0.5),
             'Group expense management is broken. Here\'s what travelers face today.',
             font_size=Pt(14), font_color=SLATE_400)

# Problem cards - 2x2 grid
problems = [
    ('😤', 'Manual Tracking Nightmare',
     'Spreadsheets, notes apps, and mental math lead to errors, forgotten expenses, and arguments among friends.',
     RED, hex_to_rgb('7f1d1d')),
    ('💸', 'Unfair Splits',
     'Not everyone participates in every expense. Current tools can\'t handle complex splitting scenarios properly.',
     AMBER, hex_to_rgb('78350f')),
    ('🔄', 'Settlement Chaos',
     'Without optimization, a group of 5 might need 10+ transactions to settle up. That\'s inefficient and confusing.',
     PURPLE, hex_to_rgb('4c1d95')),
    ('📵', 'No Offline Access',
     'Most apps require internet. Travelers in remote areas or abroad can\'t track expenses when they need it most.',
     CYAN, hex_to_rgb('164e63')),
]

card_w = Inches(5.5)
card_h = Inches(2.0)
start_x = Inches(0.8)
gap_x = Inches(0.5)
start_y = Inches(2.6)
gap_y = Inches(0.4)

for i, (emoji, title, desc, accent_color, bg_dark) in enumerate(problems):
    col = i % 2
    row = i // 2
    x = start_x + col * (card_w + gap_x)
    y = start_y + row * (card_h + gap_y)

    # Card background
    card = add_rounded_rect(slide2, x, y, card_w, card_h,
                           fill_color=CARD_BG, border_color=SLATE_700, border_width=Pt(1), shadow=True)

    # Accent line on left
    accent = slide2.shapes.add_shape(MSO_SHAPE.RECTANGLE, x, y + Inches(0.3), Inches(0.06), Inches(1.4))
    accent.fill.solid()
    accent.fill.fore_color.rgb = accent_color
    accent.line.fill.background()

    # Icon circle
    icon_bg = add_circle(slide2, x + Inches(0.3), y + Inches(0.3), Inches(0.6),
                        fill_color=bg_dark, text=emoji, font_size=Pt(20), font_color=WHITE)

    # Title
    add_text_box(slide2, x + Inches(1.1), y + Inches(0.3), card_w - Inches(1.4), Inches(0.4),
                title, font_size=Pt(16), font_color=WHITE, bold=True)

    # Description
    add_text_box(slide2, x + Inches(1.1), y + Inches(0.8), card_w - Inches(1.4), Inches(1.0),
                desc, font_size=Pt(11), font_color=SLATE_400)


# ══════════════════════════════════════════════════════════════════
# SLIDE 3: SOLUTION OVERVIEW
# ══════════════════════════════════════════════════════════════════
slide3 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg_gradient(slide3, DARK_BG, DARK_BG_2)
add_gradient_rect(slide3, Inches(0), Inches(0), SLIDE_W, Inches(0.06), EMERALD, CYAN)
add_gradient_rect(slide3, Inches(0), SLIDE_H - Inches(0.06), SLIDE_W, Inches(0.06), EMERALD, CYAN)

# Section badge
badge = add_rounded_rect(slide3, Inches(0.8), Inches(0.5), Inches(2.2), Inches(0.4),
                         fill_color=GREEN_900_BG, border_color=EMERALD_DARK, border_width=Pt(1))
add_text_box(slide3, Inches(0.8), Inches(0.52), Inches(2.2), Inches(0.4),
             '💡  The Solution', font_size=Pt(11), font_color=EMERALD_LIGHT,
             alignment=PP_ALIGN.CENTER, bold=True)

# Title
add_text_box(slide3, Inches(0.8), Inches(1.1), Inches(11), Inches(0.7),
             'How TripSplit Works', font_size=Pt(36), font_color=WHITE, bold=True)

add_text_box(slide3, Inches(0.8), Inches(1.8), Inches(11), Inches(0.5),
             'A seamless 4-step process from group creation to settlement.',
             font_size=Pt(14), font_color=SLATE_400)

# 4 step cards
steps = [
    ('1', '👥', 'Create Group', 'Start a trip group and add members. Share a simple code for others to join instantly.',
     EMERALD, hex_to_rgb('064e3b')),
    ('2', '💰', 'Add Expenses', 'Log expenses with flexible splitting — equal, exact amounts, or percentages. Tag categories too.',
     CYAN, hex_to_rgb('164e63')),
    ('3', '📊', 'Track Balances', 'Real-time balance calculations show who owes whom. Visual dashboard with spending analytics.',
     PURPLE, hex_to_rgb('4c1d95')),
    ('4', '✅', 'Settle Up', 'Optimized settlement algorithm minimizes transactions. Mark payments as complete with one tap.',
     AMBER, hex_to_rgb('78350f')),
]

step_card_w = Inches(2.7)
step_card_h = Inches(3.5)
total_w = 4 * step_card_w + 3 * Inches(0.35)
start_x = (SLIDE_W - total_w) // 2
start_y = Inches(2.5)

for i, (num, emoji, title, desc, accent, bg_dark) in enumerate(steps):
    x = start_x + i * (step_card_w + Inches(0.35))
    y = start_y

    # Card
    card = add_rounded_rect(slide3, x, y, step_card_w, step_card_h,
                           fill_color=CARD_BG, border_color=SLATE_700, border_width=Pt(1), shadow=True)

    # Top accent bar
    accent_bar = slide3.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                         x + Inches(0.2), y, step_card_w - Inches(0.4), Inches(0.05))
    accent_bar.fill.solid()
    accent_bar.fill.fore_color.rgb = accent
    accent_bar.line.fill.background()

    # Step number
    num_circle = add_circle(slide3, x + (step_card_w - Inches(0.5)) // 2, y + Inches(0.3),
                           Inches(0.5), fill_color=bg_dark, text=num, font_size=Pt(16), font_color=accent)

    # Icon
    add_text_box(slide3, x, y + Inches(1.0), step_card_w, Inches(0.5),
                emoji, font_size=Pt(32), alignment=PP_ALIGN.CENTER)

    # Title
    add_text_box(slide3, x + Inches(0.2), y + Inches(1.6), step_card_w - Inches(0.4), Inches(0.4),
                title, font_size=Pt(15), font_color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)

    # Description
    add_text_box(slide3, x + Inches(0.2), y + Inches(2.1), step_card_w - Inches(0.4), Inches(1.2),
                desc, font_size=Pt(10), font_color=SLATE_400, alignment=PP_ALIGN.CENTER)

    # Connector arrow between steps
    if i < 3:
        arrow_x = x + step_card_w + Inches(0.05)
        arrow_y = y + step_card_h // 2
        add_text_box(slide3, arrow_x, arrow_y - Inches(0.15), Inches(0.3), Inches(0.3),
                    '→', font_size=Pt(20), font_color=SLATE_500, alignment=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════
# SLIDE 4: KEY FEATURES
# ══════════════════════════════════════════════════════════════════
slide4 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg_gradient(slide4, DARK_BG, DARK_BG_2)
add_gradient_rect(slide4, Inches(0), Inches(0), SLIDE_W, Inches(0.06), EMERALD, CYAN)
add_gradient_rect(slide4, Inches(0), SLIDE_H - Inches(0.06), SLIDE_W, Inches(0.06), EMERALD, CYAN)

# Section badge
badge = add_rounded_rect(slide4, Inches(0.8), Inches(0.5), Inches(2.2), Inches(0.4),
                         fill_color=GREEN_900_BG, border_color=EMERALD_DARK, border_width=Pt(1))
add_text_box(slide4, Inches(0.8), Inches(0.52), Inches(2.2), Inches(0.4),
             '⭐  Features', font_size=Pt(11), font_color=EMERALD_LIGHT,
             alignment=PP_ALIGN.CENTER, bold=True)

# Title
add_text_box(slide4, Inches(0.8), Inches(1.1), Inches(11), Inches(0.7),
             'Key Features', font_size=Pt(36), font_color=WHITE, bold=True)

add_text_box(slide4, Inches(0.8), Inches(1.8), Inches(11), Inches(0.5),
             'Everything you need for hassle-free group expense management.',
             font_size=Pt(14), font_color=SLATE_400)

# Feature grid - 3x2
features = [
    ('📱', 'Offline-First PWA', 'Works without internet. All data syncs automatically when connection is restored. Install on any device.',
     EMERALD, hex_to_rgb('064e3b')),
    ('🔄', 'Real-Time Sync', 'Changes sync instantly across all group members\' devices using Firestore real-time listeners.',
     CYAN, hex_to_rgb('164e63')),
    ('🧮', 'Smart Splitting', '3 split modes: Equal, Exact Amount, Percentage. Handles complex scenarios with decimal precision.',
     PURPLE, hex_to_rgb('4c1d95')),
    ('📊', 'Visual Analytics', 'Interactive charts showing spending by category, per-person breakdown, and expense timeline.',
     PINK, hex_to_rgb('831843')),
    ('🔔', 'Push Notifications', 'Get notified when expenses are added, balances change, or settlements are requested.',
     AMBER, hex_to_rgb('78350f')),
    ('🔐', 'Secure & Private', 'Firebase Authentication with Google sign-in. Data encrypted in transit and at rest.',
     RED, hex_to_rgb('7f1d1d')),
]

feat_card_w = Inches(3.7)
feat_card_h = Inches(1.8)
feat_gap_x = Inches(0.4)
feat_gap_y = Inches(0.35)
feat_total_w = 3 * feat_card_w + 2 * feat_gap_x
feat_start_x = (SLIDE_W - feat_total_w) // 2
feat_start_y = Inches(2.5)

for i, (emoji, title, desc, accent, bg_dark) in enumerate(features):
    col = i % 3
    row = i // 3
    x = feat_start_x + col * (feat_card_w + feat_gap_x)
    y = feat_start_y + row * (feat_card_h + feat_gap_y)

    card = add_rounded_rect(slide4, x, y, feat_card_w, feat_card_h,
                           fill_color=CARD_BG, border_color=SLATE_700, border_width=Pt(1), shadow=True)

    # Icon
    icon = add_circle(slide4, x + Inches(0.25), y + Inches(0.25), Inches(0.55),
                     fill_color=bg_dark, text=emoji, font_size=Pt(20), font_color=WHITE)

    # Title
    add_text_box(slide4, x + Inches(1.0), y + Inches(0.25), feat_card_w - Inches(1.3), Inches(0.35),
                title, font_size=Pt(14), font_color=WHITE, bold=True)

    # Description
    add_text_box(slide4, x + Inches(1.0), y + Inches(0.65), feat_card_w - Inches(1.3), Inches(1.0),
                desc, font_size=Pt(10), font_color=SLATE_400)


# ══════════════════════════════════════════════════════════════════
# SLIDE 5: TECH STACK
# ══════════════════════════════════════════════════════════════════
slide5 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg_gradient(slide5, DARK_BG, DARK_BG_2)
add_gradient_rect(slide5, Inches(0), Inches(0), SLIDE_W, Inches(0.06), EMERALD, CYAN)
add_gradient_rect(slide5, Inches(0), SLIDE_H - Inches(0.06), SLIDE_W, Inches(0.06), EMERALD, CYAN)

# Section badge
badge = add_rounded_rect(slide5, Inches(0.8), Inches(0.5), Inches(2.2), Inches(0.4),
                         fill_color=GREEN_900_BG, border_color=EMERALD_DARK, border_width=Pt(1))
add_text_box(slide5, Inches(0.8), Inches(0.52), Inches(2.2), Inches(0.4),
             '🛠️  Tech Stack', font_size=Pt(11), font_color=EMERALD_LIGHT,
             alignment=PP_ALIGN.CENTER, bold=True)

add_text_box(slide5, Inches(0.8), Inches(1.1), Inches(11), Inches(0.7),
             'Technology Stack', font_size=Pt(36), font_color=WHITE, bold=True)

add_text_box(slide5, Inches(0.8), Inches(1.8), Inches(11), Inches(0.5),
             'Built with modern, scalable, and performant technologies.',
             font_size=Pt(14), font_color=SLATE_400)

# Tech stack cards
tech_categories = [
    ('Frontend', [
        ('⚛️', 'React 18', 'UI library with hooks & concurrent features', CYAN),
        ('📘', 'TypeScript', 'Type-safe development experience', BLUE),
        ('🎨', 'Tailwind CSS', 'Utility-first CSS framework', TEAL),
        ('⚡', 'Vite', 'Next-gen frontend build tool', PURPLE),
    ]),
    ('Backend & Infrastructure', [
        ('🔥', 'Firebase', 'Firestore, Auth, Cloud Functions, Hosting', AMBER),
        ('📱', 'PWA', 'Service Workers, Web App Manifest', EMERALD),
        ('🔄', 'Workbox', 'Advanced caching strategies', PINK),
        ('☁️', 'Cloud Functions', 'Serverless backend logic', CYAN),
    ]),
]

cat_y = Inches(2.5)
for cat_idx, (cat_name, techs) in enumerate(tech_categories):
    cat_x = Inches(0.8)
    # Category label
    add_text_box(slide5, cat_x, cat_y, Inches(3), Inches(0.4),
                cat_name, font_size=Pt(18), font_color=WHITE, bold=True)

    tech_y = cat_y + Inches(0.5)
    tech_card_w = Inches(2.7)
    tech_card_h = Inches(1.4)
    tech_gap = Inches(0.3)

    for j, (emoji, name, desc, color) in enumerate(techs):
        tx = cat_x + j * (tech_card_w + tech_gap)

        card = add_rounded_rect(slide5, tx, tech_y, tech_card_w, tech_card_h,
                               fill_color=CARD_BG, border_color=SLATE_700, border_width=Pt(1), shadow=True)

        # Top accent
        acc = slide5.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                       tx + Inches(0.15), tech_y, tech_card_w - Inches(0.3), Inches(0.04))
        acc.fill.solid()
        acc.fill.fore_color.rgb = color
        acc.line.fill.background()

        add_text_box(slide5, tx + Inches(0.2), tech_y + Inches(0.2), tech_card_w - Inches(0.4), Inches(0.4),
                    f'{emoji} {name}', font_size=Pt(13), font_color=WHITE, bold=True)

        add_text_box(slide5, tx + Inches(0.2), tech_y + Inches(0.6), tech_card_w - Inches(0.4), Inches(0.7),
                    desc, font_size=Pt(10), font_color=SLATE_400)

    cat_y += Inches(2.3)


# ══════════════════════════════════════════════════════════════════
# SLIDE 6: ARCHITECTURE
# ══════════════════════════════════════════════════════════════════
slide6 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg_gradient(slide6, DARK_BG, DARK_BG_2)
add_gradient_rect(slide6, Inches(0), Inches(0), SLIDE_W, Inches(0.06), EMERALD, CYAN)
add_gradient_rect(slide6, Inches(0), SLIDE_H - Inches(0.06), SLIDE_W, Inches(0.06), EMERALD, CYAN)

badge = add_rounded_rect(slide6, Inches(0.8), Inches(0.5), Inches(2.2), Inches(0.4),
                         fill_color=GREEN_900_BG, border_color=EMERALD_DARK, border_width=Pt(1))
add_text_box(slide6, Inches(0.8), Inches(0.52), Inches(2.2), Inches(0.4),
             '🏗️  Architecture', font_size=Pt(11), font_color=EMERALD_LIGHT,
             alignment=PP_ALIGN.CENTER, bold=True)

add_text_box(slide6, Inches(0.8), Inches(1.1), Inches(11), Inches(0.7),
             'System Architecture', font_size=Pt(36), font_color=WHITE, bold=True)

add_text_box(slide6, Inches(0.8), Inches(1.8), Inches(11), Inches(0.5),
             'Three-tier architecture designed for scalability, offline resilience, and real-time collaboration.',
             font_size=Pt(14), font_color=SLATE_400)

# Architecture layers - 3 columns
arch_layers = [
    ('📱 Presentation Layer', EMERALD, hex_to_rgb('064e3b'), [
        'React Components',
        'Tailwind CSS Styling',
        'React Router Navigation',
        'Progressive Web App Shell',
        'Responsive Design System',
    ]),
    ('⚙️ Business Logic Layer', CYAN, hex_to_rgb('164e63'), [
        'Expense Calculation Engine',
        'Settlement Optimization',
        'Balance Tracking Algorithm',
        'Split Mode Processor',
        'Data Validation Layer',
    ]),
    ('💾 Data & Services Layer', PURPLE, hex_to_rgb('4c1d95'), [
        'Cloud Firestore (NoSQL)',
        'Firebase Authentication',
        'Service Worker (Cache)',
        'IndexedDB (Offline)',
        'Cloud Functions (Triggers)',
    ]),
]

arch_card_w = Inches(3.6)
arch_card_h = Inches(3.8)
arch_gap = Inches(0.45)
arch_total = 3 * arch_card_w + 2 * arch_gap
arch_start_x = (SLIDE_W - arch_total) // 2
arch_y = Inches(2.5)

for i, (title, accent, bg_dark, items) in enumerate(arch_layers):
    x = arch_start_x + i * (arch_card_w + arch_gap)

    card = add_rounded_rect(slide6, x, arch_y, arch_card_w, arch_card_h,
                           fill_color=CARD_BG, border_color=SLATE_700, border_width=Pt(1), shadow=True)

    # Top accent
    acc = slide6.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   x + Inches(0.15), arch_y, arch_card_w - Inches(0.3), Inches(0.05))
    acc.fill.solid()
    acc.fill.fore_color.rgb = accent
    acc.line.fill.background()

    # Title
    add_text_box(slide6, x + Inches(0.2), arch_y + Inches(0.25), arch_card_w - Inches(0.4), Inches(0.4),
                title, font_size=Pt(15), font_color=WHITE, bold=True, alignment=PP_ALIGN.CENTER)

    # Items
    for j, item in enumerate(items):
        item_y = arch_y + Inches(0.85) + j * Inches(0.55)
        item_bg = add_rounded_rect(slide6, x + Inches(0.2), item_y,
                                   arch_card_w - Inches(0.4), Inches(0.42),
                                   fill_color=hex_to_rgb('0f172a'), border_color=SLATE_700, border_width=Pt(0.5))
        add_text_box(slide6, x + Inches(0.35), item_y + Inches(0.05),
                    arch_card_w - Inches(0.7), Inches(0.32),
                    f'• {item}', font_size=Pt(10), font_color=SLATE_300)

    # Arrows between columns
    if i < 2:
        arr_x = x + arch_card_w + Inches(0.08)
        arr_y_mid = arch_y + arch_card_h // 2
        add_text_box(slide6, arr_x, arr_y_mid - Inches(0.15), Inches(0.35), Inches(0.3),
                    '⇄', font_size=Pt(20), font_color=EMERALD, alignment=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════
# SLIDE 7: ALGORITHM – Settlement Optimization
# ══════════════════════════════════════════════════════════════════
slide7 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg_gradient(slide7, DARK_BG, DARK_BG_2)
add_gradient_rect(slide7, Inches(0), Inches(0), SLIDE_W, Inches(0.06), EMERALD, CYAN)
add_gradient_rect(slide7, Inches(0), SLIDE_H - Inches(0.06), SLIDE_W, Inches(0.06), EMERALD, CYAN)

badge = add_rounded_rect(slide7, Inches(0.8), Inches(0.5), Inches(2.5), Inches(0.4),
                         fill_color=GREEN_900_BG, border_color=EMERALD_DARK, border_width=Pt(1))
add_text_box(slide7, Inches(0.8), Inches(0.52), Inches(2.5), Inches(0.4),
             '🧮  Algorithm', font_size=Pt(11), font_color=EMERALD_LIGHT,
             alignment=PP_ALIGN.CENTER, bold=True)

add_text_box(slide7, Inches(0.8), Inches(1.1), Inches(11), Inches(0.7),
             'Settlement Optimization', font_size=Pt(36), font_color=WHITE, bold=True)

add_text_box(slide7, Inches(0.8), Inches(1.8), Inches(11), Inches(0.5),
             'Minimizing the number of transactions needed to settle all debts.',
             font_size=Pt(14), font_color=SLATE_400)

# Before vs After comparison
comp_y = Inches(2.5)

# "Before" card
before_w = Inches(5.5)
before_h = Inches(4.2)
before_x = Inches(0.8)
before_card = add_rounded_rect(slide7, before_x, comp_y, before_w, before_h,
                               fill_color=CARD_BG, border_color=hex_to_rgb('7f1d1d'), border_width=Pt(1.5), shadow=True)

add_text_box(slide7, before_x + Inches(0.3), comp_y + Inches(0.25), before_w - Inches(0.6), Inches(0.4),
            '❌  Without Optimization', font_size=Pt(16), font_color=RED, bold=True)

add_text_box(slide7, before_x + Inches(0.3), comp_y + Inches(0.7), before_w - Inches(0.6), Inches(0.3),
            '6 transactions needed', font_size=Pt(12), font_color=SLATE_400)

# Before transactions
before_txns = [
    'Alice → Bob: $25.00',
    'Alice → Charlie: $15.00',
    'Bob → Charlie: $10.00',
    'Bob → Diana: $20.00',
    'Charlie → Diana: $5.00',
    'Diana → Alice: $30.00',
]

for j, txn in enumerate(before_txns):
    ty = comp_y + Inches(1.15) + j * Inches(0.45)
    txn_bg = add_rounded_rect(slide7, before_x + Inches(0.3), ty,
                              before_w - Inches(0.6), Inches(0.38),
                              fill_color=hex_to_rgb('7f1d1d'))
    add_text_box(slide7, before_x + Inches(0.45), ty + Inches(0.03),
                before_w - Inches(0.9), Inches(0.32),
                txn, font_size=Pt(10), font_color=hex_to_rgb('fca5a5'))

# "After" card
after_x = Inches(6.8)
after_w = Inches(5.5)
after_h = Inches(4.2)
after_card = add_rounded_rect(slide7, after_x, comp_y, after_w, after_h,
                              fill_color=CARD_BG, border_color=hex_to_rgb('064e3b'), border_width=Pt(1.5), shadow=True)

add_text_box(slide7, after_x + Inches(0.3), comp_y + Inches(0.25), after_w - Inches(0.6), Inches(0.4),
            '✅  With TripSplit Algorithm', font_size=Pt(16), font_color=EMERALD, bold=True)

add_text_box(slide7, after_x + Inches(0.3), comp_y + Inches(0.7), after_w - Inches(0.6), Inches(0.3),
            'Only 2 transactions needed!', font_size=Pt(12), font_color=SLATE_400)

# After transactions
after_txns = [
    'Alice → Charlie: $10.00',
    'Bob → Diana: $15.00',
]

for j, txn in enumerate(after_txns):
    ty = comp_y + Inches(1.15) + j * Inches(0.45)
    txn_bg = add_rounded_rect(slide7, after_x + Inches(0.3), ty,
                              after_w - Inches(0.6), Inches(0.38),
                              fill_color=hex_to_rgb('064e3b'))
    add_text_box(slide7, after_x + Inches(0.45), ty + Inches(0.03),
                after_w - Inches(0.9), Inches(0.32),
                txn, font_size=Pt(10), font_color=EMERALD_LIGHT)

# Savings badge
savings_y = comp_y + Inches(2.2)
savings_bg = add_rounded_rect(slide7, after_x + Inches(0.5), savings_y,
                              after_w - Inches(1.0), Inches(1.2),
                              fill_color=hex_to_rgb('064e3b'), border_color=EMERALD_DARK, border_width=Pt(1))
add_text_box(slide7, after_x + Inches(0.7), savings_y + Inches(0.1),
            after_w - Inches(1.4), Inches(0.5),
            '67% fewer transactions', font_size=Pt(20), font_color=EMERALD, bold=True,
            alignment=PP_ALIGN.CENTER)
add_text_box(slide7, after_x + Inches(0.7), savings_y + Inches(0.65),
            after_w - Inches(1.4), Inches(0.4),
            'Greedy algorithm with O(n log n) complexity',
            font_size=Pt(10), font_color=EMERALD_LIGHT, alignment=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════
# SLIDE 8: DATABASE SCHEMA
# ══════════════════════════════════════════════════════════════════
slide8 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg_gradient(slide8, DARK_BG, DARK_BG_2)
add_gradient_rect(slide8, Inches(0), Inches(0), SLIDE_W, Inches(0.06), EMERALD, CYAN)
add_gradient_rect(slide8, Inches(0), SLIDE_H - Inches(0.06), SLIDE_W, Inches(0.06), EMERALD, CYAN)

badge = add_rounded_rect(slide8, Inches(0.8), Inches(0.5), Inches(2.5), Inches(0.4),
                         fill_color=GREEN_900_BG, border_color=EMERALD_DARK, border_width=Pt(1))
add_text_box(slide8, Inches(0.8), Inches(0.52), Inches(2.5), Inches(0.4),
             '🗄️  Database', font_size=Pt(11), font_color=EMERALD_LIGHT,
             alignment=PP_ALIGN.CENTER, bold=True)

add_text_box(slide8, Inches(0.8), Inches(1.1), Inches(11), Inches(0.7),
             'Database Schema', font_size=Pt(36), font_color=WHITE, bold=True)

add_text_box(slide8, Inches(0.8), Inches(1.8), Inches(11), Inches(0.5),
             'Firestore NoSQL collections designed for real-time sync and offline access.',
             font_size=Pt(14), font_color=SLATE_400)

# Schema cards - 3 columns
schemas = [
    ('👥 Users Collection', EMERALD, [
        ('uid', 'string (PK)'),
        ('displayName', 'string'),
        ('email', 'string'),
        ('photoURL', 'string | null'),
        ('createdAt', 'timestamp'),
        ('groups[ ]', 'string[ ]'),
    ]),
    ('🏕️ Groups Collection', CYAN, [
        ('groupId', 'string (PK)'),
        ('name', 'string'),
        ('description', 'string'),
        ('createdBy', 'string (FK→Users)'),
        ('members[ ]', 'MemberObject[ ]'),
        ('inviteCode', 'string (unique)'),
        ('currency', 'string'),
        ('createdAt', 'timestamp'),
    ]),
    ('💰 Expenses Collection', PURPLE, [
        ('expenseId', 'string (PK)'),
        ('groupId', 'string (FK→Groups)'),
        ('description', 'string'),
        ('amount', 'number'),
        ('paidBy', 'string (FK→Users)'),
        ('splitType', 'equal|exact|percent'),
        ('splits[ ]', 'SplitObject[ ]'),
        ('category', 'string'),
        ('date', 'timestamp'),
    ]),
]

schema_card_w = Inches(3.6)
schema_card_h = Inches(4.3)
schema_gap = Inches(0.4)
schema_total = 3 * schema_card_w + 2 * schema_gap
schema_start_x = (SLIDE_W - schema_total) // 2
schema_y = Inches(2.5)

for i, (title, accent, fields) in enumerate(schemas):
    x = schema_start_x + i * (schema_card_w + schema_gap)

    card = add_rounded_rect(slide8, x, schema_y, schema_card_w, schema_card_h,
                           fill_color=CARD_BG, border_color=SLATE_700, border_width=Pt(1), shadow=True)

    # Accent bar
    acc = slide8.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                   x + Inches(0.1), schema_y, schema_card_w - Inches(0.2), Inches(0.05))
    acc.fill.solid()
    acc.fill.fore_color.rgb = accent
    acc.line.fill.background()

    # Title
    add_text_box(slide8, x + Inches(0.2), schema_y + Inches(0.2),
                schema_card_w - Inches(0.4), Inches(0.35),
                title, font_size=Pt(14), font_color=WHITE, bold=True)

    # Fields
    for j, (field_name, field_type) in enumerate(fields):
        fy = schema_y + Inches(0.7) + j * Inches(0.42)
        field_bg = add_rounded_rect(slide8, x + Inches(0.15), fy,
                                    schema_card_w - Inches(0.3), Inches(0.36),
                                    fill_color=hex_to_rgb('0f172a'), border_color=SLATE_700, border_width=Pt(0.5))

        # Field name (left)
        add_text_box(slide8, x + Inches(0.25), fy + Inches(0.03),
                    Inches(1.4), Inches(0.3),
                    field_name, font_size=Pt(9), font_color=EMERALD_LIGHT, bold=True,
                    font_name='Consolas')

        # Field type (right)
        add_text_box(slide8, x + Inches(1.7), fy + Inches(0.03),
                    schema_card_w - Inches(2.0), Inches(0.3),
                    field_type, font_size=Pt(8), font_color=SLATE_400,
                    font_name='Consolas')


# ══════════════════════════════════════════════════════════════════
# SLIDE 9: SCREENS / UI PREVIEW
# ══════════════════════════════════════════════════════════════════
slide9 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg_gradient(slide9, DARK_BG, DARK_BG_2)
add_gradient_rect(slide9, Inches(0), Inches(0), SLIDE_W, Inches(0.06), EMERALD, CYAN)
add_gradient_rect(slide9, Inches(0), SLIDE_H - Inches(0.06), SLIDE_W, Inches(0.06), EMERALD, CYAN)

badge = add_rounded_rect(slide9, Inches(0.8), Inches(0.5), Inches(2.5), Inches(0.4),
                         fill_color=GREEN_900_BG, border_color=EMERALD_DARK, border_width=Pt(1))
add_text_box(slide9, Inches(0.8), Inches(0.52), Inches(2.5), Inches(0.4),
             '📱  UI Preview', font_size=Pt(11), font_color=EMERALD_LIGHT,
             alignment=PP_ALIGN.CENTER, bold=True)

add_text_box(slide9, Inches(0.8), Inches(1.1), Inches(11), Inches(0.7),
             'App Screens', font_size=Pt(36), font_color=WHITE, bold=True)

add_text_box(slide9, Inches(0.8), Inches(1.8), Inches(11), Inches(0.5),
             'Clean, intuitive mobile-first interface designed for ease of use.',
             font_size=Pt(14), font_color=SLATE_400)

# Mock phone screens - 4 screens
screens = [
    ('Dashboard', EMERALD, [
        ('🏕️ Bali Trip 2024', '$1,247.50 total'),
        ('👥 5 members', '$249.50/person'),
        ('📊 Your balance', '+$52.30'),
    ]),
    ('Add Expense', CYAN, [
        ('📝 Description', 'Beach dinner...'),
        ('💰 Amount', '$185.00'),
        ('👥 Split: Equal', '5 people × $37.00'),
    ]),
    ('Balances', PURPLE, [
        ('Alice', '+$52.30 owed'),
        ('Bob', '-$28.50 owes'),
        ('Charlie', '+$15.00 owed'),
    ]),
    ('Settlement', AMBER, [
        ('Bob → Alice', '$28.50'),
        ('Diana → Charlie', '$15.00'),
        ('✅ 2 transactions', 'optimized from 6'),
    ]),
]

phone_w = Inches(2.5)
phone_h = Inches(4.2)
phone_gap = Inches(0.5)
phone_total = 4 * phone_w + 3 * phone_gap
phone_start_x = (SLIDE_W - phone_total) // 2
phone_y = Inches(2.5)

for i, (screen_name, accent, items) in enumerate(screens):
    x = phone_start_x + i * (phone_w + phone_gap)

    # Phone frame
    phone = add_rounded_rect(slide9, x, phone_y, phone_w, phone_h,
                            fill_color=hex_to_rgb('0f172a'), border_color=SLATE_600, border_width=Pt(2), shadow=True)

    # Status bar area
    status = add_rounded_rect(slide9, x + Inches(0.1), phone_y + Inches(0.1),
                              phone_w - Inches(0.2), Inches(0.35),
                              fill_color=CARD_BG)

    # Screen title in status bar
    add_text_box(slide9, x + Inches(0.15), phone_y + Inches(0.13),
                phone_w - Inches(0.3), Inches(0.3),
                screen_name, font_size=Pt(11), font_color=accent, bold=True,
                alignment=PP_ALIGN.CENTER)

    # Content items
    for j, (label, value) in enumerate(items):
        iy = phone_y + Inches(0.6) + j * Inches(0.95)

        item_bg = add_rounded_rect(slide9, x + Inches(0.15), iy,
                                   phone_w - Inches(0.3), Inches(0.8),
                                   fill_color=CARD_BG, border_color=SLATE_700, border_width=Pt(0.5))

        add_text_box(slide9, x + Inches(0.25), iy + Inches(0.08),
                    phone_w - Inches(0.5), Inches(0.3),
                    label, font_size=Pt(10), font_color=WHITE, bold=True)

        add_text_box(slide9, x + Inches(0.25), iy + Inches(0.38),
                    phone_w - Inches(0.5), Inches(0.3),
                    value, font_size=Pt(9), font_color=SLATE_400)

    # Bottom nav bar
    nav = add_rounded_rect(slide9, x + Inches(0.1), phone_y + phone_h - Inches(0.45),
                          phone_w - Inches(0.2), Inches(0.35),
                          fill_color=CARD_BG)
    add_text_box(slide9, x + Inches(0.15), phone_y + phone_h - Inches(0.42),
                phone_w - Inches(0.3), Inches(0.3),
                '🏠    💰    📊    ⚙️', font_size=Pt(12), font_color=SLATE_400,
                alignment=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════
# SLIDE 10: ROADMAP / TIMELINE
# ══════════════════════════════════════════════════════════════════
slide10 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg_gradient(slide10, DARK_BG, DARK_BG_2)
add_gradient_rect(slide10, Inches(0), Inches(0), SLIDE_W, Inches(0.06), EMERALD, CYAN)
add_gradient_rect(slide10, Inches(0), SLIDE_H - Inches(0.06), SLIDE_W, Inches(0.06), EMERALD, CYAN)

badge = add_rounded_rect(slide10, Inches(0.8), Inches(0.5), Inches(2.5), Inches(0.4),
                         fill_color=GREEN_900_BG, border_color=EMERALD_DARK, border_width=Pt(1))
add_text_box(slide10, Inches(0.8), Inches(0.52), Inches(2.5), Inches(0.4),
             '🗓️  Roadmap', font_size=Pt(11), font_color=EMERALD_LIGHT,
             alignment=PP_ALIGN.CENTER, bold=True)

add_text_box(slide10, Inches(0.8), Inches(1.1), Inches(11), Inches(0.7),
             'Development Roadmap', font_size=Pt(36), font_color=WHITE, bold=True)

add_text_box(slide10, Inches(0.8), Inches(1.8), Inches(11), Inches(0.5),
             'Phased development approach ensuring quality at every stage.',
             font_size=Pt(14), font_color=SLATE_400)

# Timeline phases
phases = [
    ('Phase 1', 'Foundation', 'Weeks 1-3', EMERALD, hex_to_rgb('064e3b'), '✅', [
        'Project setup & architecture',
        'Firebase integration',
        'User authentication',
        'Basic group CRUD',
    ]),
    ('Phase 2', 'Core Features', 'Weeks 4-6', CYAN, hex_to_rgb('164e63'), '🔄', [
        'Expense management',
        'Split calculations',
        'Balance tracking',
        'Settlement algorithm',
    ]),
    ('Phase 3', 'Enhancement', 'Weeks 7-9', PURPLE, hex_to_rgb('4c1d95'), '📋', [
        'PWA & offline support',
        'Push notifications',
        'Visual analytics',
        'Performance optimization',
    ]),
    ('Phase 4', 'Polish & Launch', 'Weeks 10-12', AMBER, hex_to_rgb('78350f'), '🚀', [
        'Testing & bug fixes',
        'UI/UX refinements',
        'Documentation',
        'Production deployment',
    ]),
]

phase_card_w = Inches(2.7)
phase_card_h = Inches(3.8)
phase_gap = Inches(0.35)
phase_total = 4 * phase_card_w + 3 * phase_gap
phase_start_x = (SLIDE_W - phase_total) // 2
phase_y = Inches(2.5)

# Connecting line
line_y = phase_y + Inches(0.5)
timeline_line = slide10.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                          phase_start_x, line_y + Inches(0.12),
                                          phase_total, Inches(0.04))
timeline_line.fill.solid()
timeline_line.fill.fore_color.rgb = SLATE_700
timeline_line.line.fill.background()

for i, (phase, title, period, accent, bg_dark, status_emoji, tasks) in enumerate(phases):
    x = phase_start_x + i * (phase_card_w + phase_gap)

    # Phase dot on timeline
    dot = add_circle(slide10, x + (phase_card_w - Inches(0.3)) // 2, line_y,
                    Inches(0.3), fill_color=accent, text='', font_size=Pt(8))

    # Card below
    card = add_rounded_rect(slide10, x, phase_y + Inches(0.8), phase_card_w, phase_card_h - Inches(0.8),
                           fill_color=CARD_BG, border_color=SLATE_700, border_width=Pt(1), shadow=True)

    # Top accent
    acc_top = slide10.shapes.add_shape(MSO_SHAPE.RECTANGLE,
                                        x + Inches(0.1), phase_y + Inches(0.8),
                                        phase_card_w - Inches(0.2), Inches(0.04))
    acc_top.fill.solid()
    acc_top.fill.fore_color.rgb = accent
    acc_top.line.fill.background()

    # Phase label
    phase_badge = add_rounded_rect(slide10, x + Inches(0.15), phase_y + Inches(1.0),
                                   Inches(1.2), Inches(0.3),
                                   fill_color=bg_dark, border_color=accent, border_width=Pt(0.5))
    add_text_box(slide10, x + Inches(0.15), phase_y + Inches(1.02),
                Inches(1.2), Inches(0.28),
                f'{status_emoji} {phase}', font_size=Pt(9), font_color=accent,
                alignment=PP_ALIGN.CENTER, bold=True)

    # Title
    add_text_box(slide10, x + Inches(0.2), phase_y + Inches(1.4),
                phase_card_w - Inches(0.4), Inches(0.35),
                title, font_size=Pt(14), font_color=WHITE, bold=True)

    # Period
    add_text_box(slide10, x + Inches(0.2), phase_y + Inches(1.75),
                phase_card_w - Inches(0.4), Inches(0.25),
                period, font_size=Pt(10), font_color=SLATE_400, italic=True)

    # Tasks
    for j, task in enumerate(tasks):
        ty = phase_y + Inches(2.1) + j * Inches(0.38)
        add_text_box(slide10, x + Inches(0.25), ty,
                    phase_card_w - Inches(0.5), Inches(0.32),
                    f'• {task}', font_size=Pt(9), font_color=SLATE_300)


# ══════════════════════════════════════════════════════════════════
# SLIDE 11: TEAM
# ══════════════════════════════════════════════════════════════════
slide11 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg_gradient(slide11, DARK_BG, DARK_BG_2)
add_gradient_rect(slide11, Inches(0), Inches(0), SLIDE_W, Inches(0.06), EMERALD, CYAN)
add_gradient_rect(slide11, Inches(0), SLIDE_H - Inches(0.06), SLIDE_W, Inches(0.06), EMERALD, CYAN)

badge = add_rounded_rect(slide11, Inches(0.8), Inches(0.5), Inches(2.5), Inches(0.4),
                         fill_color=GREEN_900_BG, border_color=EMERALD_DARK, border_width=Pt(1))
add_text_box(slide11, Inches(0.8), Inches(0.52), Inches(2.5), Inches(0.4),
             '👥  Team', font_size=Pt(11), font_color=EMERALD_LIGHT,
             alignment=PP_ALIGN.CENTER, bold=True)

add_text_box(slide11, Inches(0.8), Inches(1.1), Inches(11), Inches(0.7),
             'Meet the Team', font_size=Pt(36), font_color=WHITE, bold=True)

add_text_box(slide11, Inches(0.8), Inches(1.8), Inches(11), Inches(0.5),
             'The people behind TripSplit.',
             font_size=Pt(14), font_color=SLATE_400)

# Team member cards
team = [
    ('👨‍💻', 'AJ', 'Alex Johnson', 'Full-Stack Developer', 'React, Firebase, TypeScript\nSystem architecture & algorithm design', EMERALD),
    ('🎨', 'SM', 'Sarah Miller', 'UI/UX Designer', 'Figma, Tailwind CSS\nUser research & interface design', PINK),
    ('⚙️', 'MC', 'Mike Chen', 'Backend Developer', 'Cloud Functions, Firestore\nAPI design & optimization', CYAN),
    ('🧪', 'EP', 'Emily Park', 'QA Engineer', 'Jest, Cypress, Testing\nQuality assurance & automation', PURPLE),
]

member_card_w = Inches(2.7)
member_card_h = Inches(3.5)
member_gap = Inches(0.4)
member_total = 4 * member_card_w + 3 * member_gap
member_start_x = (SLIDE_W - member_total) // 2
member_y = Inches(2.5)

for i, (emoji, initials, name, role, skills, accent) in enumerate(team):
    x = member_start_x + i * (member_card_w + member_gap)

    card = add_rounded_rect(slide11, x, member_y, member_card_w, member_card_h,
                           fill_color=CARD_BG, border_color=SLATE_700, border_width=Pt(1), shadow=True)

    # Avatar circle
    avatar_size = Inches(0.9)
    avatar_x = x + (member_card_w - avatar_size) // 2
    avatar_y = member_y + Inches(0.35)
    avatar = add_circle(slide11, avatar_x, avatar_y, avatar_size,
                       fill_color=accent)

    # Initials in avatar
    add_text_box(slide11, avatar_x, avatar_y + Inches(0.15),
                avatar_size, Inches(0.6),
                initials, font_size=Pt(24), font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER)

    # Name
    add_text_box(slide11, x + Inches(0.15), member_y + Inches(1.5),
                member_card_w - Inches(0.3), Inches(0.35),
                name, font_size=Pt(14), font_color=WHITE, bold=True,
                alignment=PP_ALIGN.CENTER)

    # Role badge
    role_badge_w = Inches(2.0)
    role_badge_x = x + (member_card_w - role_badge_w) // 2
    role_badge = add_rounded_rect(slide11, role_badge_x, member_y + Inches(1.9),
                                  role_badge_w, Inches(0.3),
                                  fill_color=hex_to_rgb('064e3b'))
    add_text_box(slide11, role_badge_x, member_y + Inches(1.92),
                role_badge_w, Inches(0.28),
                role, font_size=Pt(9), font_color=EMERALD_LIGHT,
                alignment=PP_ALIGN.CENTER, bold=True)

    # Skills
    add_text_box(slide11, x + Inches(0.2), member_y + Inches(2.4),
                member_card_w - Inches(0.4), Inches(0.9),
                skills, font_size=Pt(9), font_color=SLATE_400,
                alignment=PP_ALIGN.CENTER)


# ══════════════════════════════════════════════════════════════════
# SLIDE 12: THANK YOU / CLOSING
# ══════════════════════════════════════════════════════════════════
slide12 = prs.slides.add_slide(prs.slide_layouts[6])
set_slide_bg_gradient(slide12, DARK_BG, DARK_BG_2)
add_gradient_rect(slide12, Inches(0), Inches(0), SLIDE_W, Inches(0.06), EMERALD, CYAN)
add_gradient_rect(slide12, Inches(0), SLIDE_H - Inches(0.06), SLIDE_W, Inches(0.06), EMERALD, CYAN)

# Center card
end_card_w = Inches(8)
end_card_h = Inches(4.5)
end_card_x = (SLIDE_W - end_card_w) // 2
end_card_y = Inches(1.5)

# Glow
glow2 = add_rounded_rect(slide12, end_card_x - Inches(0.3), end_card_y - Inches(0.3),
                          end_card_w + Inches(0.6), end_card_h + Inches(0.6),
                          fill_color=EMERALD)
spPr2 = glow2._element.spPr
fill_elem2 = glow2.fill._fill.find(qn('a:solidFill'))
if fill_elem2 is not None:
    for child in fill_elem2:
        alpha2 = etree.SubElement(child, qn('a:alpha'))
        alpha2.set('val', '8000')

main_end_card = add_rounded_rect(slide12, end_card_x, end_card_y, end_card_w, end_card_h,
                                 fill_color=CARD_BG, border_color=SLATE_700, border_width=Pt(1), shadow=True)

# Thank you icon
ty_icon = add_circle(slide12, (SLIDE_W - Inches(1.0)) // 2, end_card_y + Inches(0.4),
                    Inches(1.0), fill_color=EMERALD, text='🎉', font_size=Pt(36), font_color=WHITE)

# Thank You text
add_text_box(slide12, end_card_x, end_card_y + Inches(1.6), end_card_w, Inches(0.7),
             'Thank You!', font_size=Pt(40), font_color=WHITE, bold=True,
             alignment=PP_ALIGN.CENTER)

add_text_box(slide12, end_card_x, end_card_y + Inches(2.3), end_card_w, Inches(0.5),
             'Ready to revolutionize group expense management', font_size=Pt(16),
             font_color=EMERALD_LIGHT, alignment=PP_ALIGN.CENTER)

# Contact info
contact_items = [
    ('🌐  github.com/tripsplit', SLATE_300),
    ('📧  team@tripsplit.app', SLATE_300),
    ('🔗  tripsplit.web.app', SLATE_300),
]

for j, (contact, color) in enumerate(contact_items):
    cy = end_card_y + Inches(3.0) + j * Inches(0.35)
    add_text_box(slide12, end_card_x, cy, end_card_w, Inches(0.3),
                contact, font_size=Pt(12), font_color=color,
                alignment=PP_ALIGN.CENTER)

# ── Save the presentation ────────────────────────────────────────
output_path = 'TripSplit_Presentation.pptx'
prs.save(output_path)
print(f"✅ Presentation saved to: {output_path}")
print(f"📊 Total slides: {len(prs.slides)}")
