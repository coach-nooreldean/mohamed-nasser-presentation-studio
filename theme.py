"""
Sofian Modern Cinema Design System & Typography Engine
Tailored for Mr. Mohamed Nasser Badawy (مستر محمد ناصر بدوي)

Color Palette (High-Contrast Dark Modern Cinema):
- Deep Navy Canvas: #0A1128 (Background)
- Slate Glass Surface: #131D31 (Card Base)
- Elevated Surface: #1A263E (Card Hover / Pills)
- Crisp Slate Border: #2A3F5F (Refined Structural Lines)
- Vibrant Gold: #FFB703 (Primary Accent, Badges, Highlights)
- Electric Sky Cyan: #38BDF8 (Phonetics, Subtitles, Rules)
- Emerald Mint: #10B981 (Examples, Positive forms, WhatsApp)
- Crisp White: #FFFFFF (Titles & Core Text)
- High-Contrast Body: #F1F5F9 (Paragraphs & Explanations)
- Slate Secondary: #CBD5E1 (Sub-bullets & Insights)
- Soft Muted: #94A3B8 (Footnotes & Metadata)
"""

import re
from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

class SofianColors:
    # Backgrounds
    BG_DARK = RGBColor(10, 17, 40)          # #0A1128 Deep Midnight Navy
    BG_GLOW_TOP = RGBColor(16, 31, 60)      # #101F3C Ambient top shade
    BG_GLOW_BOTTOM = RGBColor(13, 23, 44)   # #0D172C Ambient bottom shade

    # Card Surfaces
    CARD_BG = RGBColor(19, 29, 49)          # #131D31 Slate Glass
    CARD_SURFACE_ALT = RGBColor(26, 38, 62) # #1A263E Elevated pill surface
    CARD_BORDER = RGBColor(42, 63, 95)      # #2A3F5F Crisp border
    CARD_BORDER_GLOW = RGBColor(56, 189, 248) # #38BDF8 Cyan highlight border

    # Accents & Vibrancy
    AMBER_GOLD = RGBColor(255, 183, 3)      # #FFB703 Vibrant Golden Amber
    ELECTRIC_CYAN = RGBColor(56, 189, 248)  # #38BDF8 Electric Sky Cyan
    EMERALD_MINT = RGBColor(16, 185, 129)   # #10B981 Emerald Spring Green
    CORAL_ROSE = RGBColor(244, 63, 94)      # #F43F5E Warning / Pitfall Rose

    # High-Contrast Typography Colors
    TEXT_LIGHT = RGBColor(255, 255, 255)    # #FFFFFF Pure crisp white (Titles)
    TEXT_BODY = RGBColor(241, 245, 249)     # #F1F5F9 Bright high-contrast body
    TEXT_SECONDARY = RGBColor(203, 213, 225)# #CBD5E1 Slate 300 (Sub-points)
    TEXT_MUTED = RGBColor(148, 163, 184)    # #94A3B8 Slate 400 (Metadata)

    # Backward compatibility aliases
    DARK_BG = BG_DARK
    DARK_SEA_BLUE = BG_GLOW_TOP
    CHARCOAL_CARD = CARD_BG
    CHARCOAL_LIGHT = CARD_SURFACE_ALT
    ORANGE_ACCENT = AMBER_GOLD
    LIGHT_BLUE = ELECTRIC_CYAN
    MINT_GREEN = EMERALD_MINT
    GOLDEN_ACCENT = AMBER_GOLD

class SofianFonts:
    # High-compatibility, crisp modern system fonts on Windows and macOS
    ARABIC = "Segoe UI"
    ENGLISH = "Segoe UI"
    FALLBACK_AR = "Dubai"
    FALLBACK_EN = "Calibri"

def is_primarily_arabic(text):
    """
    Returns True ONLY if the text is primarily Arabic.
    Prevents English sentences with Arabic letters from inverting brackets or BiDi flow!
    """
    if not text:
        return False
    clean = str(text).strip()
    ar_letters = sum(1 for c in clean if '\u0600' <= c <= '\u06FF')
    en_letters = sum(1 for c in clean if 'a' <= c.lower() <= 'z')
    if ar_letters == 0:
        return False
    if en_letters == 0:
        return True
    words = clean.split()[:2]
    first_ar = any('\u0600' <= c <= '\u06FF' for c in words[0]) if words else False
    if first_ar and ar_letters > en_letters:
        return True
    return ar_letters > (en_letters * 1.5)

def normalize_quranic_brackets(text):
    """
    Cleanses and standardizes Quranic brackets to prevent PowerPoint BiDi mirroring bugs.
    In PowerPoint, ornate brackets ﴿ and ﴾ (U+FD3E and U+FD3F) have Bidi_Mirrored=Yes,
    causing them to flip outward like ') verse (' in mixed Arabic/English contexts.
    This function standardizes them to elegant Arabic quotation marks « verse » (guillemets)
    which never flip and are standard in Islamic publishing.
    """
    if not text or not isinstance(text, str):
        return text
    s = re.sub(r'[\uFD3E\uFD3F]\s*([^«»\uFD3E\uFD3F]+?)\s*[\uFD3E\uFD3F]', r'« \1 »', text)
    s = s.replace('\uFD3F', '« ').replace('\uFD3E', ' »')
    s = s.replace('﴿', '« ').replace('﴾', ' »')
    s = re.sub(r'»\s*([\u0600-\u06FF\s]+?)\s*«', r'« \1 »', s)
    s = re.sub(r'«\s+', '« ', s)
    s = re.sub(r'\s+»', ' »', s)
    s = re.sub(r' +', ' ', s)
    return s.strip()

def apply_typography(run, font_name="Segoe UI", size_pt=14, color_rgb=None, bold=False, italic=False, is_arabic=False):
    """
    Applies unified Latin and Complex Script (Arabic) font typography to avoid
    Windows falling back to Times New Roman or misaligning Arabic letters.
    """
    if color_rgb is None:
        color_rgb = SofianColors.TEXT_BODY

    run.font.name = font_name
    run.font.size = Pt(size_pt)
    run.font.color.rgb = color_rgb
    run.font.bold = bold
    run.font.italic = italic

    # Explicitly inject Complex Script (<a:cs>) font for pristine Arabic kerning
    try:
        rpr = run._r.get_or_add_rPr()
        for child in list(rpr):
            if child.tag.endswith('cs'):
                rpr.remove(child)
        cs_font = "Dubai" if is_arabic else "Segoe UI"
        cs_elem = parse_xml(f'<a:cs {nsdecls("a")} typeface="{cs_font}"/>')
        rpr.append(cs_elem)
    except Exception:
        pass

def format_paragraph(paragraph, align=PP_ALIGN.LEFT, is_arabic=False, space_after_pt=4, line_spacing=None):
    """Formats paragraph alignment and enables native Right-To-Left (RTL) for Arabic."""
    paragraph.alignment = align
    if is_arabic:
        try:
            pPr = paragraph._p.get_or_add_pPr()
            pPr.set('rtl', '1')
        except Exception:
            pass
    if space_after_pt:
        paragraph.space_after = Pt(space_after_pt)
    if line_spacing:
        paragraph.line_spacing = line_spacing

def apply_slide_fade_transition(slide):
    """Applies smooth PowerPoint native fade transition to the slide."""
    try:
        for child in slide._element:
            if child.tag.endswith('transition'):
                return
        transition_xml = parse_xml(
            f'<p:transition {nsdecls("p")} spd="med"><p:fade/></p:transition>'
        )
        slide._element.append(transition_xml)
    except Exception:
        pass

def apply_slide_background(slide, prs):
    """
    Creates a deep, harmonious cinema background without harsh distracting circles.
    Features subtle ambient lighting and a silky slide fade transition.
    """
    # 1. Base Dark Midnight Canvas
    bg_shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0), Inches(0), prs.slide_width, prs.slide_height
    )
    bg_shape.fill.solid()
    bg_shape.fill.fore_color.rgb = SofianColors.BG_DARK
    bg_shape.line.fill.background()

    # 2. Subtle top-right ambient glow
    glow_top = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(9.0), Inches(-1.8), Inches(6.0), Inches(4.5)
    )
    glow_top.fill.solid()
    glow_top.fill.fore_color.rgb = SofianColors.BG_GLOW_TOP
    glow_top.line.fill.background()

    # 3. Subtle bottom-left ambient glow
    glow_bottom = slide.shapes.add_shape(
        MSO_SHAPE.OVAL, Inches(-1.5), Inches(5.0), Inches(5.0), Inches(4.0)
    )
    glow_bottom.fill.solid()
    glow_bottom.fill.fore_color.rgb = SofianColors.BG_GLOW_BOTTOM
    glow_bottom.line.fill.background()

    # 4. Inject slide transition
    apply_slide_fade_transition(slide)

def add_header(slide, title_text, category_badge=None, is_arabic=None):
    """
    Ultra-clean, modern header with optional sleek pill badge and refined glowing accent bar.
    """
    if is_arabic is None:
        is_arabic = is_primarily_arabic(title_text)

    # Pill Badge Shape (rendered only when category_badge is explicitly provided)
    if category_badge:
        badge = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(0.38), Inches(3.4), Inches(0.38)
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = SofianColors.CARD_SURFACE_ALT
        badge.line.color.rgb = SofianColors.AMBER_GOLD
        badge.line.width = Pt(1.4)

        tf_badge = badge.text_frame
        tf_badge.word_wrap = True
        tf_badge.margin_left = Inches(0.1)
        tf_badge.margin_right = Inches(0.1)
        tf_badge.margin_top = Inches(0.04)
        tf_badge.margin_bottom = Inches(0.04)

        p_b = tf_badge.paragraphs[0]
        p_b.text = ""
        run_b = p_b.add_run()
        run_b.text = f"✦  {category_badge.upper() if not is_arabic else category_badge}"
        apply_typography(run_b, SofianFonts.ARABIC if is_arabic else SofianFonts.ENGLISH, 10, SofianColors.AMBER_GOLD, bold=True, is_arabic=is_arabic)
        format_paragraph(p_b, align=PP_ALIGN.CENTER, is_arabic=is_arabic)
        title_y = Inches(0.82)
    else:
        title_y = Inches(0.60)

    # Main Slide Title
    title_box = slide.shapes.add_textbox(Inches(0.8), title_y, Inches(11.733), Inches(0.85))
    tf_title = title_box.text_frame
    tf_title.word_wrap = True
    tf_title.margin_left = Inches(0.05)
    tf_title.margin_top = Inches(0.05)

    p_t = tf_title.paragraphs[0]
    p_t.text = ""
    run_t = p_t.add_run()
    run_t.text = title_text
    apply_typography(run_t, SofianFonts.ARABIC if is_arabic else SofianFonts.ENGLISH, 26, SofianColors.TEXT_LIGHT, bold=True, is_arabic=is_arabic)
    format_paragraph(p_t, align=PP_ALIGN.RIGHT if is_arabic else PP_ALIGN.LEFT, is_arabic=is_arabic)

    # Elegant subtle accent divider line
    divider = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.65), Inches(11.733), Inches(0.025)
    )
    divider.fill.solid()
    divider.fill.fore_color.rgb = SofianColors.CARD_BORDER
    divider.line.fill.background()

    # Glowing amber accent bar in divider
    amber_tab = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE, Inches(0.8), Inches(1.64), Inches(2.2), Inches(0.04)
    )
    amber_tab.fill.solid()
    amber_tab.fill.fore_color.rgb = SofianColors.AMBER_GOLD
    amber_tab.line.fill.background()
