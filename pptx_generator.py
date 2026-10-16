"""
Cinema-Grade Presentation Generator for Mr. Mohamed Nasser Badawy
Implements Dark Modern Cinema Aesthetic (Deep Navy, Slate Glass, Amber Gold, Electric Cyan, Mint)
Includes:
- Native PowerPoint Click Animations (Progressive Disclosure)
- Silk Slide Fade Transitions
- Cartoon Explainer Hero-Split Layouts
- Strict Typography Engine with RTL and Windows Font Compatibility
"""

import os
import json
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE
from pptx.oxml import parse_xml
from pptx.oxml.ns import nsdecls

from theme import (
    SofianColors, SofianFonts, apply_slide_background, add_header,
    apply_typography, format_paragraph, is_primarily_arabic, normalize_quranic_brackets
)
from illustration_engine import default_engine

PROFILE_PATH = os.path.join(os.path.dirname(__file__), "assets", "profile.json")

def load_profile():
    if os.path.exists(PROFILE_PATH):
        with open(PROFILE_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    return {
        "name_ar": "مستر محمد ناصر بدوي",
        "name_en": "MR. Mohamed Nasser Badawy",
        "title_ar_students": "معلم لغة إنجليزية",
        "title_en_students": "English Language Teacher",
        "title_ar_foreigners": "مدرس تجويد ولغة عربية وقرآن للأعاجم",
        "title_en_foreigners": "Quran, Tajweed & Arabic Instructor for Non-Native Speakers",
        "phone": "+201159330109",
        "whatsapp": "+201159330109",
        "whatsapp_link": "https://wa.me/201159330109",
        "facebook_url": "https://www.facebook.com/profile.php?id=100035868428460&mibextid=rS40aB7S9Ucbxw6v",
        "facebook_name": "Mohamed Nasser Badawy",
        "avatar_path": "assets/teacher_avatar_card.png"
    }

def add_click_animation_sequence(slide, shape_ids):
    """
    Injects valid OpenXML timing elements so that the provided shape_ids
    appear sequentially with each click during the PowerPoint presentation.
    """
    if not shape_ids:
        return

    child_pars = []
    bld_graphics = []
    curr_id = 3

    for sp_id in shape_ids:
        c1 = curr_id
        c2 = curr_id + 1
        c3 = curr_id + 2
        curr_id += 3

        click_par = f"""
        <p:par>
          <p:cTn id="{c1}" fill="hold">
            <p:stCondLst><p:cond delay="0"/></p:stCondLst>
            <p:childTnLst>
              <p:par>
                <p:cTn id="{c2}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" nodeType="clickEffect">
                  <p:stCondLst><p:cond delay="0"/></p:stCondLst>
                  <p:childTnLst>
                    <p:set>
                      <p:cBhvr>
                        <p:cTn id="{c3}" dur="1" fill="hold"/>
                        <p:tgtEl><p:spTgt spid="{sp_id}"/></p:tgtEl>
                        <p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst>
                      </p:cBhvr>
                      <p:to><p:strVal val="visible"/></p:to>
                    </p:set>
                  </p:childTnLst>
                </p:cTn>
              </p:par>
            </p:childTnLst>
          </p:cTn>
        </p:par>
        """
        child_pars.append(click_par)
        bld_graphics.append(f'<p:bldGraphic spid="{sp_id}" grpId="0" build="asOne"/>')

    timing_xml = f"""
    <p:timing {nsdecls("p")}>
      <p:tnLst>
        <p:par>
          <p:cTn id="1" dur="indefinite" restart="always" nodeType="tmRoot">
            <p:childTnLst>
              <p:seq concurrent="1" nextAc="seek">
                <p:cTn id="2" dur="indefinite" nodeType="mainSeq">
                  <p:childTnLst>
                    {''.join(child_pars)}
                  </p:childTnLst>
                </p:cTn>
                <p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>
                <p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst>
              </p:seq>
            </p:childTnLst>
          </p:cTn>
        </p:par>
      </p:tnLst>
      <p:bldLst>
        {''.join(bld_graphics)}
      </p:bldLst>
    </p:timing>
    """
    try:
        slide._element.append(parse_xml(timing_xml))
    except Exception as e:
        print(f"[Warning] Animation injection note: {e}")

class PresentationGenerator:
    def __init__(self, track="students", title="Masterclass Lesson", subtitle="Comprehensive Educational Guide"):
        self.prs = Presentation()
        # 16:9 Widescreen dimensions
        self.prs.slide_width = Inches(13.333)
        self.prs.slide_height = Inches(7.5)
        self.blank_layout = self.prs.slide_layouts[6]
        self.track = track
        self.profile = load_profile()
        self.is_foreigners = (track == "foreigners")
        self.title = self.clean_text(title)
        self.subtitle = self.clean_text(subtitle)

    def clean_text(self, text):
        if not text:
            return ""
        return normalize_quranic_brackets(str(text))

    def create_title_slide(self):
        """Slide 1: Hero Cover Slide with Teacher's Portrait & Cinema Branding"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        apply_slide_background(slide, self.prs)

        # Left Column: Portrait Card
        base_dir = os.path.dirname(__file__)
        avatar_file = os.path.join(base_dir, self.profile.get("avatar_path", "assets/teacher_avatar_card.png"))
        if not os.path.exists(avatar_file):
            avatar_file = os.path.join(base_dir, "assets", "teacher_photo.jpg")

        # Outer border frame for teacher photo
        glow = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.85), Inches(1.05), Inches(4.4), Inches(5.4)
        )
        glow.fill.solid()
        glow.fill.fore_color.rgb = SofianColors.CARD_BG
        glow.line.color.rgb = SofianColors.AMBER_GOLD
        glow.line.width = Pt(1.8)

        if os.path.exists(avatar_file):
            slide.shapes.add_picture(avatar_file, Inches(0.9), Inches(1.1), Inches(4.3), Inches(5.3))

        # Right Column Container
        # Track Badge Pill
        badge = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.6), Inches(1.1), Inches(6.9), Inches(0.44)
        )
        badge.fill.solid()
        badge.fill.fore_color.rgb = SofianColors.CARD_SURFACE_ALT
        badge.line.color.rgb = SofianColors.AMBER_GOLD
        badge.line.width = Pt(1.2)

        tf_b = badge.text_frame
        tf_b.margin_left = Inches(0.1)
        tf_b.margin_right = Inches(0.1)
        tf_b.margin_top = Inches(0.04)
        tf_b.margin_bottom = Inches(0.04)
        p = tf_b.paragraphs[0]
        p.text = ""
        r_b = p.add_run()
        if self.is_foreigners:
            r_b.text = "✦  TAJWEED & QURANIC ARABIC MASTERCLASS • NON-NATIVE SPEAKERS"
            apply_typography(r_b, SofianFonts.ENGLISH, 11, SofianColors.AMBER_GOLD, bold=True)
            format_paragraph(p, align=PP_ALIGN.CENTER)
        else:
            r_b.text = "✦  دورة اللغة الإنجليزية المتخصصة • شروحات تفاعلية واحترافية"
            apply_typography(r_b, SofianFonts.ARABIC, 11, SofianColors.AMBER_GOLD, bold=True, is_arabic=True)
            format_paragraph(p, align=PP_ALIGN.CENTER, is_arabic=True)

        # Main Title Box
        title_box = slide.shapes.add_textbox(Inches(5.6), Inches(1.75), Inches(6.9), Inches(2.3))
        tf = title_box.text_frame
        tf.word_wrap = True
        tf.margin_left = Inches(0.05)
        tf.margin_top = Inches(0.05)

        p_title = tf.paragraphs[0]
        p_title.text = ""
        r_title = p_title.add_run()
        r_title.text = self.title
        is_ar_title = self.is_arabic_text(self.title)
        apply_typography(r_title, SofianFonts.ARABIC if is_ar_title else SofianFonts.ENGLISH, 34, SofianColors.TEXT_LIGHT, bold=True, is_arabic=is_ar_title)
        format_paragraph(p_title, align=PP_ALIGN.RIGHT if is_ar_title else PP_ALIGN.LEFT, is_arabic=is_ar_title, space_after_pt=6)

        # Subtitle
        p_sub = tf.add_paragraph()
        r_sub = p_sub.add_run()
        r_sub.text = self.subtitle
        is_ar_sub = self.is_arabic_text(self.subtitle)
        apply_typography(r_sub, SofianFonts.ARABIC if is_ar_sub else SofianFonts.ENGLISH, 16, SofianColors.ELECTRIC_CYAN, is_arabic=is_ar_sub)
        format_paragraph(p_sub, align=PP_ALIGN.RIGHT if is_ar_sub else PP_ALIGN.LEFT, is_arabic=is_ar_sub)

        # Teacher Info Glass Card (100% English for foreigners, 100% Arabic for Arabs)
        teacher_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(5.6), Inches(4.35), Inches(6.9), Inches(2.1)
        )
        teacher_card.fill.solid()
        teacher_card.fill.fore_color.rgb = SofianColors.CARD_BG
        teacher_card.line.color.rgb = SofianColors.CARD_BORDER
        teacher_card.line.width = Pt(1.5)

        tc_tf = teacher_card.text_frame
        tc_tf.word_wrap = True
        tc_tf.margin_left = Inches(0.2)
        tc_tf.margin_right = Inches(0.2)
        tc_tf.margin_top = Inches(0.15)
        tc_tf.margin_bottom = Inches(0.15)

        p1 = tc_tf.paragraphs[0]
        p1.text = ""
        r1 = p1.add_run()
        if self.is_foreigners:
            r1.text = "INSTRUCTOR & PEDAGOGICAL LEAD"
            apply_typography(r1, SofianFonts.ENGLISH, 11, SofianColors.AMBER_GOLD, bold=True)
            format_paragraph(p1, align=PP_ALIGN.LEFT, space_after_pt=4)
        else:
            r1.text = "المُحاضر والمعلم المعتمد"
            apply_typography(r1, SofianFonts.ARABIC, 11, SofianColors.AMBER_GOLD, bold=True, is_arabic=True)
            format_paragraph(p1, align=PP_ALIGN.RIGHT, space_after_pt=4, is_arabic=True)

        p2 = tc_tf.add_paragraph()
        r2 = p2.add_run()
        if self.is_foreigners:
            r2.text = self.profile.get("name_en", "MR. Mohamed Nasser Badawy")
            apply_typography(r2, SofianFonts.ENGLISH, 20, SofianColors.TEXT_LIGHT, bold=True)
            format_paragraph(p2, align=PP_ALIGN.LEFT, space_after_pt=4)
        else:
            r2.text = self.profile.get("name_ar", "مستر محمد ناصر بدوي")
            apply_typography(r2, SofianFonts.ARABIC, 22, SofianColors.TEXT_LIGHT, bold=True, is_arabic=True)
            format_paragraph(p2, align=PP_ALIGN.RIGHT, space_after_pt=4, is_arabic=True)

        p3 = tc_tf.add_paragraph()
        r3 = p3.add_run()
        if self.is_foreigners:
            r3.text = self.profile.get("title_en_foreigners", "Quran, Tajweed & Arabic Instructor for Non-Native Speakers")
            apply_typography(r3, SofianFonts.ENGLISH, 13, SofianColors.TEXT_SECONDARY)
            format_paragraph(p3, align=PP_ALIGN.LEFT)
        else:
            r3.text = self.profile.get("title_ar_students", "معلم لغة إنجليزية متخصص ومحاضر معتمد")
            apply_typography(r3, SofianFonts.ARABIC, 13, SofianColors.TEXT_SECONDARY, is_arabic=True)
            format_paragraph(p3, align=PP_ALIGN.RIGHT, is_arabic=True)

    def create_overview_slide(self, topics):
        """Slide 2: Roadmap Overview with glowing numbered cards"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        apply_slide_background(slide, self.prs)

        if self.is_foreigners:
            title = "Lesson Roadmap & Core Objectives"
            badge = "LESSON ROADMAP"
        else:
            title = "خريطة الدرس والمحاور الرئيسية"
            badge = "محاور الدرس"
        add_header(slide, title, badge, is_arabic=(not self.is_foreigners))

        num_cards = min(len(topics), 4)
        if num_cards == 0:
            return

        card_w = Inches(11.733 / num_cards - 0.25)
        card_h = Inches(4.8)
        top_y = Inches(2.0)
        animated_shape_ids = []

        for i, item in enumerate(topics[:4]):
            left_x = Inches(0.8 + i * (11.733 / num_cards))
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, left_x, top_y, card_w, card_h
            )
            card.fill.solid()
            card.fill.fore_color.rgb = SofianColors.CARD_BG
            card.line.color.rgb = SofianColors.AMBER_GOLD if i == 0 else SofianColors.CARD_BORDER
            card.line.width = Pt(1.5)
            animated_shape_ids.append(card.shape_id)

            tf = card.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.18)
            tf.margin_right = Inches(0.18)
            tf.margin_top = Inches(0.16)
            tf.margin_bottom = Inches(0.16)

            # Number badge
            p_num = tf.paragraphs[0]
            p_num.text = ""
            r_num = p_num.add_run()
            r_num.text = f"0{i+1}"
            apply_typography(r_num, SofianFonts.ENGLISH, 30, SofianColors.AMBER_GOLD, bold=True)
            format_paragraph(p_num, align=PP_ALIGN.LEFT, space_after_pt=6)

            # Topic Title
            p_topic = tf.add_paragraph()
            t_text = self.clean_text(item.get("title", f"Topic {i+1}"))
            is_ar_top = self.is_arabic_text(t_text)
            r_topic = p_topic.add_run()
            r_topic.text = t_text
            apply_typography(r_topic, SofianFonts.ARABIC if is_ar_top else SofianFonts.ENGLISH, 16, SofianColors.TEXT_LIGHT, bold=True, is_arabic=is_ar_top)
            format_paragraph(p_topic, align=PP_ALIGN.RIGHT if is_ar_top else PP_ALIGN.LEFT, is_arabic=is_ar_top, space_after_pt=8)

            # Bullet points
            desc_items = item.get("bullets", [item.get("description", "")])
            for b in desc_items[:4]:
                if not b: continue
                b = self.clean_text(b)
                p_b = tf.add_paragraph()
                is_ar_b = self.is_arabic_text(b)
                r_b = p_b.add_run()
                r_b.text = f"•  {b}"
                apply_typography(r_b, SofianFonts.ARABIC if is_ar_b else SofianFonts.ENGLISH, 13, SofianColors.TEXT_SECONDARY, is_arabic=is_ar_b)
                format_paragraph(p_b, align=PP_ALIGN.RIGHT if is_ar_b else PP_ALIGN.LEFT, is_arabic=is_ar_b, space_after_pt=4)

        add_click_animation_sequence(slide, animated_shape_ids)

    def create_hero_split_slide(self, title, phonetics, cartoon_img_path, definition, bullet_points, example_box):
        """
        Hero Split Slide:
        - Left: Framed Cartoon Explainer Illustration with glowing ambient card.
        - Right: Word/Concept, Phonetics Badge, Clear Definition, Bullet Breakdown & Practical Example Box.
        - Sequenced with click animations and high-contrast typography!
        """
        title = self.clean_text(title)
        phonetics = self.clean_text(phonetics)
        definition = self.clean_text(definition)
        bullet_points = [self.clean_text(pt) for pt in bullet_points if pt]

        slide = self.prs.slides.add_slide(self.blank_layout)
        apply_slide_background(slide, self.prs)
        is_ar = self.is_arabic_text(title)
        # Category badge omitted to keep slide clean and unobstructed
        add_header(slide, title, category_badge=None, is_arabic=is_ar)

        animated_shape_ids = []

        # ================= LEFT COLUMN: CARTOON ILLUSTRATION CARD =================
        left_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(0.8), Inches(2.0), Inches(4.8), Inches(4.9)
        )
        left_card.fill.solid()
        left_card.fill.fore_color.rgb = SofianColors.CARD_BG
        left_card.line.color.rgb = SofianColors.AMBER_GOLD
        left_card.line.width = Pt(1.5)
        animated_shape_ids.append(left_card.shape_id)

        # Insert Cartoon Image if exists, else fallback generate
        if not cartoon_img_path or not os.path.exists(cartoon_img_path):
            cartoon_img_path = default_engine.get_illustration_for_topic(title, definition, self.track)

        if os.path.exists(cartoon_img_path):
            slide.shapes.add_picture(
                cartoon_img_path, Inches(1.0), Inches(2.15), Inches(4.4), Inches(4.55)
            )

        # ================= RIGHT COLUMN: EDUCATIONAL CONTENT =================
        col2_left = Inches(5.9)
        col2_width = Inches(6.633)

        # 1. Phonetics / Category Pill
        if phonetics:
            ph_pill = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, col2_left, Inches(1.95), Inches(3.6), Inches(0.36)
            )
            ph_pill.fill.solid()
            ph_pill.fill.fore_color.rgb = SofianColors.CARD_SURFACE_ALT
            ph_pill.line.color.rgb = SofianColors.ELECTRIC_CYAN
            ph_pill.line.width = Pt(1)

            p_ph = ph_pill.text_frame.paragraphs[0]
            p_ph.text = ""
            r_ph = p_ph.add_run()
            r_ph.text = f"🗣  {phonetics}"
            is_ar_ph = self.is_arabic_text(phonetics)
            apply_typography(r_ph, SofianFonts.ARABIC if is_ar_ph else SofianFonts.ENGLISH, 11, SofianColors.ELECTRIC_CYAN, bold=True, is_arabic=is_ar_ph)
            format_paragraph(p_ph, align=PP_ALIGN.RIGHT if is_ar_ph else PP_ALIGN.LEFT, is_arabic=is_ar_ph)

        # 2. Definition Box (Card)
        def_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, col2_left, Inches(2.42), col2_width, Inches(1.4)
        )
        def_card.fill.solid()
        def_card.fill.fore_color.rgb = SofianColors.CARD_BG
        def_card.line.color.rgb = SofianColors.CARD_BORDER
        def_card.line.width = Pt(1.5)
        animated_shape_ids.append(def_card.shape_id)

        tf_def = def_card.text_frame
        tf_def.word_wrap = True
        tf_def.margin_left = Inches(0.18)
        tf_def.margin_right = Inches(0.18)
        tf_def.margin_top = Inches(0.12)
        tf_def.margin_bottom = Inches(0.12)

        p_dt = tf_def.paragraphs[0]
        p_dt.text = ""
        r_dt = p_dt.add_run()
        if self.is_foreigners:
            r_dt.text = "Meaning & Core Rule:"
            apply_typography(r_dt, SofianFonts.ENGLISH, 11, SofianColors.AMBER_GOLD, bold=True)
            format_paragraph(p_dt, align=PP_ALIGN.LEFT, space_after_pt=3)
        else:
            r_dt.text = "المعنى والقاعدة الأساسية:"
            apply_typography(r_dt, SofianFonts.ARABIC, 11, SofianColors.AMBER_GOLD, bold=True, is_arabic=True)
            format_paragraph(p_dt, align=PP_ALIGN.RIGHT, space_after_pt=3, is_arabic=True)

        p_db = tf_def.add_paragraph()
        is_ar_def = self.is_arabic_text(definition)
        r_db = p_db.add_run()
        r_db.text = definition
        apply_typography(r_db, SofianFonts.ARABIC if is_ar_def else SofianFonts.ENGLISH, 14, SofianColors.TEXT_BODY, bold=False, is_arabic=is_ar_def)
        format_paragraph(p_db, align=PP_ALIGN.RIGHT if is_ar_def else PP_ALIGN.LEFT, is_arabic=is_ar_def)

        # 3. Key Breakdown Bullets Card
        bullets_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, col2_left, Inches(3.95), col2_width, Inches(1.55)
        )
        bullets_card.fill.solid()
        bullets_card.fill.fore_color.rgb = SofianColors.CARD_BG
        bullets_card.line.color.rgb = SofianColors.CARD_BORDER
        bullets_card.line.width = Pt(1.2)
        animated_shape_ids.append(bullets_card.shape_id)

        tf_b = bullets_card.text_frame
        tf_b.word_wrap = True
        tf_b.margin_left = Inches(0.18)
        tf_b.margin_right = Inches(0.18)
        tf_b.margin_top = Inches(0.12)
        tf_b.margin_bottom = Inches(0.12)

        p_bt = tf_b.paragraphs[0]
        p_bt.text = ""
        r_bt = p_bt.add_run()
        if self.is_foreigners:
            r_bt.text = "Key Breakdown & Usage Notes:"
            apply_typography(r_bt, SofianFonts.ENGLISH, 11, SofianColors.ELECTRIC_CYAN, bold=True)
            format_paragraph(p_bt, align=PP_ALIGN.LEFT, space_after_pt=3)
        else:
            r_bt.text = "أهم نقاط التركيز والاستخدام:"
            apply_typography(r_bt, SofianFonts.ARABIC, 11, SofianColors.ELECTRIC_CYAN, bold=True, is_arabic=True)
            format_paragraph(p_bt, align=PP_ALIGN.RIGHT, space_after_pt=3, is_arabic=True)

        for pt in bullet_points[:3]:
            p_item = tf_b.add_paragraph()
            is_ar_pt = self.is_arabic_text(pt)
            r_item = p_item.add_run()
            r_item.text = f"✔  {pt}"
            apply_typography(r_item, SofianFonts.ARABIC if is_ar_pt else SofianFonts.ENGLISH, 13, SofianColors.TEXT_SECONDARY, is_arabic=is_ar_pt)
            format_paragraph(p_item, align=PP_ALIGN.RIGHT if is_ar_pt else PP_ALIGN.LEFT, is_arabic=is_ar_pt, space_after_pt=2)

        # 4. Example Callout Card (Emerald Mint Accent)
        ex_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, col2_left, Inches(5.62), col2_width, Inches(1.28)
        )
        ex_card.fill.solid()
        ex_card.fill.fore_color.rgb = SofianColors.BG_GLOW_TOP
        ex_card.line.color.rgb = SofianColors.EMERALD_MINT
        ex_card.line.width = Pt(1.5)
        animated_shape_ids.append(ex_card.shape_id)

        tf_ex = ex_card.text_frame
        tf_ex.word_wrap = True
        tf_ex.margin_left = Inches(0.18)
        tf_ex.margin_right = Inches(0.18)
        tf_ex.margin_top = Inches(0.10)
        tf_ex.margin_bottom = Inches(0.10)

        p_ext = tf_ex.paragraphs[0]
        p_ext.text = ""
        r_ext = p_ext.add_run()
        if self.is_foreigners:
            r_ext.text = "💡 Practical Example:"
            apply_typography(r_ext, SofianFonts.ENGLISH, 11, SofianColors.EMERALD_MINT, bold=True)
            format_paragraph(p_ext, align=PP_ALIGN.LEFT, space_after_pt=2)
        else:
            r_ext.text = "💡 مثال عملي تطبيقي:"
            apply_typography(r_ext, SofianFonts.ARABIC, 11, SofianColors.EMERALD_MINT, bold=True, is_arabic=True)
            format_paragraph(p_ext, align=PP_ALIGN.RIGHT, space_after_pt=2, is_arabic=True)
        p_exb = tf_ex.add_paragraph()
        raw_ex = example_box.get('example', '')
        clean_ex = self.clean_text(raw_ex.strip('“"\'” '))
        if clean_ex.startswith('«') and clean_ex.endswith('»'):
            ex_text = clean_ex
        elif clean_ex:
            ex_text = f"“ {clean_ex} ”"
        else:
            ex_text = ""
        is_ar_ex = self.is_arabic_text(ex_text)
        r_exb = p_exb.add_run()
        r_exb.text = ex_text
        apply_typography(r_exb, SofianFonts.ARABIC if is_ar_ex else SofianFonts.ENGLISH, 13, SofianColors.TEXT_LIGHT, bold=True, is_arabic=is_ar_ex)
        format_paragraph(p_exb, align=PP_ALIGN.RIGHT if is_ar_ex else PP_ALIGN.LEFT, is_arabic=is_ar_ex)

        add_click_animation_sequence(slide, animated_shape_ids)

    def create_concept_card_slide(self, title, concept_box, explanation_points, example_box, cartoon_img_path=None):
        """Slide Type: Core Concept with Rule Highlight, Breakdown & Example"""
        if not cartoon_img_path:
            cartoon_img_path = default_engine.get_illustration_for_topic(
                title, concept_box.get("text", ""), self.track
            )
        phonetics = concept_box.get("badge", "القاعدة الأساسية")
        definition = concept_box.get("text", "")
        self.create_hero_split_slide(title, phonetics, cartoon_img_path, definition, explanation_points, example_box)

    def create_comparison_table_slide(self, title, headers, rows):
        """Slide Type: Comparison Grid with styled headers and alternating dark rows"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        apply_slide_background(slide, self.prs)
        is_ar = self.is_arabic_text(title)
        badge = "COMPARISON MATRIX" if self.is_foreigners else "مصفوفة المقارنة"
        add_header(slide, title, badge, is_arabic=is_ar)

        num_cols = len(headers)
        num_rows = len(rows) + 1
        table_shape = slide.shapes.add_table(
            num_rows, num_cols, Inches(0.8), Inches(2.0), Inches(11.733), Inches(4.8)
        )
        table = table_shape.table

        # Headers
        for col_idx, h_text in enumerate(headers):
            cell = table.cell(0, col_idx)
            cell.fill.solid()
            cell.fill.fore_color.rgb = SofianColors.CARD_SURFACE_ALT if col_idx % 2 == 0 else SofianColors.BG_GLOW_TOP
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            cell.margin_left = Inches(0.12)
            cell.margin_right = Inches(0.12)
            p = cell.text_frame.paragraphs[0]
            p.text = ""
            r = p.add_run()
            clean_h = self.clean_text(h_text)
            r.text = clean_h
            is_ar_h = self.is_arabic_text(clean_h)
            c_color = SofianColors.AMBER_GOLD if col_idx == 0 else SofianColors.ELECTRIC_CYAN
            apply_typography(r, SofianFonts.ARABIC if is_ar_h else SofianFonts.ENGLISH, 15, c_color, bold=True, is_arabic=is_ar_h)
            format_paragraph(p, align=PP_ALIGN.CENTER, is_arabic=is_ar_h)

        # Rows
        for row_idx, row_data in enumerate(rows):
            for col_idx, cell_value in enumerate(row_data):
                cell = table.cell(row_idx + 1, col_idx)
                cell.fill.solid()
                cell.fill.fore_color.rgb = SofianColors.CARD_BG if row_idx % 2 == 0 else SofianColors.BG_DARK
                cell.vertical_anchor = MSO_ANCHOR.MIDDLE
                cell.margin_left = Inches(0.12)
                cell.margin_right = Inches(0.12)
                p = cell.text_frame.paragraphs[0]
                p.text = ""
                r = p.add_run()
                clean_val = self.clean_text(str(cell_value))
                r.text = clean_val
                is_ar_v = self.is_arabic_text(clean_val)
                apply_typography(r, SofianFonts.ARABIC if is_ar_v else SofianFonts.ENGLISH, 13, SofianColors.TEXT_BODY, is_arabic=is_ar_v)
                format_paragraph(p, align=PP_ALIGN.CENTER, is_arabic=is_ar_v)

    def create_process_steps_slide(self, title, steps):
        """Slide Type: Sequential Process Flow / Step Cards"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        apply_slide_background(slide, self.prs)
        is_ar = self.is_arabic_text(title)
        badge = "STEP-BY-STEP FLOW" if self.is_foreigners else "خطوات التطبيق"
        add_header(slide, title, badge, is_arabic=is_ar)

        n = min(len(steps), 4)
        if n == 0:
            return

        card_w = Inches(11.733 / n - 0.3)
        animated_shape_ids = []

        for i, step in enumerate(steps[:4]):
            left_x = Inches(0.8 + i * (11.733 / n))
            card = slide.shapes.add_shape(
                MSO_SHAPE.ROUNDED_RECTANGLE, left_x, Inches(2.2), card_w, Inches(4.5)
            )
            card.fill.solid()
            card.fill.fore_color.rgb = SofianColors.CARD_BG
            card.line.color.rgb = SofianColors.AMBER_GOLD if i == 0 else SofianColors.CARD_BORDER
            card.line.width = Pt(1.5)
            animated_shape_ids.append(card.shape_id)

            tf = card.text_frame
            tf.word_wrap = True
            tf.margin_left = Inches(0.18)
            tf.margin_right = Inches(0.18)
            tf.margin_top = Inches(0.15)
            tf.margin_bottom = Inches(0.15)

            # Step Badge
            p_step = tf.paragraphs[0]
            p_step.text = ""
            r_step = p_step.add_run()
            if self.is_foreigners:
                r_step.text = f"STEP 0{i+1}"
                apply_typography(r_step, SofianFonts.ENGLISH, 12, SofianColors.AMBER_GOLD, bold=True)
                format_paragraph(p_step, align=PP_ALIGN.LEFT, space_after_pt=4)
            else:
                r_step.text = f"الخطوة 0{i+1}"
                apply_typography(r_step, SofianFonts.ARABIC, 12, SofianColors.AMBER_GOLD, bold=True, is_arabic=True)
                format_paragraph(p_step, align=PP_ALIGN.RIGHT, space_after_pt=4, is_arabic=True)

            # Step Title
            p_title = tf.add_paragraph()
            st_text = self.clean_text(step.get("title", ""))
            is_ar_st = self.is_arabic_text(st_text)
            r_title = p_title.add_run()
            r_title.text = st_text
            apply_typography(r_title, SofianFonts.ARABIC if is_ar_st else SofianFonts.ENGLISH, 16, SofianColors.TEXT_LIGHT, bold=True, is_arabic=is_ar_st)
            format_paragraph(p_title, align=PP_ALIGN.RIGHT if is_ar_st else PP_ALIGN.LEFT, is_arabic=is_ar_st, space_after_pt=6)

            # Step Description
            p_desc = tf.add_paragraph()
            sd_text = self.clean_text(step.get("description", ""))
            is_ar_sd = self.is_arabic_text(sd_text)
            r_desc = p_desc.add_run()
            r_desc.text = sd_text
            apply_typography(r_desc, SofianFonts.ARABIC if is_ar_sd else SofianFonts.ENGLISH, 13, SofianColors.TEXT_SECONDARY, is_arabic=is_ar_sd)
            format_paragraph(p_desc, align=PP_ALIGN.RIGHT if is_ar_sd else PP_ALIGN.LEFT, is_arabic=is_ar_sd, space_after_pt=6)

            # Example or Tip
            if "example" in step:
                p_ex = tf.add_paragraph()
                clean_step_ex = self.clean_text(step['example'])
                ex_t = f"💡 {clean_step_ex}"
                is_ar_ex = self.is_arabic_text(ex_t)
                r_ex = p_ex.add_run()
                r_ex.text = ex_t
                apply_typography(r_ex, SofianFonts.ARABIC if is_ar_ex else SofianFonts.ENGLISH, 12, SofianColors.ELECTRIC_CYAN, is_arabic=is_ar_ex)
                format_paragraph(p_ex, align=PP_ALIGN.RIGHT if is_ar_ex else PP_ALIGN.LEFT, is_arabic=is_ar_ex)

        add_click_animation_sequence(slide, animated_shape_ids)

    def create_closing_slide(self):
        """Slide: Redesigned Executive Closing Hub without overlapping boxes or BiDi collisions"""
        slide = self.prs.slides.add_slide(self.blank_layout)
        apply_slide_background(slide, self.prs)

        # 1. Top Centered Badge Pill
        pill = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(4.3), Inches(0.45), Inches(4.733), Inches(0.38)
        )
        pill.fill.solid()
        pill.fill.fore_color.rgb = SofianColors.CARD_SURFACE_ALT
        pill.line.color.rgb = SofianColors.AMBER_GOLD
        pill.line.width = Pt(1.2)
        p_pill = pill.text_frame.paragraphs[0]
        p_pill.text = ""
        r_pill = p_pill.add_run()
        if self.is_foreigners:
            r_pill.text = "🌟  JAZAKUM ALLAHU KHAYRAN"
            apply_typography(r_pill, SofianFonts.ENGLISH, 10, SofianColors.AMBER_GOLD, bold=True)
            format_paragraph(p_pill, align=PP_ALIGN.CENTER)
        else:
            r_pill.text = "🌟  بالتوفيق والنجاح الدائم"
            apply_typography(r_pill, SofianFonts.ARABIC, 10, SofianColors.AMBER_GOLD, bold=True, is_arabic=True)
            format_paragraph(p_pill, align=PP_ALIGN.CENTER, is_arabic=True)

        # 2. Main Title Box
        title_box = slide.shapes.add_textbox(Inches(1.0), Inches(0.9), Inches(11.333), Inches(0.65))
        tf_t = title_box.text_frame
        p_t = tf_t.paragraphs[0]
        p_t.text = ""
        r_t = p_t.add_run()
        if self.is_foreigners:
            r_t.text = "Jazakum Allahu Khayran • Best Wishes"
            apply_typography(r_t, SofianFonts.ENGLISH, 24, SofianColors.TEXT_LIGHT, bold=True)
            format_paragraph(p_t, align=PP_ALIGN.CENTER)
        else:
            r_t.text = "مع تمنياتي لكم بالتفوق والنجاح المستمر دائمًا"
            apply_typography(r_t, SofianFonts.ARABIC, 24, SofianColors.TEXT_LIGHT, bold=True, is_arabic=True)
            format_paragraph(p_t, align=PP_ALIGN.CENTER, is_arabic=True)

        # 3. Dedicated Hadith / Inspirational Quote Card (Clean, Isolated, Never Overlapping)
        quote_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(1.65), Inches(10.933), Inches(1.4)
        )
        quote_card.fill.solid()
        quote_card.fill.fore_color.rgb = SofianColors.CARD_BG
        quote_card.line.color.rgb = SofianColors.CARD_BORDER
        quote_card.line.width = Pt(1.2)

        tf_q = quote_card.text_frame
        tf_q.word_wrap = True
        tf_q.margin_left = Inches(0.3)
        tf_q.margin_right = Inches(0.3)
        tf_q.margin_top = Inches(0.18)
        tf_q.margin_bottom = Inches(0.18)

        p_q1 = tf_q.paragraphs[0]
        p_q1.text = ""
        r_q1 = p_q1.add_run()
        if self.is_foreigners:
            r_q1.text = "قال رسول الله ﷺ: «خَيْرُكُمْ مَنْ تَعَلَّمَ الْقُرْآنَ وَعَلَّمَهُ»"
            apply_typography(r_q1, SofianFonts.ARABIC, 16, SofianColors.AMBER_GOLD, bold=True, is_arabic=True)
            format_paragraph(p_q1, align=PP_ALIGN.CENTER, space_after_pt=4, is_arabic=True)

            p_q2 = tf_q.add_paragraph()
            r_q2 = p_q2.add_run()
            r_q2.text = "\"The best among you are those who learn the Quran and teach it.\""
            apply_typography(r_q2, SofianFonts.ENGLISH, 13, SofianColors.ELECTRIC_CYAN, italic=True)
            format_paragraph(p_q2, align=PP_ALIGN.CENTER)
        else:
            r_q1.text = "“Continuous effort and daily practice are the keys to true English fluency.”"
            apply_typography(r_q1, SofianFonts.ENGLISH, 16, SofianColors.AMBER_GOLD, bold=True)
            format_paragraph(p_q1, align=PP_ALIGN.CENTER, space_after_pt=4)

            p_q2 = tf_q.add_paragraph()
            r_q2 = p_q2.add_run()
            r_q2.text = "«التعلم والممارسة اليومية المستمرة هما مفتاح طلاقة التحدث بالإنجليزية»"
            apply_typography(r_q2, SofianFonts.ARABIC, 13, SofianColors.ELECTRIC_CYAN, italic=True, is_arabic=True)
            format_paragraph(p_q2, align=PP_ALIGN.CENTER, is_arabic=True)

        # 4. Two Contact Cards (WhatsApp and Facebook) Side-by-Side
        card_w = Inches(5.35)
        card_h = Inches(1.8)
        cards_y = Inches(3.2)

        # 4A. WhatsApp Card
        wa_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), cards_y, card_w, card_h
        )
        wa_card.fill.solid()
        wa_card.fill.fore_color.rgb = SofianColors.BG_GLOW_TOP
        wa_card.line.color.rgb = SofianColors.EMERALD_MINT
        wa_card.line.width = Pt(1.5)

        wa_tf = wa_card.text_frame
        wa_tf.word_wrap = True
        wa_tf.margin_left = Inches(0.2)
        wa_tf.margin_top = Inches(0.18)

        p_wa1 = wa_tf.paragraphs[0]
        p_wa1.text = ""
        r_wa1 = p_wa1.add_run()
        if self.is_foreigners:
            r_wa1.text = "📱 WhatsApp Direct & Inquiries"
            apply_typography(r_wa1, SofianFonts.ENGLISH, 11, SofianColors.EMERALD_MINT, bold=True)
            format_paragraph(p_wa1, align=PP_ALIGN.CENTER, space_after_pt=4)
        else:
            r_wa1.text = "📱 واتساب الحصص والمتابعة المباشرة"
            apply_typography(r_wa1, SofianFonts.ARABIC, 11, SofianColors.EMERALD_MINT, bold=True, is_arabic=True)
            format_paragraph(p_wa1, align=PP_ALIGN.CENTER, space_after_pt=4, is_arabic=True)

        p_wa2 = wa_tf.add_paragraph()
        r_wa2 = p_wa2.add_run()
        r_wa2.text = self.profile["phone"]
        apply_typography(r_wa2, SofianFonts.ENGLISH, 20, SofianColors.TEXT_LIGHT, bold=True)
        format_paragraph(p_wa2, align=PP_ALIGN.CENTER, space_after_pt=3)

        p_wa3 = wa_tf.add_paragraph()
        r_wa3 = p_wa3.add_run()
        if self.is_foreigners:
            r_wa3.text = "For 1-on-1 Sessions & Course Follow-up"
            apply_typography(r_wa3, SofianFonts.ENGLISH, 11, SofianColors.TEXT_SECONDARY)
            format_paragraph(p_wa3, align=PP_ALIGN.CENTER)
        else:
            r_wa3.text = "للتواصل المباشر وحجز الدروس والمراجعات"
            apply_typography(r_wa3, SofianFonts.ARABIC, 11, SofianColors.TEXT_SECONDARY, is_arabic=True)
            format_paragraph(p_wa3, align=PP_ALIGN.CENTER, is_arabic=True)

        # 4B. Facebook Card
        fb_card = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(6.783), cards_y, card_w, card_h
        )
        fb_card.fill.solid()
        fb_card.fill.fore_color.rgb = SofianColors.BG_GLOW_TOP
        fb_card.line.color.rgb = SofianColors.ELECTRIC_CYAN
        fb_card.line.width = Pt(1.5)

        fb_tf = fb_card.text_frame
        fb_tf.word_wrap = True
        fb_tf.margin_left = Inches(0.2)
        fb_tf.margin_top = Inches(0.18)

        p_fb1 = fb_tf.paragraphs[0]
        p_fb1.text = ""
        r_fb1 = p_fb1.add_run()
        if self.is_foreigners:
            r_fb1.text = "🌐 Official Facebook Channel"
            apply_typography(r_fb1, SofianFonts.ENGLISH, 11, SofianColors.ELECTRIC_CYAN, bold=True)
            format_paragraph(p_fb1, align=PP_ALIGN.CENTER, space_after_pt=4)
        else:
            r_fb1.text = "🌐 الصفحة الرسمية للمتابعة"
            apply_typography(r_fb1, SofianFonts.ARABIC, 11, SofianColors.ELECTRIC_CYAN, bold=True, is_arabic=True)
            format_paragraph(p_fb1, align=PP_ALIGN.CENTER, space_after_pt=4, is_arabic=True)

        p_fb2 = fb_tf.add_paragraph()
        r_fb2 = p_fb2.add_run()
        if self.is_foreigners:
            r_fb2.text = self.profile.get("facebook_name", "Mohamed Nasser Badawy")
            apply_typography(r_fb2, SofianFonts.ENGLISH, 18, SofianColors.TEXT_LIGHT, bold=True)
            format_paragraph(p_fb2, align=PP_ALIGN.CENTER, space_after_pt=3)
        else:
            r_fb2.text = self.profile.get("name_ar", "مستر محمد ناصر بدوي")
            apply_typography(r_fb2, SofianFonts.ARABIC, 18, SofianColors.TEXT_LIGHT, bold=True, is_arabic=True)
            format_paragraph(p_fb2, align=PP_ALIGN.CENTER, space_after_pt=3, is_arabic=True)

        p_fb3 = fb_tf.add_paragraph()
        r_fb3 = p_fb3.add_run()
        if self.is_foreigners:
            r_fb3.text = "Follow for Lessons & Live Sessions"
            apply_typography(r_fb3, SofianFonts.ENGLISH, 11, SofianColors.TEXT_SECONDARY)
            format_paragraph(p_fb3, align=PP_ALIGN.CENTER)
        else:
            r_fb3.text = "تابع أحدث الشروحات والتدريبات اليومية"
            apply_typography(r_fb3, SofianFonts.ARABIC, 11, SofianColors.TEXT_SECONDARY, is_arabic=True)
            format_paragraph(p_fb3, align=PP_ALIGN.CENTER, is_arabic=True)

        # 5. Bottom Teacher Signature Banner (Pure English or Pure Arabic)
        sig_banner = slide.shapes.add_shape(
            MSO_SHAPE.ROUNDED_RECTANGLE, Inches(1.2), Inches(5.15), Inches(10.933), Inches(1.6)
        )
        sig_banner.fill.solid()
        sig_banner.fill.fore_color.rgb = SofianColors.CARD_SURFACE_ALT
        sig_banner.line.color.rgb = SofianColors.AMBER_GOLD
        sig_banner.line.width = Pt(1.2)

        sig_tf = sig_banner.text_frame
        sig_tf.word_wrap = True
        sig_tf.margin_left = Inches(0.25)
        sig_tf.margin_top = Inches(0.22)

        p_sig1 = sig_tf.paragraphs[0]
        p_sig1.text = ""
        r_sig1 = p_sig1.add_run()
        if self.is_foreigners:
            r_sig1.text = self.profile.get("name_en", "MR. Mohamed Nasser Badawy")
            apply_typography(r_sig1, SofianFonts.ENGLISH, 20, SofianColors.AMBER_GOLD, bold=True)
            format_paragraph(p_sig1, align=PP_ALIGN.CENTER, space_after_pt=6)

            p_sig2 = sig_tf.add_paragraph()
            r_sig2 = p_sig2.add_run()
            r_sig2.text = self.profile.get("title_en_foreigners", "Quran, Tajweed & Arabic Instructor for Non-Native Speakers")
            apply_typography(r_sig2, SofianFonts.ENGLISH, 13, SofianColors.TEXT_SECONDARY)
            format_paragraph(p_sig2, align=PP_ALIGN.CENTER)
        else:
            r_sig1.text = self.profile.get("name_ar", "مستر محمد ناصر بدوي")
            apply_typography(r_sig1, SofianFonts.ARABIC, 22, SofianColors.AMBER_GOLD, bold=True, is_arabic=True)
            format_paragraph(p_sig1, align=PP_ALIGN.CENTER, space_after_pt=6, is_arabic=True)

            p_sig2 = sig_tf.add_paragraph()
            r_sig2 = p_sig2.add_run()
            r_sig2.text = self.profile.get("title_ar_students", "معلم لغة إنجليزية متخصص ومحاضر معتمد")
            apply_typography(r_sig2, SofianFonts.ARABIC, 14, SofianColors.TEXT_SECONDARY, is_arabic=True)
            format_paragraph(p_sig2, align=PP_ALIGN.CENTER, is_arabic=True)

    def is_arabic_text(self, text):
        return is_primarily_arabic(text)

    def save(self, output_path):
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        self.prs.save(output_path)
        return output_path
