"""
Full Pipeline: Content -> Blueprint -> Cartoon Engine -> Cinema PPTX File
"""

import os
from pptx_generator import PresentationGenerator
from ai_processor import ContentAnalyzer, extract_text_from_file, extract_text_from_youtube
from illustration_engine import default_engine

def build_presentation(
    source_type="text",
    source_content="",
    track="students",
    custom_title="",
    output_filename="lesson_presentation.pptx",
    api_key=None
):
    """
    source_type: 'text', 'file', 'youtube'
    track: 'students' (English) or 'foreigners' (Tajweed & Arabic)
    """
    # 1. Extract content
    raw_text = ""
    if source_type == "file":
        raw_text = extract_text_from_file(source_content)
    elif source_type == "youtube":
        raw_text = extract_text_from_youtube(source_content)
        if not custom_title:
            try:
                import urllib.request
                import json
                req = urllib.request.urlopen(f"https://noembed.com/embed?url={source_content}", timeout=5)
                info = json.loads(req.read().decode())
                fetched_title = info.get("title", "")
                if fetched_title:
                    custom_title = fetched_title
            except Exception:
                pass
    else:
        raw_text = source_content or "Educational Masterclass Content"

    # Auto-detect track if appropriate
    low_check = (raw_text + " " + (custom_title or "")).lower()
    if any(k in low_check for k in ["idgham", "إدغام", "tajweed", "تجويد", "noon sakinah", "tanween", "yarmaloon", "يرملون", "izhar", "إظهار", "quran", "قرآن"]):
        track = "foreigners"

    # 2. Analyze & build blueprint
    analyzer = ContentAnalyzer(api_key=api_key)
    blueprint = analyzer.process(raw_text, track=track, custom_title=custom_title)

    effective_track = blueprint.get("track", track)

    # 3. Instantiate generator
    gen = PresentationGenerator(
        track=effective_track,
        title=blueprint.get("title", custom_title or "Masterclass Lesson"),
        subtitle=blueprint.get("subtitle", "Comprehensive Educational Guide")
    )

    # 4. Assemble Slides
    # Slide 1: Hero Cover with Teacher Photo
    gen.create_title_slide()

    # Slide 2: Roadmap Overview
    roadmap_items = blueprint.get("roadmap", [])
    if roadmap_items:
        gen.create_overview_slide(roadmap_items)

    # Slide 3..N: Content Slides
    for slide_data in blueprint.get("slides", []):
        s_type = slide_data.get("type")
        s_title = slide_data.get("title", "Lesson Module")

        if s_type == "hero_split":
            topic = slide_data.get("title", "")
            desc = slide_data.get("illustration_topic", slide_data.get("definition", ""))
            cartoon_img = default_engine.get_illustration_for_topic(topic, desc, category=track)
            gen.create_hero_split_slide(
                s_title,
                slide_data.get("phonetics", ""),
                cartoon_img,
                slide_data.get("definition", ""),
                slide_data.get("bullet_points", []),
                slide_data.get("example_box", {})
            )
        elif s_type == "concept":
            gen.create_concept_card_slide(
                s_title,
                slide_data.get("concept_box", {}),
                slide_data.get("explanation_points", []),
                slide_data.get("example_box", {})
            )
        elif s_type == "comparison":
            gen.create_comparison_table_slide(
                s_title,
                slide_data.get("headers", ["Aspect", "Item A", "Item B"]),
                slide_data.get("rows", [])
            )
        elif s_type == "steps":
            gen.create_process_steps_slide(
                s_title,
                slide_data.get("steps", [])
            )

    # Slide Last: Closing Thank You & Contact Hub
    gen.create_closing_slide()

    # 5. Save output
    output_dir = os.path.join(os.path.dirname(__file__), "output")
    os.makedirs(output_dir, exist_ok=True)
    out_path = os.path.join(output_dir, output_filename)
    gen.save(out_path)
    return out_path, blueprint

if __name__ == "__main__":
    print("Testing pipeline generation for English track with cartoon explainers...")
    out1, b1 = build_presentation(
        source_type="text",
        source_content="Present Perfect vs Past Simple explanation with rules and examples.",
        track="students",
        custom_title="Mastering English Verb Tenses & Phrasal Verbs",
        output_filename="english_masterclass_v2.pptx"
    )
    print(f"Generated English presentation: {out1}")

    print("Testing pipeline generation for Foreigners Tajweed track...")
    out2, b2 = build_presentation(
        source_type="text",
        source_content="Rules of Noon Sakinah and Tanween: Izhar and Idgham.",
        track="foreigners",
        custom_title="Tajweed Masterclass: Rules of Noon Sakinah & Tanween",
        output_filename="tajweed_masterclass_v2.pptx"
    )
    print(f"Generated Tajweed presentation: {out2}")
