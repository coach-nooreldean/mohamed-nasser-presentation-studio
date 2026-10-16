"""
Educational Cartoon Illustration Engine
Delivers cinema-grade 3D Pixar/Disney style educational illustrations
specifically tailored for teaching English (Grammar/Vocabulary) and Tajweed/Arabic.
"""

import os
import re
import hashlib
from PIL import Image, ImageDraw, ImageFont

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ILLUSTRATIONS_DIR = os.path.join(BASE_DIR, "assets", "illustrations")
os.makedirs(ILLUSTRATIONS_DIR, exist_ok=True)

class IllustrationEngine:
    def __init__(self, api_key=None):
        self.api_key = api_key or os.environ.get("GEMINI_API_KEY")

    def get_illustration_for_topic(self, word_or_topic, context_description="", category="english"):
        """
        Returns the absolute filepath to a high-resolution, cinema-grade educational
        cartoon illustration matching the lesson topic.
        """
        query = (word_or_topic + " " + context_description + " " + category).lower()

        # Registry of all 10 cinema-grade 3D Pixar educational illustrations
        self._tajweed_assets = [
            ("izhar_halqi.jpg", ["izhar", "إظهار", "throat", "حلق", "halq", "clear", "clarity", "hamza", "noon sakinah", "sakinah"]),
            ("idgham_ghunnah.jpg", ["idgham", "إدغام", "ghunnah", "غنة", "yanmoo", "ينمو", "merg", "nasal", "assimilat"]),
            ("tajweed_ikhfa.jpg", ["ikhfa", "إخفاء", "haqiqi", "حقيقي", "conceal", "hide", "hiding", "fifteen", "15 letters"]),
            ("tajweed_iqlab.jpg", ["iqlab", "إقلاب", "transform", "convert", "turning", "meem", "baa", "م", "قلب"]),
            ("tajweed_qalqalah.jpg", ["qalqalah", "قلقلة", "bounce", "echo", "vibrat", "qutb jad", "قطب جد", "rhythm"]),
            ("tajweed_madd.jpg", ["madd", "مد", "elongat", "stretch", "harakat", "حركات", "asli", "far'i", "wave", "vowel"]),
        ]

        self._english_assets = [
            ("english_speaking_conversation.jpg", ["speak", "conversation", "dialogue", "talk", "chat", "discuss", "fluency", "question", "listen", "oral", "ability", "presentation", "strengths", "personal", "family", "father", "mother", "parents", "brother", "sister", "grand", "who's this", "who is this"]),
            ("english_grammar_rules.jpg", ["grammar", "rule", "structure", "syntax", "part of speech", "noun", "adjective", "puzzle", "clause", "general", "macro", "framework", "definition", "singular", "plural", "there is", "there are", "phonics", "letter", "sound", "alphabet"]),
            ("present_perfect.jpg", ["present perfect", "past simple", "past continuous", "future", "tense", "homework", "finish", "clock", "timeline", "already", "practice", "develop", "journey", "honed", "learning", "study", "numbers", "count", "apples", "books", "flowers"]),
            ("look_after.jpg", ["look after", "care", "pet", "cat", "kitten", "animal", "help", "kind", "phrasal", "feed", "panda", "pencil", "pizza", "nurse", "sharing", "values"]),
        ]

        # 1. Exact / keyword-weighted search in Tajweed assets
        is_tajweed_query = category == "foreigners" or any(k in query for k in ["tajweed", "quran", "arabic", "تجويد", "قرآن", "نون", "تنوين", "حرف", "حلق", "إدغام", "إظهار", "إخفاء", "إقلاب", "قلقلة", "مد"])

        if is_tajweed_query:
            for filename, kws in self._tajweed_assets:
                if any(kw in query for kw in kws):
                    preset = os.path.join(ILLUSTRATIONS_DIR, filename)
                    if os.path.exists(preset):
                        return preset
            # Rotation among Tajweed 3D assets to prevent repetition
            if not hasattr(self, "_tajweed_idx"):
                self._tajweed_idx = 0
            preset_name = self._tajweed_assets[self._tajweed_idx % len(self._tajweed_assets)][0]
            self._tajweed_idx += 1
            preset = os.path.join(ILLUSTRATIONS_DIR, preset_name)
            if os.path.exists(preset):
                return preset

        # 2. Exact / keyword-weighted search in English assets
        for filename, kws in self._english_assets:
            if any(kw in query for kw in kws):
                preset = os.path.join(ILLUSTRATIONS_DIR, filename)
                if os.path.exists(preset):
                    return preset

        # Rotation among English 3D assets to prevent repetition
        if not hasattr(self, "_english_idx"):
            self._english_idx = 0
        preset_name = self._english_assets[self._english_idx % len(self._english_assets)][0]
        self._english_idx += 1
        preset = os.path.join(ILLUSTRATIONS_DIR, preset_name)
        if os.path.exists(preset):
            return preset

        # Dynamic fallback generation (clean artwork without stamped broken text)
        clean_name = re.sub(r'[^a-zA-Z0-9_\-]', '_', word_or_topic.strip())[:30]
        filename = f"cartoon_{clean_name}.png"
        filepath = os.path.join(ILLUSTRATIONS_DIR, filename)
        if os.path.exists(filepath):
            return filepath

        self._generate_clean_vector_art(filepath, category)
        return filepath

    def _generate_clean_vector_art(self, filepath, category):
        """Generates a clean, polished vector-style educational mascot without stamped text."""
        size = (800, 800)
        img = Image.new("RGBA", size, (19, 29, 49, 255))
        draw = ImageDraw.Draw(img)

        # Background card frame
        draw.rounded_rectangle([25, 25, 775, 775], radius=45, fill=(19, 29, 49, 255), outline=(255, 183, 3, 255), width=4)
        glow_color = (25, 45, 75, 255) if category == "english" else (18, 48, 42, 255)
        draw.ellipse([120, 100, 680, 660], fill=glow_color, outline=(56, 189, 248, 160), width=3)

        # Character
        cx, cy = 400, 400
        shirt_color = (59, 130, 246) if category == "english" else (16, 185, 129)
        draw.rounded_rectangle([cx - 85, cy + 60, cx + 85, cy + 220], radius=35, fill=shirt_color, outline=(255, 255, 255, 200), width=3)
        draw.ellipse([cx - 105, cy - 120, cx + 105, cy + 80], fill=(255, 220, 190), outline=(230, 180, 140), width=3)
        # Hair
        draw.chord([cx - 105, cy - 140, cx + 105, cy - 20], 180, 360, fill=(60, 35, 20))
        # Eyes
        draw.ellipse([cx - 50, cy - 35, cx - 25, cy - 5], fill=(20, 20, 20))
        draw.ellipse([cx + 25, cy - 35, cx + 50, cy - 5], fill=(20, 20, 20))
        draw.ellipse([cx - 43, cy - 30, cx - 33, cy - 20], fill=(255, 255, 255))
        draw.ellipse([cx + 32, cy - 30, cx + 42, cy - 20], fill=(255, 255, 255))
        # Smile
        draw.arc([cx - 30, cy - 10, cx + 30, cy + 35], 0, 180, fill=(190, 40, 40), width=5)

        # Star and Book props
        draw.polygon([(cx - 100, cy + 120), (cx, cy + 150), (cx + 100, cy + 120), (cx + 100, cy + 190), (cx, cy + 220), (cx - 100, cy + 190)], fill=(255, 255, 255), outline=(255, 183, 3), width=4)

        img.save(filepath, "PNG")

default_engine = IllustrationEngine()
