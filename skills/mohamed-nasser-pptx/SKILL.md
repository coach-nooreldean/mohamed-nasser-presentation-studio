---
name: mohamed-nasser-pptx
description: Generates bespoke, cinema-quality PowerPoint presentations for Mr. Mohamed Nasser Badawy (English for students & Tajweed/Quran/Arabic for non-native speakers) matching the Dark Modern Cinema aesthetic with educational cartoon illustrations and native animations.
---

# Mohamed Nasser Badawy - Presentation Generator Skill

This skill generates high-impact, widescreen (16:9) PowerPoint presentations (.pptx) tailored specifically for **Mr. Mohamed Nasser Badawy** (`MR. Mohamed Nasser Badawy` / `مستر محمد ناصر بدوي`).

## Core Instructor Profile
- **Name:** Mr. Mohamed Nasser Badawy / مستر محمد ناصر بدوي
- **Tracks Supported:**
  1. **English for Students (`students`):** English Language Teacher / معلم لغة إنجليزية
  2. **Tajweed & Arabic for Non-Native Speakers (`foreigners`):** Quran, Tajweed & Arabic Instructor for Non-Native Speakers / مدرس تجويد ولغة عربية وقرآن للأعاجم
- **Phone & WhatsApp:** `+201159330109`
- **Facebook:** [Mr. Mohamed Nasser Badawy](https://www.facebook.com/profile.php?id=100035868428460&mibextid=rS40aB7S9Ucbxw6v)
- **Photo Asset:** `/mnt/data2/documents/githup/محمد ناصر/assets/teacher_avatar_card.png`

## Design Aesthetic (Dark Modern Cinema & Cartoon System)
- **Palette:**
  - Deep Backdrop: `#0A1128` (Midnight Navy)
  - Card Surface: Slate Glass (`#162032` / `#1E293B`)
  - Accent / Badges: Warm Amber Gold (`#E5A93C`)
  - Secondary Accent / Phonetics: Soft Electric Cyan (`#4CC9F0`)
  - Examples / Success: Emerald Mint (`#2EC4B6`)
  - Typography: `Cairo` (Arabic) & `Poppins` (English)
- **Slide Structures:**
  - **Slide 1:** Hero Cover with teacher's framed portrait card, track badge pill, lesson title, subtitle, and instructor signature bar.
  - **Slide 2:** Module Roadmap / Overview with glowing numbered cards (`01`, `02`, `03`, `04`) and click entrance animations.
  - **Hero Split Slides (New):** Balanced 16:9 split layout with an educational cartoon explainer illustration card on one side, and word/rule title, phonetics pill, definition, key points, and practical example callout box on the other.
  - **Comparison & Process Slides:** Clean comparison matrices and sequential step-by-step flows.
  - **Last Slide:** Closing Thank You / Jazakum Allahu Khayran hub with Hadith/Quote, WhatsApp card, Facebook link, and tutoring invitation.
- **Animations:** Native PowerPoint fade slide transitions and sequential click entrance animations for progressive disclosure.

## How to Trigger & Generate Presentations

### Option A: Via Python Pipeline Directly
```bash
python3 /mnt/data2/documents/githup/محمد\ ناصر/cli.py \
  --title "Lesson Title" \
  --track students \
  --text "Your notes or outline here"
```

Or from a file (PDF/Word/TXT):
```bash
python3 /mnt/data2/documents/githup/محمد\ ناصر/cli.py \
  --title "Lesson Title" \
  --track foreigners \
  --file /path/to/document.pdf
```

Or from a YouTube video:
```bash
python3 /mnt/data2/documents/githup/محمد\ ناصر/cli.py \
  --title "Lesson Title" \
  --track students \
  --youtube "https://www.youtube.com/watch?v=..."
```

### Option B: Via Local Web Studio
Launch the local web server:
```bash
python3 /mnt/data2/documents/githup/محمد\ ناصر/start_app.py
```
And navigate to `http://127.0.0.1:8505` to use the drag-and-drop web UI.
