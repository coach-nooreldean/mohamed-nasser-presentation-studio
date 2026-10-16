"""
Command Line Interface for PPTX Generation
Example usage:
  python cli.py --title "Past Simple vs Present Perfect" --track students --text "Lesson notes..."
  python cli.py --title "Tajweed Noon Sakinah" --track foreigners --file lesson.pdf
  python cli.py --title "English Video" --track students --youtube "https://youtube.com/watch?v=..."
"""

import argparse
import sys
from generator_pipeline import build_presentation

def main():
    parser = argparse.ArgumentParser(description="Mr. Mohamed Nasser Badawy PPTX Presentation Studio")
    parser.add_argument("--title", required=True, help="Lesson Title")
    parser.add_argument("--track", choices=["students", "foreigners"], default="students",
                        help="Track: 'students' (English) or 'foreigners' (Tajweed & Arabic)")
    parser.add_argument("--file", help="Path to input document (PDF, DOCX, TXT)")
    parser.add_argument("--youtube", help="YouTube video link")
    parser.add_argument("--text", help="Direct text content or outline")
    parser.add_argument("--output", default=None, help="Output filename (.pptx)")
    parser.add_argument("--api-key", default=None, help="Optional Gemini API Key")

    args = parser.parse_args()

    source_type = "text"
    source_content = ""

    if args.file:
        source_type = "file"
        source_content = args.file
    elif args.youtube:
        source_type = "youtube"
        source_content = args.youtube
    elif args.text:
        source_type = "text"
        source_content = args.text
    else:
        print("Error: You must provide one of --file, --youtube, or --text", file=sys.stderr)
        sys.exit(1)

    clean_title = "".join(c for c in args.title if c.isalnum() or c in (" ", "_", "-")).strip()
    output_filename = args.output or f"{clean_title.replace(' ', '_')}.pptx"
    if not output_filename.endswith(".pptx"):
        output_filename += ".pptx"

    print(f"Generating presentation for '{args.title}' [Track: {args.track}]...")
    out_path, blueprint = build_presentation(
        source_type=source_type,
        source_content=source_content,
        track=args.track,
        custom_title=args.title,
        output_filename=output_filename,
        api_key=args.api_key
    )

    print(f"✅ Success! Presentation saved to:\n{out_path}")

if __name__ == "__main__":
    main()
