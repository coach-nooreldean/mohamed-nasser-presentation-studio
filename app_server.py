"""
Local Web Server & API for Mr. Mohamed Nasser Badawy Presentation Studio
Zero external web framework dependencies (runs on standard library http.server).
Fully compatible with Python 3.10 through 3.13+ (no deprecated cgi module).
"""

import os
import sys
import json
import urllib.parse
from http.server import HTTPServer, SimpleHTTPRequestHandler
import webbrowser
import threading
import time
import email
from email.policy import default

from generator_pipeline import build_presentation

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
WEB_DIR = os.path.join(BASE_DIR, "web")
OUTPUT_DIR = os.path.join(BASE_DIR, "output")
UPLOADS_DIR = os.path.join(BASE_DIR, "uploads")
os.makedirs(OUTPUT_DIR, exist_ok=True)
os.makedirs(UPLOADS_DIR, exist_ok=True)

class StudioHTTPHandler(SimpleHTTPRequestHandler):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, directory=WEB_DIR, **kwargs)

    def do_GET(self):
        # Serve generated presentations for download
        if self.path.startswith("/download/"):
            filename = urllib.parse.unquote(self.path[len("/download/"):])
            filepath = os.path.join(OUTPUT_DIR, filename)
            if os.path.exists(filepath):
                self.send_response(200)
                self.send_header("Content-Type", "application/vnd.openxmlformats-officedocument.presentationml.presentation")
                self.send_header("Content-Disposition", f'attachment; filename="{filename}"')
                self.send_header("Content-Length", str(os.path.getsize(filepath)))
                self.end_headers()
                with open(filepath, "rb") as f:
                    self.wfile.write(f.read())
                return
            else:
                self.send_error(404, "File not found")
                return

        return super().do_GET()

    def do_POST(self):
        if self.path == "/api/generate":
            self.handle_generate()
        else:
            self.send_error(404, "Unknown API endpoint")

    def handle_generate(self):
        try:
            content_type = self.headers.get("Content-Type", "")
            content_length = int(self.headers.get("Content-Length", 0))
            raw_body = self.rfile.read(content_length)

            form_fields = {}
            uploaded_file = None

            if content_type.startswith("multipart/form-data"):
                # Use standard email parser for multipart
                header_bytes = f"Content-Type: {content_type}\r\n\r\n".encode("utf-8")
                msg = email.message_from_bytes(header_bytes + raw_body, policy=default)

                for part in msg.iter_parts():
                    cd = part.get("Content-Disposition", "")
                    name = part.get_param("name", header="content-disposition")
                    filename = part.get_filename()

                    if filename:
                        uploaded_file = {
                            "filename": filename,
                            "data": part.get_payload(decode=True)
                        }
                    elif name:
                        payload = part.get_payload(decode=True)
                        form_fields[name] = payload.decode("utf-8", errors="ignore") if payload else ""
            else:
                # Regular form encoded
                parsed = urllib.parse.parse_qs(raw_body.decode("utf-8", errors="ignore"))
                form_fields = {k: v[0] for k, v in parsed.items()}

            track = form_fields.get("track", "students")
            title = form_fields.get("title", "Educational Masterclass")
            source_type = form_fields.get("source_type", "text")
            api_key = form_fields.get("api_key", "").strip() or None

            source_content = ""
            if source_type == "file":
                if not uploaded_file:
                    self.send_json_response(400, {"success": False, "error": "No file uploaded"})
                    return
                safe_name = os.path.basename(uploaded_file["filename"])
                upload_path = os.path.join(UPLOADS_DIR, safe_name)
                with open(upload_path, "wb") as f:
                    f.write(uploaded_file["data"])
                source_content = upload_path

            elif source_type == "youtube":
                source_content = form_fields.get("youtube_url", "").strip() or form_fields.get("source_content", "").strip()
                if not source_content:
                    self.send_json_response(400, {"success": False, "error": "Missing YouTube URL"})
                    return
            else:
                source_content = form_fields.get("text_content", "").strip() or form_fields.get("source_content", "").strip()
                if not source_content:
                    self.send_json_response(400, {"success": False, "error": "Missing text content"})
                    return

            clean_title = "".join(c for c in title if c.isalnum() or c in (" ", "_", "-")).strip()
            if not clean_title:
                clean_title = "Lesson_Presentation"
            output_filename = f"{clean_title.replace(' ', '_')}.pptx"

            out_path, blueprint = build_presentation(
                source_type=source_type,
                source_content=source_content,
                track=track,
                custom_title=title,
                output_filename=output_filename,
                api_key=api_key
            )

            slides_count = len(blueprint.get("slides", [])) + 3
            self.send_json_response(200, {
                "success": True,
                "filename": output_filename,
                "slides_count": slides_count,
                "download_url": f"/download/{urllib.parse.quote(output_filename)}",
                "title": blueprint.get("title", title),
                "track": track
            })

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_json_response(500, {
                "success": False,
                "error": str(e)
            })

    def send_json_response(self, code, data):
        response_bytes = json.dumps(data, ensure_ascii=False).encode("utf-8")
        self.send_response(code)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(response_bytes)))
        self.end_headers()
        self.wfile.write(response_bytes)

def start_server(port=None, open_browser=True, host=None):
    env_port = os.environ.get("PORT")
    if port is None:
        port = int(env_port) if env_port else 8505
    else:
        port = int(port)
        
    if host is None:
        host = "0.0.0.0" if env_port else "127.0.0.1"

    server_address = (host, port)
    httpd = HTTPServer(server_address, StudioHTTPHandler)
    display_host = "localhost" if host in ("0.0.0.0", "127.0.0.1") else host
    url = f"http://{display_host}:{port}"

    print(f"\n=======================================================")
    print(f"🎬 Mohamed Nasser Presentation Studio running at:")
    print(f"👉 {url}")
    print(f"=======================================================\n")

    if open_browser and not env_port:
        def _open():
            time.sleep(1.2)
            try:
                webbrowser.open(url)
            except Exception:
                pass
        threading.Thread(target=_open, daemon=True).start()

    httpd.serve_forever()

def run_server(port=8505, host=None):
    start_server(port=port, open_browser=False, host=host)

if __name__ == "__main__":
    port = None
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            pass
    start_server(port=port, open_browser=True)

