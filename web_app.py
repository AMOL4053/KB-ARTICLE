import json
import mimetypes
import os
import sys
import uuid
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from urllib.parse import parse_qs, unquote, urlparse

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import generate_sop_kb as sop


ROOT = Path(__file__).resolve().parent
FRONTEND_DIR = ROOT / "frontend"
DRAFTS: dict[str, dict] = {}


def relative_path(path: Path) -> str:
    return str(path.resolve().relative_to(ROOT)).replace("\\", "/")


def add_log(logs: list[dict], message: str, level: str = "info") -> None:
    logs.append({"level": level, "message": message})


def build_sop(topic: str, model: str) -> dict:
    logs: list[dict] = []
    add_log(logs, f"Request received: {topic}", "info")

    add_log(logs, "1/5 Generating SOP content from Ollama...")
    sop_data = sop.run_ollama_sop(topic, model)
    add_log(logs, f"Ollama returned SOP title: {sop_data.get('title', topic)}", "success")

    add_log(logs, "2/5 Creating image plan...")
    image_plan = sop.generate_image_plan_from_sop(sop_data, model)
    layout = image_plan.get("layout", "workflow").upper()
    add_log(logs, f"Selected diagram layout: {layout}", "success")

    add_log(logs, "3/5 Rendering process diagram...")
    image_path = sop.create_process_image(sop_data.get("title", topic), image_plan)
    add_log(logs, f"Process image created: {relative_path(image_path)}", "success")

    add_log(logs, "4/5 Creating Word document...")
    docx_path = sop.create_document_with_image(sop_data, image_path, image_plan)
    add_log(logs, f"Word document created: {relative_path(docx_path)}", "success")

    add_log(logs, "5/5 Preparing ServiceNow KB HTML draft...")
    html_content = sop.convert_docx_to_html(docx_path, sop_data, image_path, image_plan)
    add_log(logs, "HTML draft prepared for preview and ServiceNow.", "success")

    draft_id = uuid.uuid4().hex
    DRAFTS[draft_id] = {
        "topic": topic,
        "html": html_content,
        "title": sop_data.get("title", topic),
    }
    add_log(logs, "Generated DOCX and KB HTML are ready before publishing.", "success")

    return {
        "ok": True,
        "draft_id": draft_id,
        "topic": topic,
        "title": sop_data.get("title", topic),
        "layout": layout,
        "html": html_content,
        "paths": {
            "docx": relative_path(docx_path),
            "image": relative_path(image_path),
        },
        "logs": logs,
    }


def publish_draft(draft_id: str, topic: str = "", html_content: str = "") -> dict:
    draft = DRAFTS.get(draft_id)
    if draft:
        topic = draft["topic"]
        html_content = draft["html"]

    if not topic or not html_content:
        return {"ok": False, "error": "Generate an SOP before saving it as a ServiceNow draft."}

    result = sop.publish_to_servicenow(topic, html_content)
    return {
        "ok": True,
        "service_now": result,
        "logs": [
            {"level": "success", "message": f"ServiceNow draft created. KB Number: {result.get('number', 'N/A')}"},
            {"level": "success", "message": f"Sys ID: {result.get('sys_id', 'N/A')} | Status: {result.get('workflow_state', 'draft')}"},
        ],
    }


class AppHandler(BaseHTTPRequestHandler):
    server_version = "SOPWeb/1.0"

    def send_json(self, payload: dict, status: int = 200) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def send_file(self, path: Path) -> None:
        if not path.exists() or not path.is_file():
            self.send_error(404, "File not found")
            return

        content_type = mimetypes.guess_type(str(path))[0] or "application/octet-stream"
        data = path.read_bytes()
        self.send_response(200)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(data)))
        self.end_headers()
        self.wfile.write(data)

    def do_GET(self) -> None:
        parsed = urlparse(self.path)

        if parsed.path == "/api/health":
            self.send_json(
                {
                    "ok": True,
                    "service": "sop-web",
                    "port": int(os.environ.get("SOP_WEB_PORT", "8090")),
                }
            )
            return

        if parsed.path == "/":
            self.send_file(FRONTEND_DIR / "index.html")
            return

        if parsed.path == "/download":
            params = parse_qs(parsed.query)
            requested = unquote(params.get("file", [""])[0])
            target = (ROOT / requested).resolve()
            try:
                target.relative_to(ROOT)
            except ValueError:
                self.send_error(403, "Invalid file path")
                return
            self.send_file(target)
            return

        static_path = (FRONTEND_DIR / parsed.path.lstrip("/")).resolve()
        try:
            static_path.relative_to(FRONTEND_DIR)
        except ValueError:
            self.send_error(403, "Invalid static file path")
            return
        self.send_file(static_path)

    def do_POST(self) -> None:
        path = urlparse(self.path).path
        if path not in {"/api/generate", "/api/publish"}:
            self.send_error(404, "Not found")
            return

        try:
            length = int(self.headers.get("Content-Length", "0"))
            data = json.loads(self.rfile.read(length).decode("utf-8"))

            if path == "/api/publish":
                draft_id = str(data.get("draft_id", "")).strip()
                topic = str(data.get("topic", "")).strip()
                html_content = str(data.get("html", "")).strip()
                self.send_json(publish_draft(draft_id, topic, html_content))
                return

            topic = str(data.get("topic", "")).strip()
            model = str(data.get("model", sop.DEFAULT_MODEL)).strip() or sop.DEFAULT_MODEL

            if not topic:
                self.send_json({"ok": False, "error": "Please enter the problem or SOP topic."}, 400)
                return

            self.send_json(build_sop(topic, model))
        except Exception as exc:
            self.send_json(
                {"ok": False, "error": str(exc)},
                500,
            )


class ReusableThreadingHTTPServer(ThreadingHTTPServer):
    allow_reuse_address = True
    daemon_threads = True


def main() -> None:
    port = int(os.environ.get("SOP_WEB_PORT", "8090"))
    server = ReusableThreadingHTTPServer(("127.0.0.1", port), AppHandler)
    print(f"SOP web app running at http://127.0.0.1:{port}")
    server.serve_forever()


if __name__ == "__main__":
    main()
