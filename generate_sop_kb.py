import argparse
import base64
import datetime as dt
import html
import json
import re
import subprocess
import sys
from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.enum.style import WD_STYLE_TYPE
from bs4 import BeautifulSoup
from urllib import request
from urllib.error import HTTPError, URLError
from urllib.parse import quote

from PIL import Image, ImageDraw, ImageFont

import servicenow_config as sn_config


if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

DEFAULT_MODEL = "llama3.2:latest"
OUTPUT_DIR = Path("generated_docs")
IMAGE_DIR = Path("generated_images")


def slugify(value: str, max_length: int = 80) -> str:
    value = value.strip().lower()
    value = re.sub(r"[^a-z0-9]+", "-", value)
    value = value.strip("-") or "sop"
    return value[:max_length].rstrip("-") or "sop"


def clean_ollama_output(value: str) -> str:
    value = re.sub(r"\x1b\[[0-?]*[ -/]*[@-~]", "", value)
    value = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]", "", value)
    return value.replace("\r", "").strip()


def extract_first_json_object(value: str) -> str | None:
    start = value.find("{")
    if start == -1:
        return None

    depth = 0
    in_string = False
    escaped = False
    for index in range(start, len(value)):
        char = value[index]
        if in_string:
            if escaped:
                escaped = False
            elif char == "\\":
                escaped = True
            elif char == '"':
                in_string = False
            continue

        if char == '"':
            in_string = True
        elif char == "{":
            depth += 1
        elif char == "}":
            depth -= 1
            if depth == 0:
                return value[start:index + 1]
    return None


def run_ollama_prompt(prompt: str, model: str) -> subprocess.CompletedProcess[str]:
    return subprocess.run(
        ["ollama", "run", "--nowordwrap", model],
        input=prompt,
        text=True,
        capture_output=True,
        check=False,
        encoding="utf-8",
    )


def parse_sop_json(content: str) -> dict:
    content = clean_ollama_output(content)
    json_text = extract_first_json_object(content)
    if not json_text:
        json_match = re.search(r'```(?:json)?\s*(\{.*?\})\s*```', content, re.DOTALL)
        if json_match:
            json_text = json_match.group(1)

    if not json_text:
        raise ValueError(f"Could not extract JSON from Ollama response: {content[:200]}")

    return json.loads(json_text)


def repair_sop_json(markdown_or_text: str, model: str) -> dict:
    repair_prompt = f"""
Convert the following SOP draft into ONLY valid JSON.
Do not include markdown, commentary, or code fences.

Required JSON shape:
{{
  "title": "Short title under 90 characters",
  "purpose": "Clear statement of why this SOP exists",
  "scope": ["scope item 1", "scope item 2", "scope item 3"],
  "prerequisites": ["prerequisite 1", "prerequisite 2", "prerequisite 3"],
  "roles_responsibilities": [
    {{"role": "Role name", "responsibility": "What they do"}}
  ],
  "procedure": [
    "Step 1: Detailed action with specific instructions",
    "Step 2: Another step with clear guidance",
    "Step 3: Final step with completion criteria"
  ],
  "validation": ["validation check 1", "validation check 2"],
  "troubleshooting": [
    {{"issue": "Problem description", "cause": "Root cause", "resolution": "How to fix"}}
  ],
  "rollback": ["rollback step 1", "rollback step 2"],
  "references": ["reference 1", "reference 2"]
}}

SOP draft to convert:
<<<SOP_DRAFT
{markdown_or_text[:12000]}
SOP_DRAFT
>>>
""".strip()
    result = run_ollama_prompt(repair_prompt, model)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "Ollama JSON repair failed.")
    try:
        return parse_sop_json(result.stdout)
    except (ValueError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"Could not extract JSON from Ollama response: {clean_ollama_output(markdown_or_text)[:200]}") from exc


def run_ollama_sop(topic: str, model: str) -> dict:
    """Generate SOP content as structured data from LLM"""
    example_title = topic if len(topic) <= 120 else "Short SOP title based on the request"
    prompt = f"""
You are creating structured SOP data for a ServiceNow KB draft.
Return ONLY valid JSON. Do not include markdown, headings, table of contents, or explanatory text.

Use this exact JSON structure:
{{
  "title": "{example_title}",
  "purpose": "Clear statement of why this SOP exists",
  "scope": ["scope item 1", "scope item 2", "scope item 3"],
  "prerequisites": ["prerequisite 1", "prerequisite 2", "prerequisite 3"],
  "roles_responsibilities": [
    {{"role": "Role name", "responsibility": "What they do"}},
    {{"role": "Another role", "responsibility": "Their responsibility"}}
  ],
  "procedure": [
    "Step 1: Detailed action with specific instructions",
    "Step 2: Another step with clear guidance",
    "Step 3: Final step with completion criteria"
  ],
  "validation": ["validation check 1", "validation check 2"],
  "troubleshooting": [
    {{"issue": "Problem description", "cause": "Root cause", "resolution": "How to fix"}},
    {{"issue": "Another issue", "cause": "Another cause", "resolution": "Another resolution"}}
  ],
  "rollback": ["rollback step 1", "rollback step 2"],
  "references": ["reference 1", "reference 2"]
}}

Requirements:
- Use 3-5 items for each list
- Use 3-5 procedure steps
- Keep title short, specific, and under 90 characters
- Make content specific to the source request below
- Be detailed and actionable
- Use professional technical writing style

Source request:
<<<SOURCE
{topic[:18000]}
SOURCE
>>>
""".strip()

    result = run_ollama_prompt(prompt, model)
    if result.returncode != 0:
        raise RuntimeError(result.stderr.strip() or "Ollama generation failed.")

    try:
        return parse_sop_json(result.stdout)
    except (ValueError, json.JSONDecodeError):
        return repair_sop_json(result.stdout, model)


def generate_image_plan_from_sop(sop_data: dict, model: str) -> dict:
    """Generate image plan based on SOP content - LLM chooses the layout dynamically"""
    title = sop_data.get('title', 'SOP')
    purpose = sop_data.get('purpose', '')
    procedure_steps = sop_data.get('procedure', ['Step 1', 'Step 2', 'Step 3'])
    validation_items = sop_data.get('validation', ['Validation 1'])
    roles = sop_data.get('roles_responsibilities', [])
    
    prompt = f"""
Based on this SOP content, determine the BEST visual layout and create a plan:

SOP Title: {title}
Purpose: {purpose}
Procedure Steps: {', '.join(procedure_steps[:4])}
Validation: {', '.join(validation_items[:2])}
Roles: {', '.join([r.get('role', '') for r in roles[:2]])}

Choose layout based on content type:
- "workflow" → sequential technical processes, step-by-step procedures
- "checklist" → audits, inspections, readiness, compliance checks
- "swimlane" → multiple roles/teams, handoffs, approvals, request fulfillment
- "matrix" → troubleshooting, decision criteria, risk assessment

Return ONLY JSON (no markdown):
{{
  "layout": "workflow or checklist or swimlane or matrix",
  "title": "short visual title (max 40 chars)",
  "context": "one sentence explaining the scenario (max 100 chars)",
  "steps": [
    {{"label": "short label (max 25 chars)", "description": "brief description (max 70 chars)"}},
    {{"label": "second step", "description": "description"}},
    {{"label": "third step", "description": "description"}}
  ],
  "evidence": ["evidence 1", "evidence 2", "evidence 3"]
}}

Make steps match the actual procedure from the SOP.
""".strip()

    result = run_ollama_prompt(prompt, model)
    
    if result.returncode != 0:
        return fallback_image_plan(title, procedure_steps)
    
    raw = clean_ollama_output(result.stdout)
    json_text = extract_first_json_object(raw)
    
    if not json_text:
        return fallback_image_plan(title, procedure_steps)
    
    try:
        plan = json.loads(json_text)
        
        valid_layouts = ["workflow", "checklist", "swimlane", "matrix"]
        if plan.get("layout") not in valid_layouts:
            plan["layout"] = "workflow"
        
        plan["steps"] = plan.get("steps", [])[:4]
        while len(plan["steps"]) < 3:
            plan["steps"].append({"label": f"Step {len(plan['steps'])+1}", "description": "Complete this step"})
        
        plan["evidence"] = plan.get("evidence", [])[:3]
        while len(plan["evidence"]) < 3:
            plan["evidence"].append(f"Evidence {len(plan['evidence'])+1}")
        
        return plan
    except json.JSONDecodeError:
        return fallback_image_plan(title, procedure_steps)


def fallback_image_plan(topic: str, steps: list = None) -> dict:
    if steps is None:
        steps = ["Review Request", "Execute Action", "Validate Result"]
    
    return {
        "layout": "workflow",
        "title": f"{topic[:40]} Process",
        "context": f"Process flow for {topic[:100]}",
        "steps": [
            {"label": steps[0][:25], "description": "Confirm details and requirements"},
            {"label": steps[1][:25] if len(steps) > 1 else "Execute", "description": "Perform the approved procedure"},
            {"label": steps[2][:25] if len(steps) > 2 else "Validate", "description": "Check outcome and document evidence"},
        ],
        "evidence": ["Request confirmed", "Action completed", "Result validated"],
    }


def load_font(size: int, bold: bool = False):
    candidates = [
        r"C:\Windows\Fonts\arialbd.ttf" if bold else r"C:\Windows\Fonts\arial.ttf",
        r"C:\Windows\Fonts\segoeuib.ttf" if bold else r"C:\Windows\Fonts\segoeui.ttf",
        "/usr/share/fonts/truetype/liberation/LiberationSans-Bold.ttf" if bold else "/usr/share/fonts/truetype/liberation/LiberationSans-Regular.ttf",
    ]
    for font_path in candidates:
        if Path(font_path).exists():
            return ImageFont.truetype(font_path, size)
    return ImageFont.load_default()


def draw_wrapped_text(draw, xy, text, font, fill, max_width, line_gap=8, max_lines=None):
    words = str(text).split()
    lines = []
    current = ""
    for word in words:
        trial = f"{current} {word}".strip()
        bbox = draw.textbbox((0, 0), trial, font=font)
        if bbox[2] - bbox[0] <= max_width or not current:
            current = trial
        else:
            lines.append(current)
            current = word
    if current:
        lines.append(current)
    if max_lines and len(lines) > max_lines:
        lines = lines[:max_lines]
        lines[-1] = lines[-1].rstrip(" .") + "..."

    x, y = xy
    for line in lines:
        draw.text((x, y), line, font=font, fill=fill)
        bbox = draw.textbbox((x, y), line, font=font)
        y = bbox[3] + line_gap
    return y


def render_workflow_image(draw, width, height, topic, plan, title_font, heading_font, body_font):
    steps = plan.get("steps", [])[:4]
    card_w = 300
    gap = 60
    total_w = len(steps) * card_w + (len(steps) - 1) * gap
    start_x = (width - total_w) // 2
    card_y = 300
    
    palette = [("#dbeafe", "#2563eb"), ("#dcfce7", "#16a34a"), ("#fef3c7", "#d97706"), ("#fce7f3", "#db2777")]
    
    for idx, step in enumerate(steps):
        fill, accent = palette[idx % len(palette)]
        left = start_x + idx * (card_w + gap)
        right = left + card_w
        
        draw.rounded_rectangle((left, card_y, right, card_y + 200), radius=15, fill=fill, outline="#cbd5e1", width=2)
        draw.ellipse((left + 20, card_y + 20, left + 70, card_y + 70), fill=accent)
        draw.text((left + 38, card_y + 32), str(idx + 1), font=load_font(28, bold=True), fill="#ffffff")
        
        label = str(step.get("label", f"Step {idx+1}"))[:25]
        draw_wrapped_text(draw, (left + 90, card_y + 30), label, load_font(22, bold=True), "#0f172a", card_w - 110, max_lines=2)
        
        desc = str(step.get("description", ""))[:70]
        draw_wrapped_text(draw, (left + 20, card_y + 100), desc, load_font(18), "#475569", card_w - 40, max_lines=3)
        
        if idx < len(steps) - 1:
            arrow_x = right + 20
            arrow_y = card_y + 100
            draw.line((arrow_x, arrow_y, arrow_x + gap - 20, arrow_y), fill="#2563eb", width=4)
            draw.polygon([(arrow_x + gap - 20, arrow_y - 10), (arrow_x + gap, arrow_y), (arrow_x + gap - 20, arrow_y + 10)], fill="#2563eb")
    
    evidence_y = 580
    draw.rounded_rectangle((48, evidence_y, width - 48, height - 48), radius=12, fill="#ffffff", outline="#e2e8f0", width=2)
    draw.text((80, evidence_y + 20), "📋 Evidence to Capture", font=load_font(24, bold=True), fill="#0f172a")
    
    evidence_items = plan.get("evidence", [])[:3]
    for idx, item in enumerate(evidence_items):
        y = evidence_y + 70 + idx * 45
        draw.rounded_rectangle((80, y, 110, y + 28), radius=6, fill="#16a34a")
        draw.line((86, y + 14, 92, y + 20, 102, y + 8), fill="#ffffff", width=3)
        draw.text((130, y + 3), str(item)[:50], font=body_font, fill="#334155")


def render_checklist_image(draw, width, height, topic, plan, title_font, heading_font, body_font):
    steps = plan.get("steps", [])[:4]
    left, top = 90, 280
    
    draw.rounded_rectangle((left, top, width - 90, height - 60), radius=18, fill="#ffffff", outline="#cbd5e1", width=2)
    draw.text((left + 38, top + 30), "✅ Completion Checklist", font=load_font(28, bold=True), fill="#0f172a")
    
    for idx, step in enumerate(steps):
        y = top + 100 + idx * 85
        
        draw.rounded_rectangle((left + 38, y, left + 78, y + 40), radius=8, fill="#16a34a")
        draw.line((left + 48, y + 20, left + 58, y + 30, left + 70, y + 12), fill="#ffffff", width=4)
        
        label = str(step.get("label", f"Item {idx + 1}"))[:35]
        draw.text((left + 100, y + 5), label, font=load_font(22, bold=True), fill="#0f172a")
        
        desc = str(step.get("description", ""))[:80]
        draw_wrapped_text(draw, (left + 100, y + 40), desc, load_font(18), "#475569", 1100, max_lines=2)
    
    evidence = plan.get("evidence", [])[:3]
    side_left = width - 380
    evidence_top = top + 100
    draw.rounded_rectangle((side_left, evidence_top, width - 110, height - 100), radius=14, fill="#f0fdf4", outline="#86efac", width=2)
    draw.text((side_left + 25, evidence_top + 20), "📎 Required Evidence", font=load_font(22, bold=True), fill="#166534")
    
    for idx, item in enumerate(evidence):
        draw_wrapped_text(draw, (side_left + 25, evidence_top + 70 + idx * 50), f"• {str(item)[:40]}", load_font(18), "#14532d", 250, max_lines=2)


def render_swimlane_image(draw, width, height, topic, plan, title_font, heading_font, body_font):
    steps = plan.get("steps", [])[:4]
    lanes = ["Requester", "Support Team", "Approver"]
    lane_top = 280
    lane_h = 150
    
    for idx, lane in enumerate(lanes):
        y = lane_top + idx * lane_h
        draw.rectangle((90, y, width - 90, y + lane_h - 15), fill="#ffffff", outline="#cbd5e1", width=2)
        draw.rectangle((90, y, 320, y + lane_h - 15), fill=["#dbeafe", "#dcfce7", "#fef3c7"][idx])
        
        lane_font = load_font(20, bold=True)
        bbox = draw.textbbox((0, 0), lane, font=lane_font)
        text_width = bbox[2] - bbox[0]
        draw.text((115 + (230 - text_width) // 2, y + 60), lane, font=lane_font, fill="#1e293b")
    
    for idx, step in enumerate(steps):
        lane_idx = idx % 3
        x = 350 + (idx % 3) * 290
        y = lane_top + lane_idx * lane_h + 30
        
        draw.rounded_rectangle((x, y, x + 250, y + 95), radius=12, fill="#f8fafc", outline="#94a3b8", width=2)
        draw.ellipse((x + 12, y + 12, x + 42, y + 42), fill="#3b82f6")
        draw.text((x + 22, y + 18), str(idx + 1), font=load_font(18, bold=True), fill="#ffffff")
        
        label = str(step.get("label", f"Step {idx+1}"))[:22]
        draw_wrapped_text(draw, (x + 55, y + 12), label, load_font(18, bold=True), "#0f172a", 180, max_lines=2)
        
        desc = str(step.get("description", ""))[:50]
        draw_wrapped_text(draw, (x + 55, y + 50), desc, load_font(16), "#475569", 180, max_lines=2)
        
        if idx < len(steps) - 1:
            next_lane = (idx + 1) % 3
            if next_lane == lane_idx:
                arrow_x = x + 255
                arrow_y = y + 45
                draw.line((arrow_x, arrow_y, arrow_x + 25, arrow_y), fill="#3b82f6", width=3)
                draw.polygon([(arrow_x + 25, arrow_y - 8), (arrow_x + 35, arrow_y), (arrow_x + 25, arrow_y + 8)], fill="#3b82f6")
    
    evidence = plan.get("evidence", [])[:3]
    draw.text((105, 800), "🔍 Evidence: " + " → ".join(str(e)[:30] for e in evidence), font=load_font(18), fill="#475569")


def render_matrix_image(draw, width, height, topic, plan, title_font, heading_font, body_font):
    steps = plan.get("steps", [])[:4]
    x1, y1 = 80, 280
    cols = [0, 350, 950, width - 120]
    headers = ["Area / Step", "Action / Check", "Evidence / Outcome"]
    
    for col_idx, header in enumerate(headers):
        draw.rectangle((x1 + cols[col_idx], y1, x1 + cols[col_idx + 1], y1 + 55), fill="#1e293b")
        draw.text((x1 + cols[col_idx] + 20, y1 + 15), header, font=load_font(20, bold=True), fill="#ffffff")
    
    evidence = plan.get("evidence", [])
    for idx, step in enumerate(steps):
        y = y1 + 55 + idx * 90
        fill = "#ffffff" if idx % 2 == 0 else "#f8fafc"
        
        for col_idx in range(3):
            draw.rectangle((x1 + cols[col_idx], y, x1 + cols[col_idx + 1], y + 90), fill=fill, outline="#cbd5e1", width=1)
        
        label = str(step.get("label", f"Item {idx+1}"))[:30]
        draw.text((x1 + 20, y + 12), f"{idx + 1}. {label}", font=load_font(19, bold=True), fill="#0f172a")
        
        desc = str(step.get("description", ""))[:80]
        draw_wrapped_text(draw, (x1 + cols[1] + 15, y + 12), desc, load_font(18), "#475569", cols[2] - cols[1] - 30, max_lines=3)
        
        ev = str(evidence[idx % len(evidence)])[:45] if evidence else "Document completion"
        draw_wrapped_text(draw, (x1 + cols[2] + 15, y + 12), f"✓ {ev}", load_font(18), "#16a34a", width - cols[2] - 40, max_lines=3)
    
    draw.text((x1, y1 + 55 + len(steps) * 90 + 20), "💡 Tip: Complete each action and capture evidence before moving to next step", 
              font=load_font(18), fill="#64748b")


def create_process_image(topic: str, plan: dict) -> Path:
    IMAGE_DIR.mkdir(exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    image_path = IMAGE_DIR / f"{stamp}_{slugify(topic)}.png"
    
    width, height = 1600, 940
    image = Image.new("RGB", (width, height), "#f8fafc")
    draw = ImageDraw.Draw(image)
    
    title_font = load_font(42, bold=True)
    heading_font = load_font(28, bold=True)
    body_font = load_font(24)
    
    draw.rectangle((0, 0, width, 110), fill="#1f2937")
    draw.text((48, 32), str(plan.get("title", "SOP Process Guide"))[:50], font=title_font, fill="#ffffff")
    draw.text((48, 145), "📌 Context", font=heading_font, fill="#475569")
    context = str(plan.get("context", topic))[:120]
    draw_wrapped_text(draw, (48, 185), context, body_font, "#0f172a", width - 100)
    
    layout = plan.get("layout", "workflow").lower()
    
    print(f"   🎨 Generating {layout.upper()} layout image...")
    
    if layout == "checklist":
        render_checklist_image(draw, width, height, topic, plan, title_font, heading_font, body_font)
    elif layout == "swimlane":
        render_swimlane_image(draw, width, height, topic, plan, title_font, heading_font, body_font)
    elif layout == "matrix":
        render_matrix_image(draw, width, height, topic, plan, title_font, heading_font, body_font)
    else:
        render_workflow_image(draw, width, height, topic, plan, title_font, heading_font, body_font)
    
    image.save(image_path, "PNG", optimize=True)
    return image_path


def create_document_with_image(sop_data: dict, image_path: Path, image_plan: dict) -> Path:
    """Create a Word document with SOP content and embedded image"""
    OUTPUT_DIR.mkdir(exist_ok=True)
    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    doc_path = OUTPUT_DIR / f"{stamp}_{slugify(sop_data.get('title', 'SOP'))}.docx"
    
    doc = Document()
    
    # Title
    title = doc.add_heading(sop_data.get('title', 'Standard Operating Procedure'), level=1)
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Date
    date_para = doc.add_paragraph(f"Effective Date: {dt.datetime.now().strftime('%B %d, %Y')}")
    date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
    doc.add_paragraph()
    
    # Purpose
    doc.add_heading('1. Purpose', level=2)
    doc.add_paragraph(sop_data.get('purpose', 'To provide clear instructions for this procedure.'))
    
    # Scope
    doc.add_heading('2. Scope', level=2)
    for item in sop_data.get('scope', ['All relevant systems and users']):
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(item)
    
    # Prerequisites
    doc.add_heading('3. Prerequisites', level=2)
    for item in sop_data.get('prerequisites', ['Appropriate access and approvals']):
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(item)
    
    # Roles and Responsibilities
    doc.add_heading('4. Roles and Responsibilities', level=2)
    roles = sop_data.get('roles_responsibilities', [])
    if roles:
        table = doc.add_table(rows=len(roles)+1, cols=2)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Role'
        header_cells[1].text = 'Responsibility'
        
        for idx, role_item in enumerate(roles, 1):
            cells = table.rows[idx].cells
            cells[0].text = role_item.get('role', '')
            cells[1].text = role_item.get('responsibility', '')
    
    # Insert Image
    if image_path and image_path.exists():
        doc.add_heading('5. Process Overview', level=2)
        layout_type = image_plan.get('layout', 'workflow').upper()
        p = doc.add_paragraph()
        p.add_run(f'The following {layout_type} diagram illustrates the key steps in this procedure:').italic = True
        doc.add_picture(str(image_path), width=Inches(6))
        caption = doc.add_paragraph(f"Figure: {image_plan.get('title', 'Process Diagram')}")
        caption.style = 'Intense Quote'
        doc.add_paragraph()
    
    # Procedure
    doc.add_heading('6. Procedure', level=2)
    procedure_steps = sop_data.get('procedure', ['Follow standard operating procedure'])
    for idx, step in enumerate(procedure_steps, 1):
        p = doc.add_paragraph(style='List Number')
        p.add_run(f"{step}")
    
    # Validation
    doc.add_heading('7. Validation', level=2)
    for item in sop_data.get('validation', ['Verify all steps completed successfully']):
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(item)
    
    # Troubleshooting
    doc.add_heading('8. Troubleshooting', level=2)
    troubleshooting = sop_data.get('troubleshooting', [])
    if troubleshooting:
        table = doc.add_table(rows=len(troubleshooting)+1, cols=3)
        table.style = 'Table Grid'
        header_cells = table.rows[0].cells
        header_cells[0].text = 'Issue'
        header_cells[1].text = 'Cause'
        header_cells[2].text = 'Resolution'
        
        for idx, issue_item in enumerate(troubleshooting, 1):
            cells = table.rows[idx].cells
            cells[0].text = issue_item.get('issue', '')
            cells[1].text = issue_item.get('cause', '')
            cells[2].text = issue_item.get('resolution', '')
    
    # Rollback
    doc.add_heading('9. Rollback / Recovery', level=2)
    for item in sop_data.get('rollback', ['Contact support if rollback is needed']):
        p = doc.add_paragraph(style='List Number')
        p.add_run(item)
    
    # References
    doc.add_heading('10. References', level=2)
    for item in sop_data.get('references', ['Related documentation']):
        p = doc.add_paragraph(style='List Bullet')
        p.add_run(item)
    
    doc.save(doc_path)
    return doc_path


def convert_docx_to_html(docx_path: Path, sop_data: dict, image_path: Path = None, image_plan: dict = None) -> str:
    """Build ServiceNow-friendly KB HTML directly from the SOP data."""
    title = html.escape(str(sop_data.get("title", "Standard Operating Procedure")))
    effective_date = dt.datetime.now().strftime("%B %d, %Y")
    layout_type = html.escape(str(image_plan.get("layout", "process")).upper()) if image_plan else "PROCESS"

    def esc(value) -> str:
        return html.escape(str(value or ""))

    def section(number: int, heading: str, body: str) -> str:
        return f"""
<div style="margin:24px 0 0 0;">
  <h2 style="margin:0 0 10px 0;padding:8px 0 8px 12px;border-left:4px solid #2f65d5;color:#172033;font-size:22px;line-height:1.25;">{number}. {esc(heading)}</h2>
  {body}
</div>"""

    def unordered(items: list, fallback: str) -> str:
        values = items or [fallback]
        lis = "".join(f'<li style="margin:6px 0;">{esc(item)}</li>' for item in values)
        return f'<ul style="margin:8px 0 0 24px;padding:0;line-height:1.6;">{lis}</ul>'

    def ordered(items: list, fallback: str) -> str:
        values = items or [fallback]
        lis = "".join(f'<li style="margin:8px 0;">{esc(item)}</li>' for item in values)
        return f'<ol style="margin:8px 0 0 24px;padding:0;line-height:1.6;">{lis}</ol>'

    def table(headers: list[str], rows: list[list[str]]) -> str:
        head = "".join(
            f'<th style="border:1px solid #cfd7e6;background:#eef3f8;color:#172033;padding:10px;text-align:left;vertical-align:top;">{esc(header)}</th>'
            for header in headers
        )
        body_rows = []
        for row in rows:
            cells = "".join(
                f'<td style="border:1px solid #cfd7e6;padding:10px;text-align:left;vertical-align:top;line-height:1.5;">{esc(cell)}</td>'
                for cell in row
            )
            body_rows.append(f"<tr>{cells}</tr>")
        return f'<table style="border-collapse:collapse;width:100%;margin:10px 0 0 0;">' \
               f"<thead><tr>{head}</tr></thead><tbody>{''.join(body_rows)}</tbody></table>"

    roles_rows = [
        [role.get("role", ""), role.get("responsibility", "")]
        for role in sop_data.get("roles_responsibilities", [])
        if isinstance(role, dict)
    ] or [["Support Team", "Owns execution and documentation of the procedure."]]

    troubleshooting_rows = [
        [item.get("issue", ""), item.get("cause", ""), item.get("resolution", "")]
        for item in sop_data.get("troubleshooting", [])
        if isinstance(item, dict)
    ] or [["Issue encountered", "Cause to be confirmed", "Follow standard escalation and recovery steps."]]

    image_html = ""
    if image_path and image_path.exists():
        img_data = base64.b64encode(image_path.read_bytes()).decode("utf-8")
        image_title = esc(image_plan.get("title", "Process Diagram") if image_plan else "Process Diagram")
        image_html = f"""
<div style="margin:12px 0 0 0;text-align:center;">
  <img src="data:image/png;base64,{img_data}" alt="{layout_type} diagram for {title}" style="max-width:100%;height:auto;border:1px solid #cfd7e6;" />
  <div style="margin-top:6px;color:#657086;font-size:12px;font-style:italic;">Figure: {image_title}</div>
</div>"""

    return f"""
<div style="font-family:Arial,Helvetica,sans-serif;color:#172033;line-height:1.55;max-width:1100px;margin:0 auto;">
  <div style="border:1px solid #cfd7e6;border-radius:8px;overflow:hidden;margin:0 0 22px 0;">
    <div style="background:#172033;color:#ffffff;padding:18px 22px;">
      <div style="font-size:12px;text-transform:uppercase;letter-spacing:0.04em;color:#c9d6ea;">Standard Operating Procedure</div>
      <h1 style="margin:5px 0 0 0;font-size:30px;line-height:1.2;color:#ffffff;">{title}</h1>
    </div>
    <div style="background:#f7f9fc;padding:12px 22px;color:#657086;font-size:14px;">
      Effective Date: {effective_date} | Workflow State: Draft | Article Type: HTML KB Draft
    </div>
  </div>

  {section(1, "Purpose", f'<p style="margin:0;line-height:1.6;">{esc(sop_data.get("purpose", "To provide clear instructions for this procedure."))}</p>')}
  {section(2, "Scope", unordered(sop_data.get("scope", []), "All relevant systems, users, and support teams."))}
  {section(3, "Prerequisites", unordered(sop_data.get("prerequisites", []), "Required access, approvals, and supporting information are available."))}
  {section(4, "Roles and Responsibilities", table(["Role", "Responsibility"], roles_rows))}
  {section(5, f"Process Overview ({layout_type} Diagram)", image_html or '<p style="margin:0;">Process diagram was not generated.</p>')}
  {section(6, "Procedure", ordered(sop_data.get("procedure", []), "Follow the approved operating procedure."))}
  {section(7, "Validation", unordered(sop_data.get("validation", []), "Verify the procedure completed successfully."))}
  {section(8, "Troubleshooting", table(["Issue", "Cause", "Resolution"], troubleshooting_rows))}
  {section(9, "Rollback / Recovery", ordered(sop_data.get("rollback", []), "Escalate to the responsible support team for recovery."))}
  {section(10, "References", unordered(sop_data.get("references", []), "Related internal documentation and vendor guidance."))}
</div>
""".strip()


# ServiceNow functions
def get_config_value(name: str, default: str = "") -> str:
    return str(getattr(sn_config, name, default) or default)


def servicenow_request(method: str, endpoint: str, username: str, password: str, payload: bytes | None, content_type: str) -> dict:
    encoded_auth = base64.b64encode(f"{username}:{password}".encode("utf-8")).decode("ascii")
    req = request.Request(
        endpoint,
        data=payload,
        method=method,
        headers={
            "Authorization": f"Basic {encoded_auth}",
            "Accept": "application/json",
            "Content-Type": content_type,
        },
    )
    try:
        opener = request.build_opener(request.ProxyHandler({}))
        with opener.open(req, timeout=60) as response:
            response_body = response.read().decode("utf-8")
    except HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="replace")
        raise RuntimeError(f"ServiceNow returned HTTP {exc.code}: {detail}") from exc
    except URLError as exc:
        raise RuntimeError(f"Could not connect to ServiceNow: {exc.reason}") from exc
    return json.loads(response_body)


def publish_to_servicenow(topic: str, html_content: str) -> dict:
    """Publish HTML content as draft to ServiceNow KB"""
    instance_url = get_config_value("SERVICENOW_INSTANCE_URL").rstrip("/")
    username = get_config_value("SERVICENOW_USERNAME")
    password = get_config_value("SERVICENOW_PASSWORD")
    table = get_config_value("SERVICENOW_KB_TABLE", get_config_value("SERVICENOW_TABLE", "kb_knowledge"))
    kb_sys_id = get_config_value("SERVICENOW_KB_SYS_ID")

    if not all([instance_url, username, password, table, kb_sys_id]):
        raise RuntimeError("Missing ServiceNow configuration in servicenow_config.py")

    endpoint = f"{instance_url}/api/now/table/{table}"
    payload = {
        "short_description": topic[:160],
        "text": html_content,
        "kb_knowledge_base": kb_sys_id,
        "workflow_state": "draft",
        "active": "true",
    }
    
    data = servicenow_request("POST", endpoint, username, password, json.dumps(payload).encode("utf-8"), "application/json")
    return data.get("result", data)


def main() -> int:
    parser = argparse.ArgumentParser(description="Generate SOP with dynamic images and publish to ServiceNow")
    parser.add_argument("topic", help="SOP topic")
    parser.add_argument("--model", default=DEFAULT_MODEL, help=f"Ollama model (default: {DEFAULT_MODEL})")
    parser.add_argument("--publish", action="store_true", help="Publish to ServiceNow as draft")
    args = parser.parse_args()
    
    print(f"📝 Generating SOP for: {args.topic}")
    
    # Step 1: Generate SOP content from LLM
    print("  1/5 Generating SOP content from LLM...")
    sop_data = run_ollama_sop(args.topic, args.model)
    
    # Step 2: Generate image plan based on SOP (LLM chooses layout dynamically)
    print("  2/5 Creating dynamic image plan (LLM chooses layout)...")
    image_plan = generate_image_plan_from_sop(sop_data, args.model)
    print(f"      📐 Selected layout: {image_plan.get('layout', 'workflow').upper()}")
    
    # Step 3: Create process image with dynamic layout
    print("  3/5 Generating process diagram...")
    image_path = create_process_image(sop_data.get('title', args.topic), image_plan)
    
    # Step 4: Create Word document with embedded image
    print("  4/5 Creating Word document with all content...")
    docx_path = create_document_with_image(sop_data, image_path, image_plan)
    
    # Step 5: Convert to HTML with proper formatting
    print("  5/5 Converting to HTML (preserving all formatting)...")
    html_content = convert_docx_to_html(docx_path, sop_data, image_path, image_plan)
    
    # Save HTML locally for inspection
    html_path = OUTPUT_DIR / f"{docx_path.stem}.html"
    html_path.write_text(html_content, encoding='utf-8')
    
    print(f"\n✅ Successfully generated:")
    print(f"   📄 Word Document: {docx_path}")
    print(f"   🖼️  Process Image: {image_path} ({image_plan.get('layout', 'workflow').upper()} layout)")
    print(f"   🌐 HTML File: {html_path}")
    
    # Publish to ServiceNow if requested
    if args.publish:
        print("\n📤 Publishing to ServiceNow...")
        try:
            result = publish_to_servicenow(args.topic, html_content)
            print(f"✅ Successfully published to ServiceNow!")
            print(f"   KB Number: {result.get('number', 'N/A')}")
            print(f"   Sys ID: {result.get('sys_id', 'N/A')}")
            print(f"   Status: {result.get('workflow_state', 'draft')}")
        except Exception as e:
            print(f"❌ Failed to publish to ServiceNow: {e}")
            return 1
    
    return 0


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        raise SystemExit(1)
