from flask import Blueprint, request, render_template, jsonify, send_file
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import os
import json
import uuid
import pythoncom
from werkzeug.utils import secure_filename

# ── Blueprint — all routes live under /alignment/ ──────────────────────────
alignment_bp = Blueprint(
    "alignment",
    __name__,
    template_folder="templates",
    url_prefix="/alignment"
)

UPLOAD = "uploads/alignment"
OUTPUT = "outputs/alignment"
os.makedirs(UPLOAD, exist_ok=True)
os.makedirs(OUTPUT, exist_ok=True)


# ---------------- ANALYSE ----------------

def analyse_docx(path):
    doc = Document(path)
    elements = []
    types_present = set()

    for i, p in enumerate(doc.paragraphs):
        text = p.text.strip()
        if not text:
            continue
        is_list = bool(p._element.xpath('.//w:numPr'))
        if is_list:
            elements.append({"type": "list", "index": i})
            types_present.add("list")
        else:
            elements.append({"type": "paragraph", "index": i})
            types_present.add("paragraph")

    return elements, sorted(types_present)


# ---------------- FORMAT ----------------

def format_docx(input_path, output_path, detected, cfg):
    doc = Document(input_path)

    alignment_map = {
        "left":    WD_ALIGN_PARAGRAPH.LEFT,
        "right":   WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
    }

    idx_map = {d["index"]: d["type"] for d in detected}

    for i, p in enumerate(doc.paragraphs):
        etype = idx_map.get(i)
        if not etype:
            continue

        pf = p.paragraph_format

        align_value = cfg.get(f"{etype}_alignment", "")
        if align_value:
            p.alignment = alignment_map.get(align_value)

        spacing_value = cfg.get(f"{etype}_spacing", "")
        if spacing_value:
            pf.line_spacing = float(spacing_value)

        if etype == "paragraph":
            options = cfg.get("paragraph_spacing_options", [])
            if "add_before" in options:
                pf.space_before = Pt(12)
            if "add_after" in options:
                pf.space_after = Pt(12)
            if "remove_before" in options:
                pf.space_before = Pt(0)
            if "remove_after" in options:
                pf.space_after = Pt(0)

    doc.save(output_path)


# ---------------- ROUTES ----------------

@alignment_bp.route("/")
def home():
    return render_template("alignmentUI.html")


@alignment_bp.route("/analyse", methods=["POST"])
def analyse():
    f = request.files["file"]
    ext = os.path.splitext(secure_filename(f.filename))[1]
    path = os.path.join(UPLOAD, f"{uuid.uuid4()}{ext}")
    f.save(path)
    detected, types = analyse_docx(path)
    return jsonify({"types": types})


@alignment_bp.route("/preview", methods=["POST"])
def preview():
    pythoncom.CoInitialize()
    f = request.files["file"]
    cfg = json.loads(request.form["config"])

    ext = os.path.splitext(secure_filename(f.filename))[1]
    in_path = os.path.join(UPLOAD, f"{uuid.uuid4()}{ext}")
    f.save(in_path)

    out_docx = os.path.join(OUTPUT, f"formatted_{uuid.uuid4()}.docx")
    detected, _ = analyse_docx(in_path)
    format_docx(in_path, out_docx, detected, cfg)

    try:
        from docx2pdf import convert
        pdf_path = os.path.join(OUTPUT, f"preview_{uuid.uuid4()}.pdf")
        convert(out_docx, pdf_path)
        pythoncom.CoUninitialize()
        return send_file(pdf_path)
    except Exception as e:
        pythoncom.CoUninitialize()
        return jsonify({"error": str(e)}), 500


@alignment_bp.route("/format", methods=["POST"])
def format_doc():
    f = request.files["file"]
    cfg = json.loads(request.form["config"])

    ext = os.path.splitext(secure_filename(f.filename))[1]
    in_path = os.path.join(UPLOAD, f"{uuid.uuid4()}{ext}")
    f.save(in_path)

    out_path = os.path.join(OUTPUT, f"formatted_{uuid.uuid4()}.docx")
    detected, _ = analyse_docx(in_path)
    format_docx(in_path, out_path, detected, cfg)

    return send_file(out_path, as_attachment=True,
                     download_name="formatted.docx")


# ── Allows running alignment.py standalone for testing ─────────────────────
if __name__ == "__main__":
    from flask import Flask
    _app = Flask(__name__, template_folder="templates")
    _app.register_blueprint(alignment_bp)
    _app.run(debug=True, threaded=False)
