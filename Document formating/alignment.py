import os
import json
import uuid
import platform
import html as _html

from flask import (Blueprint, request, render_template,
                   jsonify, send_file, session, current_app)
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import qn
from werkzeug.utils import secure_filename

IS_WINDOWS = platform.system() == "Windows"
WORD_AVAILABLE = False
if IS_WINDOWS:
    try:
        import pythoncom
        from docx2pdf import convert
        WORD_AVAILABLE = True
    except ImportError:
        pass

alignment_bp = Blueprint(
    "alignment",
    __name__,
    template_folder="templates",
    url_prefix="/alignment"
)

OUTPUT = "outputs/alignment"
os.makedirs(OUTPUT, exist_ok=True)


# ── Session store helpers ─────────────────────────────────────────────────

def _get_path():
    sid = session.get("sid")
    if not sid:
        return None
    return current_app.config["GET_WORKING_PATH"](sid)


def _set_path(path: str):
    sid = session.get("sid")
    if sid:
        current_app.config["SET_WORKING_PATH"](sid, path)


# ── Analyse ───────────────────────────────────────────────────────────────

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
            elements.append({"type": "list",      "index": i})
            types_present.add("list")
        else:
            elements.append({"type": "paragraph", "index": i})
            types_present.add("paragraph")

    return elements, sorted(types_present)


# ── Format ────────────────────────────────────────────────────────────────

def format_docx(input_path, output_path, detected, cfg):
    doc = Document(input_path)

    alignment_map = {
        "left":    WD_ALIGN_PARAGRAPH.LEFT,
        "right":   WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
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


# ── HTML preview fallback ─────────────────────────────────────────────────

def _docx_to_html(docx_path: str) -> str:
    doc = Document(docx_path)
    page_css = (
        '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
        'body{font-family:Arial,sans-serif;margin:0;padding:0;background:#e8e8e8;}'
        '.page{background:white;width:794px;min-height:1123px;margin:30px auto;'
        'padding:72px 80px;box-shadow:0 2px 12px rgba(0,0,0,0.18);box-sizing:border-box;}'
        'h1,h2,h3,h4,h5,h6{margin:0.4em 0 0.2em;}'
        'p{margin:0.3em 0;line-height:1.5;}'
        'ul,ol{margin:0.5em 0 0.5em 1.8em;padding:0;}'
        'li{margin:0.2em 0;line-height:1.5;}'
        '</style></head><body><div class="page">'
    )
    parts = [page_css]
    in_list = False
    for para in doc.paragraphs:
        text = (para.text or "").strip()
        if not text:
            if in_list:
                parts.append("</ul>")
                in_list = False
            parts.append("<p>&nbsp;</p>")
            continue
        is_list = bool(
            para._element.find(qn("w:pPr")) is not None
            and para._element.find(qn("w:pPr")).find(qn("w:numPr")) is not None
        )
        sn = para.style.name.lower() if para.style else ""
        if is_list:
            if not in_list:
                parts.append("<ul>")
                in_list = True
            parts.append(f"<li>{_html.escape(text)}</li>")
        else:
            if in_list:
                parts.append("</ul>")
                in_list = False
            tag = "p"
            for i in range(1, 7):
                if f"heading {i}" in sn:
                    tag = f"h{i}"
                    break
            align = ""
            if para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
                align = ' style="text-align:center"'
            elif para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                align = ' style="text-align:right"'
            elif para.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
                align = ' style="text-align:justify"'
            parts.append(f"<{tag}{align}>{_html.escape(text)}</{tag}>")
    if in_list:
        parts.append("</ul>")
    parts.append("</div></body></html>")
    return "".join(parts)


def _session_out_dir():
    sid = session.get("sid", "default")
    d = os.path.join(OUTPUT, sid)
    os.makedirs(d, exist_ok=True)
    return d


# ── Routes ────────────────────────────────────────────────────────────────

@alignment_bp.route("/")
def home():
    return render_template("alignmentUI.html")


@alignment_bp.route("/analyse", methods=["POST"])
def analyse():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file uploaded yet. Please upload from the home page."}), 400

    _, types = analyse_docx(path)
    return jsonify({"types": types})


@alignment_bp.route("/original", methods=["POST"])
def original():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file uploaded yet. Please upload from the home page."}), 400

    try:
        if IS_WINDOWS and WORD_AVAILABLE:
            try:
                pythoncom.CoInitialize()
                pdf_path = os.path.join(
                    _session_out_dir(), f"orig_{uuid.uuid4().hex}.pdf")
                convert(path, pdf_path)
                pythoncom.CoUninitialize()
                return send_file(pdf_path, mimetype="application/pdf")
            except Exception:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass
        html_content = _docx_to_html(path)
        return html_content, 200, {"Content-Type": "text/html; charset=utf-8"}
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@alignment_bp.route("/preview", methods=["POST"])
def preview():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file uploaded yet. Please upload from the home page."}), 400

    cfg = json.loads(request.form.get("config", "{}"))
    out_docx = os.path.join(
        _session_out_dir(), f"preview_{uuid.uuid4().hex}.docx")
    detected, _ = analyse_docx(path)
    format_docx(path, out_docx, detected, cfg)

    try:
        if IS_WINDOWS and WORD_AVAILABLE:
            pythoncom.CoInitialize()
            pdf_path = os.path.join(
                _session_out_dir(), f"preview_{uuid.uuid4().hex}.pdf")
            convert(out_docx, pdf_path)
            pythoncom.CoUninitialize()
            return send_file(pdf_path, mimetype="application/pdf")
    except Exception:
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass

    html_content = _docx_to_html(out_docx)
    return html_content, 200, {"Content-Type": "text/html; charset=utf-8"}


@alignment_bp.route("/format", methods=["POST"])
def format_doc():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file uploaded yet. Please upload from the home page."}), 400

    cfg = json.loads(request.form.get("config", "{}"))
    out_path = os.path.join(
        _session_out_dir(), f"aligned_{uuid.uuid4().hex}.docx")
    detected, _ = analyse_docx(path)
    format_docx(path, out_path, detected, cfg)

    # Update working file so the next tool in the pipeline uses this output.
    _set_path(out_path)

    return send_file(
        out_path,
        as_attachment=True,
        download_name="formatted.docx",
    )


if __name__ == "__main__":
    from flask import Flask
    _app = Flask(__name__, template_folder="templates")
    _app.secret_key = "dev"
    _app.register_blueprint(alignment_bp)
    _app.run(debug=True, threaded=False)
