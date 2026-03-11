import re as _re
from flask import Blueprint, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from docx import Document
from textblob import TextBlob
import language_tool_python
import os
import json
import io
import uuid

grammar_bp = Blueprint(
    "grammar",
    __name__,
    template_folder="templates",
    url_prefix="/grammar"
)

UPLOAD = "uploads/grammar"
os.makedirs(UPLOAD, exist_ok=True)

tool = language_tool_python.LanguageTool('en-US')


def extract_sentences_from_docx(path):
    """
    Extract sentences from both normal paragraphs AND table cells.
    Each unit gets a unique paragraph_index so corrections map back correctly.
    Table cells carry a human-readable location_label shown in the UI.
    """
    doc = Document(path)
    sentences = []
    paragraph_index = 0

    # 1. Normal body paragraphs
    for para in doc.paragraphs:
        paragraph_index += 1
        if para.text.strip() == "":
            continue
        blob = TextBlob(para.text)
        for sentence in blob.sentences:
            sentences.append({
                "paragraph": paragraph_index,
                "location_label": f"\u00b6{paragraph_index}",
                "source": "paragraph",
                "sentence": str(sentence)
            })

    # 2. Table cells
    for t_idx, table in enumerate(doc.tables, start=1):
        for r_idx, row in enumerate(table.rows, start=1):
            for c_idx, cell in enumerate(row.cells, start=1):
                cell_text = cell.text.strip()
                if not cell_text:
                    continue
                paragraph_index += 1
                label = f"T{t_idx} R{r_idx} C{c_idx}"
                blob = TextBlob(cell_text)
                for sentence in blob.sentences:
                    sentences.append({
                        "paragraph": paragraph_index,
                        "location_label": label,
                        "source": "table",
                        "sentence": str(sentence)
                    })

    return sentences


# Matches patterns that should be silently skipped:
#   - ALL-CAPS acronyms: NASA, FBI, HTML (2+ capital letters)
#   - Dot-separated abbreviations: U.S.A., e.g., i.e., Dr., Mr., etc.
#   - Mixed-case abbreviations ending in dot: Ltd., Corp., Inc.

_ABBREV_PATTERN = _re.compile(
    r'^('
    r'[A-Z]{2,}'                          # pure acronym: NASA, HTTP
    r'|[A-Za-z]([.][A-Za-z])+[.]?'         # dot-separated: U.S.A., e.g.
    r'|[A-Z][a-z]+[.]'                     # title abbrev: Dr., Mr., Inc.
    r')$'
)


def _is_abbrev_or_acronym(text):
    """Return True if the flagged token looks like an abbreviation or acronym."""
    token = text.strip().rstrip('.')
    return bool(_ABBREV_PATTERN.match(text.strip()) or
                _ABBREV_PATTERN.match(token) or
                (len(token) >= 2 and token.isupper()))


def analyze_sentence(sentence):
    matches = tool.check(sentence)
    errors = []
    for match in matches:
        error_text = sentence[match.offset:match.offset + match.error_length]
        # Skip if the flagged text is an abbreviation or acronym
        if _is_abbrev_or_acronym(error_text):
            continue
        errors.append({
            "error_text": error_text,
            "message": match.message,
            "suggestions": match.replacements[:6]
        })
    return errors


def analyze_docx(path):
    sentences = extract_sentences_from_docx(path)

    # --- Deduplication ---
    # seen_sentences: sentence_text -> index in report (first occurrence)
    seen_sentences = {}
    report = []
    total_errors = 0

    for s in sentences:
        sentence = s["sentence"]
        para_idx = s["paragraph"]
        label = s.get("location_label", f"\u00b6{para_idx}")
        source = s.get("source", "paragraph")

        if sentence in seen_sentences:
            first_idx = seen_sentences[sentence]
            report[first_idx]["occurrences"].append(para_idx)
            report[first_idx]["occurrence_labels"].append(label)
        else:
            errors = analyze_sentence(sentence)
            total_errors += len(errors)
            entry = {
                "paragraph": para_idx,
                "location_label": label,
                "source": source,
                "occurrences": [para_idx],
                "occurrence_labels": [label],
                "sentence": sentence,
                "errors": errors,
                "error_count": len(errors)
            }
            seen_sentences[sentence] = len(report)
            report.append(entry)

    return report, len(sentences), total_errors


def set_paragraph_text(para, new_text):
    """Replace paragraph text while preserving run formatting."""
    if not para.runs:
        para.add_run(new_text)
        return
    para.runs[0].text = new_text
    for run in para.runs[1:]:
        run.text = ""


def apply_corrections_to_docx(doc_path, report, decisions):
    """
    Apply accepted corrections to ALL occurrences of each sentence.
    - report items carry an `occurrences` list of every paragraph index
      that contains the sentence, so a single decision fixes every duplicate.
    - Corrections are keyed by paragraph_index and applied all at once per paragraph.
    """
    doc = Document(doc_path)

    # Build: para_idx -> list of (error_text, replacement)
    para_corrections = {}

    for si, item in enumerate(report):
        for ei, err in enumerate(item["errors"]):
            key = f"{si}_{ei}"
            decision = decisions.get(key, {})
            if decision.get("action") == "accept":
                suggestion = decision.get("suggestion")
                error_text = err.get("error_text", "")
                if suggestion and error_text:
                    # Apply to EVERY occurrence of this sentence
                    for para_idx in item.get("occurrences", [item["paragraph"]]):
                        para_corrections.setdefault(para_idx, []).append(
                            (error_text, suggestion))

    if not para_corrections:
        buffer = io.BytesIO()
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    # --- Walk the same way extract_sentences_from_docx does ---
    para_counter = 0

    # 1. Normal body paragraphs
    for para in doc.paragraphs:
        para_counter += 1
        if para_counter in para_corrections:
            new_text = para.text
            for old_text, replacement in para_corrections[para_counter]:
                new_text = new_text.replace(old_text, replacement, 1)
            if new_text != para.text:
                set_paragraph_text(para, new_text)

    # 2. Table cells  (must mirror the same loop order as extraction)
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                if not cell.text.strip():
                    continue
                para_counter += 1
                if para_counter in para_corrections:
                    # Cells contain their own paragraphs; rebuild full cell text
                    original_text = cell.text
                    new_text = original_text
                    for old_text, replacement in para_corrections[para_counter]:
                        new_text = new_text.replace(old_text, replacement, 1)
                    if new_text != original_text:
                        # Apply to the cell's paragraphs (preserve internal structure)
                        remaining = new_text
                        for p in cell.paragraphs:
                            if p.text and remaining:
                                set_paragraph_text(p, remaining[:len(p.text)])
                                remaining = remaining[len(p.text):]

    buffer = io.BytesIO()
    doc.save(buffer)
    buffer.seek(0)
    return buffer


@grammar_bp.route("/", methods=["GET", "POST"])
def index():
    if request.method == "GET":
        return render_template("grammarUI.html")

    file = request.files.get("file")
    if not file:
        return jsonify({"error": "No file provided"}), 400

    ext = os.path.splitext(secure_filename(file.filename))[1]
    unique_name = f"{uuid.uuid4()}{ext}"
    path = os.path.join(UPLOAD, unique_name)
    file.save(path)

    report, sentence_count, total_errors = analyze_docx(path)

    return jsonify({
        "report": report,
        "sentence_count": sentence_count,
        "total_errors": total_errors
    })


@grammar_bp.route("/download", methods=["POST"])
def download():
    file = request.files.get("file")
    decisions_raw = request.form.get("decisions", "{}")
    report_raw = request.form.get("report", "[]")

    if not file:
        return jsonify({"error": "No file provided"}), 400

    decisions = json.loads(decisions_raw)
    report = json.loads(report_raw)

    ext = os.path.splitext(secure_filename(file.filename))[1]
    unique_name = f"{uuid.uuid4()}{ext}"
    path = os.path.join(UPLOAD, unique_name)
    file.save(path)

    corrected_buffer = apply_corrections_to_docx(path, report, decisions)

    original_name = os.path.splitext(file.filename)[0]
    download_name = f"{original_name}_corrected.docx"

    return send_file(
        corrected_buffer,
        as_attachment=True,
        download_name=download_name,
        mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )


if __name__ == "__main__":
    from flask import Flask
    _app = Flask(__name__, template_folder="templates")
    _app.register_blueprint(grammar_bp)
    _app.run(debug=True)
