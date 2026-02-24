from flask import Flask, request, jsonify, render_template_string
from werkzeug.utils import secure_filename
import os
import re
import platform
from docx import Document

# =============================
# Flask Configuration
# =============================
app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024
app.config['ALLOWED_EXTENSIONS'] = {'docx'}
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# =============================
# Platform / Dependency Check
# =============================
IS_WINDOWS = platform.system() == 'Windows'
WORD_AVAILABLE = False
LANGUAGE_TOOL_AVAILABLE = False

if IS_WINDOWS:
    try:
        import win32com.client
        import pythoncom
        WORD_AVAILABLE = True
    except ImportError:
        print("⚠ Install pywin32")

try:
    import language_tool_python
    tool = language_tool_python.LanguageTool('en-US')
    LANGUAGE_TOOL_AVAILABLE = True
except ImportError:
    print("⚠ Install language-tool-python")

# =============================
# Helpers
# =============================


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() == 'docx'


def split_into_sentences(text):
    return re.split(r'(?<=[.!?])\s+(?=[A-Z<])', text.strip())


def ai_correct_sentence(sentence):
    if not LANGUAGE_TOOL_AVAILABLE:
        return sentence

    matches = tool.check(sentence)
    corrected = sentence

    for m in sorted(matches, key=lambda x: x.offset, reverse=True):
        length = getattr(m, 'errorLength', m.error_length)
        replacement = m.replacements[0] if m.replacements else corrected[m.offset:m.offset+length]
        corrected = corrected[:m.offset] + \
            replacement + corrected[m.offset+length:]

    return corrected


def split_identifier(text):
    if not text or " " in text:
        return text

    text = re.sub(r'([A-Z]+)([A-Z][a-z])', r'\1 \2', text)
    text = re.sub(r'([a-z0-9])([A-Z])', r'\1 \2', text)
    return text.strip()


def smart_correct(original, corrected):
    original = original.strip()
    corrected = corrected.strip()

    if original == corrected:
        split_text = split_identifier(original)
        if split_text != original and len(split_text.split()) > 1:
            return split_text
        return None

    return corrected

# =============================
# Word COM Checker
# =============================


def check_with_word_com(file_path):
    pythoncom.CoInitialize()
    word = doc = None

    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        doc = word.Documents.Open(os.path.abspath(file_path))

        sentences = {}
        spelling = grammar = 0

        for para in doc.Paragraphs:
            text = para.Range.Text.strip()
            if not text:
                continue
            for s in split_into_sentences(text):
                sentences[s] = {
                    "original": s,
                    "corrected": s,
                    "errors": []
                }

        for err in doc.SpellingErrors:
            spelling += 1
            for s in sentences:
                if err.Text in s:
                    corrected_sentence = ai_correct_sentence(
                        sentences[s]["corrected"])
                    final_corrected = smart_correct(
                        sentences[s]["original"], corrected_sentence)
                    if final_corrected is None:
                        continue

                    sentences[s]["corrected"] = final_corrected
                    sentences[s]["errors"].append({
                        "type": "SPELLING",
                        "text": err.Text
                    })
                    break

        for err in doc.GrammaticalErrors:
            grammar += 1
            for s in sentences:
                if err.Text in s:
                    corrected_sentence = ai_correct_sentence(
                        sentences[s]["corrected"])
                    final_corrected = smart_correct(
                        sentences[s]["original"], corrected_sentence)
                    if final_corrected is None:
                        continue

                    sentences[s]["corrected"] = final_corrected
                    sentences[s]["errors"].append({
                        "type": "GRAMMAR",
                        "text": err.Text
                    })
                    break

        results = []
        idx = 0
        for v in sentences.values():
            if not v["errors"]:
                continue
            if v["original"] == v["corrected"]:
                continue

            results.append({
                "id": idx,
                "original": v["original"],
                "corrected": v["corrected"],
                "current": v["original"],
                "error_details": v["errors"]
            })
            idx += 1

        return results, spelling, grammar

    finally:
        if doc:
            doc.Close(False)
        if word:
            word.Quit()
        pythoncom.CoUninitialize()


# =============================
# UI Template
# =============================
HTML_TEMPLATE = """
<!DOCTYPE html>
<html>
<head>
<title>Word Grammar Checker</title>
<style>
body { font-family: Arial; background:#eef2ff; padding:40px; }
.card {
    max-width:900px; margin:auto; background:white;
    padding:30px; border-radius:10px;
    box-shadow:0 8px 25px rgba(0,0,0,0.15);
}
button {
    padding:8px 18px; border:none; border-radius:20px;
    font-size:14px; cursor:pointer; margin-right:10px;
}
.accept { background:#16a34a; color:white; }
.ignore { background:#e5e7eb; }
.section {
    margin-bottom:25px; padding:15px;
    background:#f9fafb; border-left:5px solid #dc2626;
}
</style>
</head>
<body>
<div class="card">
<h2>📝 Word Grammar Checker</h2>

<input type="file" id="file" accept=".docx"><br><br>
<button onclick="upload()">Analyze Document</button>

<div style="margin:20px 0;">
    <button onclick="showSpelling()">Spelling Errors</button>
    <button onclick="showGrammar()">Grammar Errors</button>
</div>

<div id="result"></div>
</div>

<script>
let API_DATA = null;

async function upload() {
    const file = document.getElementById("file").files[0];
    if (!file) return alert("Select a .docx file");

    const fd = new FormData();
    fd.append("file", file);

    document.getElementById("result").innerHTML = "<b>Analyzing...</b>";

    const res = await fetch("/upload", { method:"POST", body: fd });
    API_DATA = await res.json();

    showSpelling();
}

function showSpelling() {
    render("SPELLING");
}

function showGrammar() {
    render("GRAMMAR");
}

function render(type) {
    let html = "";

    API_DATA.results.forEach(item => {
        const errs = item.error_details.filter(e => e.type === type);
        if (!errs.length) return;

        html += `
        <div class="section" id="row-${item.id}">
            <p><b>Wrong sentence:</b><br>${escape(item.current)}</p>
            <p><b>Corrected:</b><br><span style="color:green">${escape(item.corrected)}</span></p>
        `;

        if (type === "SPELLING") {
            html += `
            <button class="accept" onclick="accept(${item.id})">✔ Accept</button>
            <button class="ignore" onclick="ignore(${item.id})">Ignore</button>
            `;
        }

        html += "</div>";
    });

    document.getElementById("result").innerHTML =
        html || `<p style="color:green"><b>No ${type.toLowerCase()} errors 🎉</b></p>`;
}

function accept(id) {
    const item = API_DATA.results.find(r => r.id === id);
    item.current = item.corrected;

    document.getElementById("row-" + id).innerHTML =
        "<p style='color:green'><b>✔ Accepted</b></p><p>" + escape(item.current) + "</p>";
}

function ignore(id) {
    document.getElementById("row-" + id).remove();
}

function escape(text) {
    const d = document.createElement("div");
    d.textContent = text;
    return d.innerHTML;
}
</script>
</body>
</html>
"""

# =============================
# Routes
# =============================


@app.route('/')
def index():
    return render_template_string(HTML_TEMPLATE)


@app.route('/upload', methods=['POST'])
def upload():
    file = request.files.get('file')
    if not file or not allowed_file(file.filename):
        return jsonify({"success": False})

    path = os.path.join(app.config['UPLOAD_FOLDER'],
                        secure_filename(file.filename))
    file.save(path)

    try:
        results, s, g = check_with_word_com(path)
        return jsonify({
            "success": True,
            "results": results,
            "spelling_errors": s,
            "grammar_errors": g
        })
    finally:
        os.remove(path)


# =============================
# Run
# =============================
if __name__ == "__main__":
    print("🚀 Word Grammar Checker")
    print("http://localhost:5000")
    app.run(debug=True)
