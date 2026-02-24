from flask import Flask, request, render_template_string, jsonify, send_file
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
import os
import json
from werkzeug.utils import secure_filename

app = Flask(__name__)

UPLOAD = "uploads"
OUTPUT = "outputs"
os.makedirs(UPLOAD, exist_ok=True)
os.makedirs(OUTPUT, exist_ok=True)

# ---------------- ANALYSIS ----------------


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
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY
    }

    idx_map = {d["index"]: d["type"] for d in detected}

    for i, p in enumerate(doc.paragraphs):
        etype = idx_map.get(i)
        if not etype:
            continue

        pf = p.paragraph_format

        # ALIGNMENT
        align_value = cfg.get(f"{etype}_alignment", "left")
        p.alignment = alignment_map.get(align_value, WD_ALIGN_PARAGRAPH.LEFT)

        # LINE SPACING
        spacing_value = float(cfg.get(f"{etype}_spacing", 1.0))
        pf.line_spacing = spacing_value

        # MULTI-SELECT PARAGRAPH SPACING
        if etype == "paragraph":
            options = cfg.get("paragraph_spacing_options", [])

            pf.space_before = None
            pf.space_after = None

            if "add_before" in options:
                pf.space_before = Pt(12)

            if "add_after" in options:
                pf.space_after = Pt(12)

            if "remove_before" in options:
                pf.space_before = Pt(0)

            if "remove_after" in options:
                pf.space_after = Pt(0)

            pf.line_spacing_rule = None

    doc.save(output_path)


# ---------------- UI ----------------
HTML = """
<!DOCTYPE html>
<html>
<head>
<title>DOCX Formating And Alignment</title>
<style>
body { font-family: Arial; background:#f3f4f6; padding:40px; }
.box { background:white; width:750px; margin:auto; padding:30px; border-radius:12px; }

label { font-weight:bold; display:block; margin-top:15px; color:black; }
select { width:100%; padding:8px; margin-top:5px; color:black; }

button {
  padding:12px;
  width:100%;
  background:#4f46e5;
  color:white;
  border:none;
  border-radius:20px;
  margin-top:20px;
  cursor:pointer;
}

h4 { margin-top:25px; }

/* Multiselect dropdown */
.dropdown { position: relative; margin-top:10px; }

.dropbtn {
  width:100%;
  padding:8px;
  background:white;
  border:1px solid #ccc;
  text-align:left;
  cursor:pointer;
  border-radius:6px;
  color:black;              /* FORCE BLACK */
  font-weight:500;
}

.dropdown-content {
  display:none;
  position:absolute;
  background:white;
  width:100%;
  border:1px solid #ccc;
  padding:10px;
  z-index:1;
  border-radius:6px;
  color:black;              /* FORCE BLACK */
}

.dropdown-content label {
  display:block;
  margin-bottom:6px;
  font-weight:normal;
  color:black;              /* FORCE BLACK */
}

.dropdown-content input[type="checkbox"] {
  accent-color: #4f46e5;
}

.show { display:block; }

</style>
</head>
<body>
<div class="box">
<h2>📄 DOCX Advanced Formatter</h2>

<input type="file" id="file">
<br><br>
<button onclick="analyse()">Analyse Document</button>

<div id="ui"></div>
</div>

<script>

let types = [];

// Toggle dropdown
function toggleDropdown(){
    document.getElementById("spacingDropdown").classList.toggle("show");
}

// Auto-close when clicking outside
window.onclick = function(event) {
    if (!event.target.closest('.dropdown')) {
        let dropdown = document.getElementById("spacingDropdown");
        if (dropdown && dropdown.classList.contains('show')) {
            dropdown.classList.remove('show');
        }
    }
}

// Update button text
function updateDropdownText(){

    let dropdown = document.getElementById("spacingDropdown");
    let checkboxes = document.querySelectorAll("#spacingDropdown input[type='checkbox']");
    let selected = [];

    checkboxes.forEach(cb=>{
        if(cb.checked){
            selected.push(cb.parentElement.innerText.trim());
        }
    });

    let button = document.getElementById("spacingButton");

    if(selected.length === 0){
        button.innerText = "Select Spacing Options ▼";
    }
    else{
        button.innerText = selected.join(", ") + " ▼";
    }

    button.style.color = "black";

    // ✅ AUTO CLOSE AFTER SELECTION
    dropdown.classList.remove("show");
}

async function analyse(){
    let fd = new FormData();
    fd.append("file", file.files[0]);

    let res = await fetch("/analyse",{method:"POST",body:fd});
    let data = await res.json();
    types = data.types;

    let html = "<h3>Formatting Options</h3>";

    types.forEach(t=>{
        html += `
        <h4>${t.toUpperCase()}</h4>

        <label>Alignment</label>
        <select id='${t}_alignment'>
            <option value="left">Align Left</option>
            <option value="right">Align Right</option>
            <option value="justify">Justify</option>
        </select>

        <label>Line Spacing</label>
        <select id='${t}_spacing'>
            <option value="1">Single</option>
            <option value="1.5">1.5</option>
            <option value="2">Double</option>
            <option value="2.5">2.5</option>
            <option value="3">Triple</option>
        </select>
        `;
    });

    if(types.includes("paragraph")){
        html += `
        <h4>Paragraph Spacing (Word Style)</h4>

        <div class="dropdown">
            <button type="button"
                    id="spacingButton"
                    onclick="toggleDropdown()"
                    class="dropbtn">
                Select Spacing Options ▼
            </button>

            <div id="spacingDropdown" class="dropdown-content">
                <label><input type="checkbox" value="add_before" onchange="updateDropdownText()"> Add Space Before Paragraph</label>
                <label><input type="checkbox" value="add_after" onchange="updateDropdownText()"> Add Space After Paragraph</label>
                <label><input type="checkbox" value="remove_before" onchange="updateDropdownText()"> Remove Space Before Paragraph</label>
                <label><input type="checkbox" value="remove_after" onchange="updateDropdownText()"> Remove Space After Paragraph</label>
            </div>
        </div>
        `;
    }

    html += "<button onclick='generate()'>Generate Formatted DOCX</button>";
    ui.innerHTML = html;
}

async function generate(){
    let cfg = {};

    types.forEach(t=>{
        cfg[`${t}_alignment`] = document.getElementById(`${t}_alignment`).value;
        cfg[`${t}_spacing`] = document.getElementById(`${t}_spacing`).value;
    });

    if(types.includes("paragraph")){
        let checkboxes = document.querySelectorAll("#spacingDropdown input[type='checkbox']");
        let selected = [];

        checkboxes.forEach(cb=>{
            if(cb.checked) selected.push(cb.value);
        });

        cfg["paragraph_spacing_options"] = selected;
    }

    let fd = new FormData();
    fd.append("file", file.files[0]);
    fd.append("config", JSON.stringify(cfg));

    let res = await fetch("/format",{method:"POST",body:fd});
    let blob = await res.blob();

    let a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "formatted.docx";
    a.click();
}

</script>
</body>
</html>
"""


# ---------------- ROUTES ----------------
@app.route("/")
def home():
    return render_template_string(HTML)


@app.route("/analyse", methods=["POST"])
def analyse():
    f = request.files["file"]
    path = os.path.join(UPLOAD, secure_filename(f.filename))
    f.save(path)

    detected, types = analyse_docx(path)
    return jsonify({"types": types})


@app.route("/format", methods=["POST"])
def format_doc():
    f = request.files["file"]
    cfg = json.loads(request.form["config"])

    in_path = os.path.join(UPLOAD, secure_filename(f.filename))
    f.save(in_path)

    out_path = os.path.join(OUTPUT, "formatted.docx")

    detected, _ = analyse_docx(in_path)
    format_docx(in_path, out_path, detected, cfg)

    return send_file(out_path, as_attachment=True)


if __name__ == "__main__":
    app.run(debug=True)
