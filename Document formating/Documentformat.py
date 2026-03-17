import os
import uuid
from flask import Flask, render_template, session, request, jsonify, send_file
from werkzeug.utils import secure_filename

from alignment import alignment_bp
from grammar import grammar_bp
from font1 import font_bp

app = Flask(__name__, template_folder="templates")

# ── Secret key — required for Flask sessions ─────────────────────────────
# Change this to a long random value before deploying.
app.secret_key = "docformat-local-dev-secret-key-change-me"

# ── Shared upload folder ──────────────────────────────────────────────────
SHARED_UPLOAD = "uploads/shared"
os.makedirs(SHARED_UPLOAD, exist_ok=True)

# ── In-memory working-file store ──────────────────────────────────────────
# Maps  session_id (sid)  →  absolute path of the current working file.
# On Windows + single-process dev server this is perfectly safe.
# For multi-worker / OCI deployments replace with Redis or a DB.
_working_files: dict = {}


def get_working_path(sid: str):
    """Return the current working-file path for this session, or None."""
    return _working_files.get(sid)


def set_working_path(sid: str, path: str) -> None:
    """Update the working-file path for this session."""
    _working_files[sid] = os.path.abspath(path)


# Expose helpers via app.config so blueprints can import the Flask app and
# call them without a circular import.
app.config["GET_WORKING_PATH"] = get_working_path
app.config["SET_WORKING_PATH"] = set_working_path

# ── Register blueprints ───────────────────────────────────────────────────
app.register_blueprint(grammar_bp)    # /grammar/*
app.register_blueprint(alignment_bp)  # /alignment/*
app.register_blueprint(font_bp)       # /font/*


# ── Home ──────────────────────────────────────────────────────────────────
@app.route("/")
def home():
    # Assign a stable session ID so every browser tab from this user
    # maps to the same slot in _working_files.
    if "sid" not in session:
        session["sid"] = uuid.uuid4().hex
    return render_template("DocumentformatUI.html")


# ── Upload (called once from the home page) ───────────────────────────────
@app.route("/upload", methods=["POST"])
def upload():
    """Accept the document once, store it, record path in session store."""
    if "sid" not in session:
        session["sid"] = uuid.uuid4().hex
    sid = session["sid"]

    f = request.files.get("file")
    if not f or not f.filename:
        return jsonify({"error": "No file provided"}), 400

    ext = os.path.splitext(secure_filename(f.filename))[1].lower()
    if ext != ".docx":
        return jsonify({"error": "Only .docx files are supported"}), 400

    # Each session gets its own sub-folder → no path collisions between users.
    session_dir = os.path.join(SHARED_UPLOAD, sid)
    os.makedirs(session_dir, exist_ok=True)

    filename = f"{uuid.uuid4().hex}{ext}"
    save_path = os.path.join(session_dir, filename)
    f.save(save_path)

    set_working_path(sid, save_path)

    return jsonify({
        "status": "ok",
        "original_name": f.filename,
    })


# ── Download current working file ─────────────────────────────────────────
@app.route("/download_current", methods=["GET"])
def download_current():
    """Stream the latest version of this session's working file."""
    sid = session.get("sid")
    if not sid:
        return jsonify({"error": "No session"}), 400

    path = get_working_path(sid)
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file found for this session"}), 404

    return send_file(
        path,
        as_attachment=True,
        download_name="formatted_document.docx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument"
            ".wordprocessingml.document"
        ),
    )


if __name__ == "__main__":
    app.run(debug=True, threaded=False)
