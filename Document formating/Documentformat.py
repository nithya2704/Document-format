from flask import Flask, render_template
from alignment import alignment_bp
from grammar import grammar_bp
from font1 import font_bp

app = Flask(__name__, template_folder="templates")

# ── Register blueprints ────────────────────────────────────────────────────
# Each blueprint owns its own URL prefix AND its own uploads/ folder,
# so the same filename uploaded to two different tabs never collides.
#
#   Tab                 URL prefix       Upload folder
#   ──────────────────  ───────────────  ──────────────────
#   Grammar             /grammar/*       uploads/grammar/
#   Alignment           /alignment/*     uploads/alignment/
#   Font (coming soon)  /font/*          uploads/font/
#
app.register_blueprint(grammar_bp)       # /grammar/*
app.register_blueprint(alignment_bp)     # /alignment/*
app.register_blueprint(font_bp)          # /font/*


@app.route("/")
def home():
    return render_template("DocumentformatUI.html")


if __name__ == "__main__":
    app.run(debug=True, threaded=False)
