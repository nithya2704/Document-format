from flask import Flask, render_template_string

app = Flask(__name__)

# ---------------- MAIN PAGE UI (UNCHANGED) ----------------
MAIN_HTML = """
<!DOCTYPE html>
<html>
<head>
    <title>Document Formatter</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            background-color: #f2f4f8;
            margin: 0;
            padding: 0;
        }

        .container {
            width: 70%;
            margin: 60px auto;
            text-align: center;
        }

        h1 {
            margin-bottom: 30px;
        }

        .tabs {
            display: flex;
            justify-content: center;
            gap: 20px;
        }

        .tab-button {
            padding: 12px 25px;
            border: none;
            border-radius: 8px;
            background-color: #e0e0e0;
            cursor: pointer;
            font-size: 16px;
            font-weight: bold;
        }

        .tab-button:hover {
            background-color: #d0d0d0;
        }

        .active {
            background-color: #4CAF50;
            color: white;
        }

        .tab-content {
            margin-top: 40px;
            padding: 30px;
            border-radius: 10px;
            background: white;
            box-shadow: 0px 4px 10px rgba(0,0,0,0.1);
        }
    </style>

    <script>
        function showTab(tabName) {
            document.getElementById("font").style.display = "none";
            document.getElementById("grammar").style.display = "none";
            document.getElementById("formatting").style.display = "none";

            document.getElementById(tabName).style.display = "block";

            var buttons = document.getElementsByClassName("tab-button");
            for (var i = 0; i < buttons.length; i++) {
                buttons[i].classList.remove("active");
            }

            document.getElementById(tabName + "-btn").classList.add("active");
        }

        window.onload = function() {
            showTab("font");
        }
    </script>
</head>
<body>

<div class="container">
    <h1>📄 Document Formatter</h1>

    <div class="tabs">
        <button id="font-btn" class="tab-button" onclick="showTab('font')">
            Font Style & Size
        </button>

        <button id="grammar-btn" class="tab-button" onclick="showTab('grammar')">
            Grammar & Spell Check
        </button>

        <button id="formatting-btn" class="tab-button" onclick="showTab('formatting')">
            Formatting & Alignment
        </button>
    </div>

    <div id="font" class="tab-content">
        <h3>Font Style & Size Section</h3>
        <p>Backend functionality will be added later.</p>
    </div>

    <div id="grammar" class="tab-content" style="display:none;">
        <iframe src="/grammarcheck" 
                width="100%" 
                height="700px" 
                style="border:none;">
        </iframe>
    </div>

    <div id="formatting" class="tab-content" style="display:none;">
        <iframe src="/alignment" 
                width="100%" 
                height="700px" 
                style="border:none;">
        </iframe>
    </div>

</div>

</body>
</html>
"""

# ---------------- ROUTES ----------------


@app.route("/")
def home():
    return render_template_string(MAIN_HTML)


# This route will render your existing alignment.py UI
@app.route("/alignment")
def alignment():
    from alignment import HTML  # import your existing HTML variable
    return render_template_string(HTML)


@app.route("/grammarcheck")
def grammarcheck():
    from grammarcheck import HTML_TEMPLATE  # import your existing HTML variable
    return render_template_string(HTML_TEMPLATE)


if __name__ == "__main__":
    app.run(debug=True)
