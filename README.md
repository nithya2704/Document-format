# DocFormat — Word Document Formatting Tool

DocFormat is a Flask web application for reviewing grammar and standardizing the alignment, spacing, fonts, and colors of Microsoft Word documents. Upload a `.docx` file once, apply one or more tools, and download the latest formatted document.

The application combines three tools in a shared browser session. Each applied operation becomes the input for the next tool, so changes can be built up without downloading and uploading between steps.

## Contents

- [Features](#features)
- [Requirements](#requirements)
- [Installation and startup](#installation-and-startup)
- [Using the application](#using-the-application)

## Features

| Tool                  | Capabilities                                                                                                                                                                                                             |
| --------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| Grammar review        | Checks US English using LanguageTool; analyzes body paragraphs and table cells; groups repeated sentences; supports accepting, declining, or entering custom corrections.                                                |
| Alignment and spacing | Applies left, right, or justified alignment and line spacing to paragraphs, lists, and table-cell text; adds or removes spacing before and after content.                                                                |
| Font formatting       | Detects heading levels 1–6, paragraphs, lists, and table text; applies fonts and sizes by element type; provides heading/list bold options and separate color controls for paragraph text, list text, and table headers. |
| Document previews     | Displays the current document and proposed alignment/font changes using Word-generated PDFs where available, with HTML preview paths in the implementation.                                                              |
| Shared workflow       | Maintains one current working document per session and lets users continue through additional tools before downloading.                                                                                                  |

Alignment and font processing use heuristics to identify cover-page content. Font processing also detects table-of-contents regions. Grammar processing skips body paragraphs containing Word field codes and flags code-like text for the review interface. These protections depend on document structure; review the output before distributing it.

## Requirements

### Platform and supporting software

- **Windows** for the integrated application as supplied. `alignment.py` imports `pythoncom` unconditionally, so Linux and macOS cannot start the application without a code change.
- **Python:** the bundled virtual environment records Python **3.14.2**. This is the source environment, not a guarantee that all other Python versions have been tested.
- **Java 17 or newer**, accessible through `java -version`, for the default LanguageTool version used by the bundled Python wrapper.
- **Microsoft Word desktop** for DOCX-to-PDF previews through `docx2pdf`. Importing the conversion libraries does not verify that Word is installed.
- A modern web browser and write permission in the project directory.
- Internet access for package installation, the initial LanguageTool download, and TextBlob/NLTK data downloads. The templates also request Google Fonts.

### Python dependencies

The archive does not include a `requirements.txt` or a dependency lockfile. The following versions were read from the package metadata in its bundled `.venv`:

| Package              | Bundled version | Purpose                                           |
| -------------------- | --------------- | ------------------------------------------------- |
| Flask                | 3.1.2           | Web application, templates, routing, and sessions |
| Werkzeug             | 3.1.5           | Upload filename handling and Flask utilities      |
| python-docx          | 1.2.0           | Read and modify Word documents                    |
| language-tool-python | 3.2.2           | Local LanguageTool integration                    |
| textblob             | 0.19.0          | Sentence segmentation                             |
| nltk                 | 3.9.2           | Language resources used by TextBlob               |
| docx2pdf             | 0.1.8           | Word-based PDF conversion                         |
| pywin32              | 311             | Windows COM support, including `pythoncom`        |

These are the relevant installed versions, not a complete lock of all transitive dependencies.

## Installation and startup

Run the following commands in **Windows PowerShell**, from the extracted project folder containing `Documentformat.py`.

### 1. Create a fresh virtual environment

The supplied `.venv` contains paths from the original machine. Create a new environment rather than relying on it. The commands below use `.venv-local` so they do not overwrite the supplied environment.

```powershell
cd "C:\path\to\Document formating"
py -m venv .venv-local
.\.venv-local\Scripts\python.exe -m pip install --upgrade pip
```

Activation is optional: all commands use the environment's Python directly.

### 2. Install the application dependencies

```powershell
.\.venv-local\Scripts\python.exe -m pip install Flask==3.1.2 Werkzeug==3.1.5 python-docx==1.2.0 language-tool-python==3.2.2 textblob==0.19.0 nltk==3.9.2 docx2pdf==0.1.8 pywin32==311
```

### 3. Prepare the language resources

```powershell
java -version
.\.venv-local\Scripts\python.exe -m textblob.download_corpora
```

If sentence tokenization reports a missing `punkt_tab` resource, install it explicitly:

```powershell
.\.venv-local\Scripts\python.exe -m nltk.downloader punkt_tab
```

`grammar.py` initializes LanguageTool when imported. The first application startup may download LanguageTool and take longer than later runs. Java must be available even if you intend to use only alignment or font formatting, because the main application imports all three tools at startup.

### 4. Start DocFormat

```powershell
.\.venv-local\Scripts\python.exe Documentformat.py
```

Open **http://127.0.0.1:5000/** in your browser. Keep the terminal open while using the application; press `Ctrl+C` to stop it.

Use `Documentformat.py` as the entry point. The standalone launch blocks in the individual modules do not configure the shared working-file helpers required by the integrated workflow.

## Using the application

1. Open the home page and upload a Word `.docx` document.
2. Select a tool and analyze the current document.
3. Review the detected content or suggested corrections.
4. Apply the changes using that tool's apply/download action.
5. Continue to another tool or download the current document from the home interface.

### Grammar review

- Analyze the document to see sentence-level issues and suggested replacements.
- Accept or decline suggestions, or enter a custom replacement.
- Apply the selected corrections using the interface's completion action. Accepted corrections are saved as the new working document.
- Repeated sentences are grouped; an accepted correction can affect each recorded occurrence.

### Alignment and spacing

- Analyze the document to detect paragraphs, lists, and tables.
- Configure alignment and line spacing separately for each detected content category.
- Add or remove before/after spacing. The implementation uses 12 points when adding spacing and 0 points when removing it.
- Review the preview, then apply the formatting.

Table alignment controls change paragraph alignment **inside table cells**, rather than positioning the entire table on the page. Code-like paragraphs are protected from justification by font, style, and text heuristics.

### Font formatting

- Analyze the document to identify element types, counts, and detected font information.
- Set fonts and sizes for the detected content categories.
- Configure bold headings/lists and the available text and table-header colors.
- Optionally highlight element types in the preview, then apply formatting. The final font-formatting route disables preview highlighting.

**Previewing does not replace the current working document.** Apply the changes before moving to another tool. Uploading another document replaces the current working-file pointer for that session.
