#font1.py
import os
import re
import json
import uuid
import platform
from collections import Counter

from flask import (Blueprint, Flask, request, jsonify,
                   send_file, render_template, session, current_app)

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_COLOR_INDEX, WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ── Platform / optional deps ──────────────────────────────────────────────────
IS_WINDOWS = platform.system() == "Windows"
WORD_AVAILABLE = False
if IS_WINDOWS:
    try:
        import win32com.client
        import pythoncom
        from docx2pdf import convert
        WORD_AVAILABLE = True
    except ImportError:
        pass

# ── Blueprint ─────────────────────────────────────────────────────────────────
font_bp = Blueprint("font", __name__, url_prefix="/font",
                    template_folder="templates")

OUTPUT = "outputs/font"
os.makedirs(OUTPUT, exist_ok=True)

# =============================================================================
#  USER-FACING ELEMENT TYPES  (cover / TOC / TOF are detected internally
#  and silently skipped — they never reach the UI or the formatter)
# =============================================================================
# The types the user sees and can configure:
USER_TYPES = [
    "HEADING_1", "HEADING_2", "HEADING_3",
    "HEADING_4", "HEADING_5", "HEADING_6",
    "PARAGRAPH", "LIST_ITEM", "TABLE_TEXT",
]

# Highlight colours shown in preview for each user-facing type
HIGHLIGHT_COLORS = {
    "HEADING_1": WD_COLOR_INDEX.BRIGHT_GREEN,
    "HEADING_2": WD_COLOR_INDEX.TURQUOISE,
    "HEADING_3": WD_COLOR_INDEX.PINK,
    "HEADING_4": WD_COLOR_INDEX.VIOLET,
    "HEADING_5": WD_COLOR_INDEX.TEAL,
    "HEADING_6": WD_COLOR_INDEX.DARK_BLUE,
    "PARAGRAPH": WD_COLOR_INDEX.WHITE,
    "LIST_ITEM": WD_COLOR_INDEX.YELLOW,
    "TABLE_TEXT": WD_COLOR_INDEX.RED,
}

_COLOR_MAP = {
    WD_COLOR_INDEX.BRIGHT_GREEN: "green",
    WD_COLOR_INDEX.TURQUOISE:    "cyan",
    WD_COLOR_INDEX.PINK:         "magenta",
    WD_COLOR_INDEX.VIOLET:       "darkMagenta",
    WD_COLOR_INDEX.TEAL:         "darkCyan",
    WD_COLOR_INDEX.DARK_BLUE:    "darkBlue",
    WD_COLOR_INDEX.WHITE:        "white",
    WD_COLOR_INDEX.YELLOW:       "yellow",
    WD_COLOR_INDEX.GREEN:        "green",
    WD_COLOR_INDEX.RED:          "red"
}

# =============================================================================
#  SESSION HELPERS
# =============================================================================


def _get_path():
    sid = session.get("sid")
    if not sid:
        return None
    return current_app.config["GET_WORKING_PATH"](sid)


def _set_path(path: str):
    sid = session.get("sid")
    if sid:
        current_app.config["SET_WORKING_PATH"](sid, path)


def _session_out_dir():
    sid = session.get("sid", "default")
    d = os.path.join(OUTPUT, sid)
    os.makedirs(d, exist_ok=True)
    return d

# =============================================================================
#  LOW-LEVEL XML HELPERS
#  UTILITY FUNCTIONS
# =============================================================================


def is_in_footer_or_header(para) -> bool:
    """Return True if para lives inside a header or footer frame."""
    _FH = {qn("w:hdr"), qn("w:ftr")}
    node = para._element.getparent()
    while node is not None:
        if node.tag in _FH:
            return True
        node = node.getparent()
    return False


def get_paragraph_indentation(para) -> float:
    try:
        if para.paragraph_format.left_indent:
            return para.paragraph_format.left_indent.inches
    except Exception:
        pass
    return 0.0


def has_page_break_before(para) -> bool:
    try:
        pPr = para._element.find(qn("w:pPr"))
        if pPr is not None:
            pb = pPr.find(qn("w:pageBreakBefore"))
            if pb is not None:
                val = pb.get(qn("w:val"), "true")
                if val.lower() not in ("false", "0", "off"):
                    return True
        for br in para._element.xpath(".//w:br"):
            if br.get(qn("w:type"), "") == "page":
                return True
        if "lastRenderedPageBreak" in para._element.xml:
            return True
    except Exception:
        pass
    return False


# ── Numbered list resolution ──────────────────────────────────────────────────

_NUMBERED_FMTS = frozenset({
    "decimal", "lowerLetter", "upperLetter", "lowerRoman",
    "upperRoman", "ordinal", "cardinalText", "decimalZero",
    "decimalEnclosedCircle", "decimalFullWidth", "decimalHalfWidth",
})

NUMBERED_HEADING_RE = re.compile(
    r"^\s*(\d+(?:\.\d+)+)\.?(.*)$"
)


def detect_numbered_heading(text: str):
    """Detect '1.2 Title' style headings. Returns (level, num, title) or (None,None,None)."""
    if not text:
        return None, None, None
    m = NUMBERED_HEADING_RE.match(text)
    if not m:
        return None, None, None
    num = m.group(1)
    title = m.group(2).strip()
    parts = num.split(".")
    if len(parts) < 2:
        return None, None, None
    return min(6, len(parts)), num, title


def _resolve_numFmt(para) -> str:
    try:
        numbering_part = para.part.numbering_part
    except Exception:
        numbering_part = None

    def _fmt(numId_val, ilvl_val):
        if numId_val == "0":
            return ""
        if numbering_part is None:
            return "bullet"
        nbr = numbering_part._element
        abs_nodes = nbr.xpath(
            f".//w:num[@w:numId='{numId_val}']/w:abstractNumId")
        if not abs_nodes:
            return "bullet"
        abstract_id = abs_nodes[0].get(qn("w:val"), "0")
        fmt_nodes = nbr.xpath(
            f".//w:abstractNum[@w:abstractNumId='{abstract_id}']"
            f"/w:lvl[@w:ilvl='{ilvl_val}']/w:numFmt")
        return fmt_nodes[0].get(qn("w:val"), "") if fmt_nodes else "bullet"

    pPr = para._element.find(qn("w:pPr"))
    if pPr is not None:
        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            nid = numPr.find(qn("w:numId"))
            ilvl = numPr.find(qn("w:ilvl"))
            nid_v = nid.get(qn("w:val"), "0") if nid is not None else "0"
            ilvl_v = ilvl.get(qn("w:val"), "0") if ilvl is not None else "0"
            if nid_v == "0":
            numId_node = numPr.find(qn("w:numId"))
            ilvl_node  = numPr.find(qn("w:ilvl"))
            numId_val  = numId_node.get(qn("w:val"), "0") if numId_node is not None else "0"
            ilvl_val   = ilvl_node.get(qn("w:val"),  "0") if ilvl_node  is not None else "0"
            if numId_val == "0":
                return ""
            return _fmt(nid_v, ilvl_v)

    style = para.style if para.style else None
    while style is not None:
        try:
            s_pPr = style.element.find(qn("w:pPr"))
            if s_pPr is not None:
                s_numPr = s_pPr.find(qn("w:numPr"))
                if s_numPr is not None:
                    nid = s_numPr.find(qn("w:numId"))
                    ilvl = s_numPr.find(qn("w:ilvl"))
                    nid_v = nid.get(
                        qn("w:val"), "0") if nid is not None else "0"
                    ilvl_v = ilvl.get(
                        qn("w:val"), "0") if ilvl is not None else "0"
                    if nid_v == "0":
            style_elem = style.element
            style_pPr  = style_elem.find(qn("w:pPr"))
            if style_pPr is not None:
                style_numPr = style_pPr.find(qn("w:numPr"))
                if style_numPr is not None:
                    numId_node = style_numPr.find(qn("w:numId"))
                    ilvl_node  = style_numPr.find(qn("w:ilvl"))
                    numId_val  = numId_node.get(qn("w:val"), "0") if numId_node is not None else "0"
                    ilvl_val   = ilvl_node.get(qn("w:val"),  "0") if ilvl_node  is not None else "0"
                    if numId_val == "0":
                        return ""
                    return _fmt(nid_v, ilvl_v)
        except Exception:
            pass
        try:
            style = style.base_style
        except Exception:
            break
    return ""


def get_list_type(para) -> str:
    sn = ""
    try:
        sn = para.style.name if para.style else ""
    except Exception:
        pass
    if "List Bullet" in sn or "List Number" in sn or sn.startswith("List ") or sn == "List":
        return "LIST_ITEM"
    fmt = _resolve_numFmt(para)
    if not fmt:
        return None
    return "LIST_ITEM"

# =============================================================================
#  INTERNAL-ONLY: Cover page + TOC + Table of Figures detection
#  These ranges are used to SKIP those paragraphs entirely.
# =============================================================================


def _style_base(para) -> str:

# ── Highlight colour definitions ──────────────────────────────────────────────

HIGHLIGHT_COLORS = {
    "TITLE":         WD_COLOR_INDEX.YELLOW,
    "COVER_PAGE":    WD_COLOR_INDEX.DARK_YELLOW,
    "TOC_TITLE":     WD_COLOR_INDEX.RED,
    "TOC_HEADING_1": WD_COLOR_INDEX.GRAY_25,
    "TOC_HEADING_2": WD_COLOR_INDEX.GRAY_50,
    "TOC_HEADING_3": WD_COLOR_INDEX.BLUE,
    "TOC_HEADING_4": WD_COLOR_INDEX.DARK_BLUE,
    "TOC_HEADING_5": WD_COLOR_INDEX.TEAL,
    "TOC_HEADING_6": WD_COLOR_INDEX.TURQUOISE,
    "HEADING_1":     WD_COLOR_INDEX.BRIGHT_GREEN,
    "HEADING_2":     WD_COLOR_INDEX.TURQUOISE,
    "HEADING_3":     WD_COLOR_INDEX.PINK,
    "HEADING_4":     WD_COLOR_INDEX.VIOLET,
    "HEADING_5":     WD_COLOR_INDEX.TEAL,
    "HEADING_6":     WD_COLOR_INDEX.DARK_BLUE,
    "PARAGRAPH":     WD_COLOR_INDEX.WHITE,
    "LIST_ITEM":     WD_COLOR_INDEX.RED,
    # Table zones — distinct colours for header / body / footer rows
    "TABLE_HEADER":  WD_COLOR_INDEX.GREEN,
    "TABLE_BODY":    WD_COLOR_INDEX.WHITE,
    "TABLE_FOOTER":  WD_COLOR_INDEX.DARK_YELLOW,
    # Backward-compat key
    "TABLE":         WD_COLOR_INDEX.GREEN,
}

_COLOR_MAP = {
    WD_COLOR_INDEX.YELLOW:       "yellow",
    WD_COLOR_INDEX.DARK_YELLOW:  "darkYellow",
    WD_COLOR_INDEX.RED:          "red",
    WD_COLOR_INDEX.DARK_RED:     "darkRed",
    WD_COLOR_INDEX.BRIGHT_GREEN: "green",
    WD_COLOR_INDEX.GREEN:        "green",
    WD_COLOR_INDEX.TURQUOISE:    "cyan",
    WD_COLOR_INDEX.TEAL:         "darkCyan",
    WD_COLOR_INDEX.BLUE:         "blue",
    WD_COLOR_INDEX.DARK_BLUE:    "darkBlue",
    WD_COLOR_INDEX.PINK:         "magenta",
    WD_COLOR_INDEX.VIOLET:       "darkMagenta",
    WD_COLOR_INDEX.GRAY_25:      "lightGray",
    WD_COLOR_INDEX.GRAY_50:      "darkGray",
    WD_COLOR_INDEX.WHITE:        "white",
    WD_COLOR_INDEX.BLACK:        "black",
}


# =============================================================================
#  TABLE ROW-ROLE HELPERS
# =============================================================================

def _is_table_header_row(table, row_idx: int) -> bool:
    """
    Return True when row_idx is a designated header row.
    Checks the Word tblHeader XML flag; falls back to row 0.
    """
    try:
        row  = table.rows[row_idx]
        trPr = row._tr.find(qn("w:trPr"))
        if trPr is not None:
            tblHeader = trPr.find(qn("w:tblHeader"))
            if tblHeader is not None:
                val = tblHeader.get(qn("w:val"), "true")
                if val.lower() not in ("false", "0", "off"):
                    return True
    except Exception:
        pass
    return row_idx == 0


def _is_table_footer_row(table, row_idx: int, total_rows: int) -> bool:
    """Heuristic: last row whose text contains totals/summary keywords."""
    if row_idx != total_rows - 1:
        return False
    try:
        row_text = " ".join(
            cell.text for cell in table.rows[row_idx].cells
        ).lower()
        FOOTER_KW = re.compile(
            r"\b(total|sum|grand\s+total|subtotal|average|avg|count)\b",
            re.IGNORECASE,
        )
        return bool(FOOTER_KW.search(row_text))
    except Exception:
        return False


# =============================================================================
#  COVER PAGE DETECTION (standalone, usable with any Document instance)
# =============================================================================

def _detect_cover_page_body_threshold(doc) -> int:
    """
    Returns the body-child position index of the last cover-page element,
    or -1 if no cover page detected.  Works on any Document instance.
    """
    paragraphs = doc.paragraphs
    body_children    = list(doc.element.body)
    body_elem_to_pos = {id(child): pos for pos, child in enumerate(body_children)}

    cover_end = _ft_detect_cover_page(doc)
    if cover_end < 0:
        return -1

    # Map paragraph sequence index → body position
    para_seq = 0
    para_body_pos: dict = {}
    for child in body_children:
        if child.tag == qn("w:p"):
            para_body_pos[para_seq] = body_elem_to_pos[id(child)]
            para_seq += 1

    if cover_end in para_body_pos:
        return para_body_pos[cover_end]
    return -1


# =============================================================================
#  TOC DETECTION HELPERS
# =============================================================================

def _ft_get_style_base(para) -> str:
    try:
        return (para.style.name if para.style else "").split(" Char")[0].strip()
    except Exception:
        return ""


def _is_toc_entry_by_style(para) -> bool:
    return bool(re.match(r"^TOC \d+$", _style_base(para)))


_TOC_ENTRY_RE = re.compile(
    r"^\s*(\d+(?:\.\d+)*)?\s*(.+?)[\s\.…\-]{2,}(\d{1,4})\s*$", re.UNICODE)
_TOC_TAB_RE = re.compile(r"^.+\t\d{1,4}\s*$")
_FT_TOC_TAB_RE   = re.compile(r"^.+\t\d{1,4}\s*$")


def _is_toc_entry_by_heuristic(para) -> bool:
    text = (para.text or "").strip()
    if not text:
        return False
    if _TOC_ENTRY_RE.match(text) or _TOC_TAB_RE.match(text):
        return True
    xml = para._element.xml
    return "PAGEREF" in xml or ("w:instr" in xml and "TOC" in xml)


def _detect_toc_ranges(doc):
    """
    Return a set of paragraph indices that belong to the Table of Contents
    or Table of Figures sections.  Both are detected with the same logic:
    - Style-name pass: 'TOC N' / 'TOC Heading' styles
    - Heuristic pass:  text matching 'Table of Contents / Figures … page-num'
    """
    paragraphs = doc.paragraphs
    skip_indices: set = set()

    # ── Pass 1: style-based ───────────────────────────────────────────────────
    first_idx = last_idx = -1
    for i, para in enumerate(paragraphs):
        if _is_toc_entry_by_style(para):
            if first_idx == -1:
                first_idx = i
            last_idx = i
        elif first_idx != -1 and (i - last_idx) > 5:
            break

    if first_idx != -1:
        # Walk back to include the TOC/TOF title heading
        toc_start = first_idx
        for j in range(first_idx - 1, max(-1, first_idx - 5), -1):
            sb = _style_base(paragraphs[j])
            txt = (paragraphs[j].text or "").strip().lower()
            if (sb == "TOC Heading" or
                    re.match(r"^\s*(table\s+of\s+(contents|figures|tables)|contents)\s*$", txt)):
    if first_toc_style_idx != -1:
        toc_start = first_toc_style_idx
        for j in range(first_toc_style_idx - 1, max(-1, first_toc_style_idx - 5), -1):
            text = (paragraphs[j].text or "").strip().lower()
            sb   = _ft_get_style_base(paragraphs[j])
            if sb == "TOC Heading" or "table of contents" in text or text == "contents":
                toc_start = j
                break
            if txt:
                break
        for i in range(toc_start, last_idx + 1):
            skip_indices.add(i)

    # ── Pass 2: heuristic (handles docs without TOC styles) ──────────────────
    TOC_HDR_RE = re.compile(
        r"^\s*(table\s+of\s+(contents|figures|tables)|contents)\s*$", re.IGNORECASE)
    toc_start = -1
    for i, para in enumerate(paragraphs[:120]):
        sb = _style_base(para)
        txt = (para.text or "").strip()
        if (sb == "TOC Heading" or TOC_HDR_RE.match(txt) or
                ("TOC" in para._element.xml and "w:fldChar" in para._element.xml)):
    for i, para in enumerate(paragraphs[:100]):
        text = (para.text or "").strip()
        sb   = _ft_get_style_base(para)
        if sb == "TOC Heading" or TOC_TITLE_RE.match(text):
            toc_start = i
            break

    if toc_start != -1 and toc_start not in skip_indices:
        toc_end = toc_start
        consec = 0
        for i in range(toc_start + 1, min(len(paragraphs), toc_start + 300)):
            txt = (paragraphs[i].text or "").strip()
            if not txt:
                continue
            if _is_toc_entry_by_heuristic(paragraphs[i]) or _is_toc_entry_by_style(paragraphs[i]):
                toc_end = i
                consec = 0
            else:
                consec += 1
                if consec >= 3:
                    break
        for i in range(toc_start, toc_end + 1):
            skip_indices.add(i)
    toc_end = toc_start
    consec  = 0
    for i in range(toc_start + 1, min(len(paragraphs), toc_start + 300)):
        text = (paragraphs[i].text or "").strip()
        if not text:
            continue
        if (_ft_is_toc_entry_by_heuristic(paragraphs[i]) or
                _ft_is_toc_entry_by_style(paragraphs[i])):
            toc_end = i
            consec  = 0
        else:
            consec += 1
            if consec >= 3:
                break
    return toc_start, toc_end, True

    return skip_indices


def _detect_cover_end(doc) -> int:
    """Return the last paragraph index that belongs to the cover page, or -1."""
    paragraphs = doc.paragraphs

    # Strategy 1: explicit page-break run
    for i, para in enumerate(paragraphs[:60]):
        for run in para.runs:
            for br in run._element.xpath(".//w:br"):
                if br.get(qn("w:type"), "") == "page":
                    return i
        if i > 0 and has_page_break_before(para):
            return i - 1

    # Strategy 2: first Heading style signals body start
    for i, para in enumerate(paragraphs[:40]):
        sn = para.style.name if para.style else ""
        if sn.startswith("Heading") and i > 0:
            return i - 1

    # Strategy 3: keyword heuristic
    COVER_KW = re.compile(
        r"\b(prepared\s+(by|for)|author|version|copyright|confidential|restricted|"
        r"january|february|march|april|may|june|july|august|"
        r"september|october|november|december|\d{4})\b", re.IGNORECASE)
    non_empty = [(i, p) for i, p in enumerate(
        paragraphs) if (p.text or "").strip()]
        r"january|february|march|april|may|june|july|august|september|october|"
        r"november|december|\d{4})\b",
        re.IGNORECASE)
    non_empty      = [(i, p) for i, p in enumerate(paragraphs) if (p.text or "").strip()]
    cover_candidate = -1
    for rank, (i, para) in enumerate(non_empty[:25]):
        text = (para.text or "").strip()
        sn = (para.style.name if para.style else "")
        if (len(text) < 200 or para.alignment == WD_ALIGN_PARAGRAPH.CENTER or
                bool(COVER_KW.search(text)) or sn in ("Title", "Subtitle")):
            cover_candidate = i
        elif rank > 2:
            break
    return cover_candidate

# =============================================================================
#  FONT RESOLUTION
# =============================================================================


def _make_theme_resolver(doc):
    _cache = {}
    try:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        theme_part = doc.part.part_related_by(RT.THEME)
        theme_root = theme_part._element
        A = "http://schemas.openxmlformats.org/drawingml/2006/main"

        def _find(tags):
            el = theme_root
            for tag in tags:
                el = el.find(f"{{{A}}}{tag}")
                if el is None:
                    return ""
            return el.get("typeface", "") if el is not None else ""

        major = _find(["themeElements", "fontScheme", "majorFont", "latin"])
        minor = _find(["themeElements", "fontScheme", "minorFont", "latin"])
        _cache = {
            "majorHAnsi": major, "majorAscii": major,
            "majorBidi":  major, "majorEastAsia": major,
            "minorHAnsi": minor, "minorAscii": minor,
            "minorBidi":  minor, "minorEastAsia": minor,
        }
    except Exception:
        pass

    return lambda v: _cache.get(v, "")


def _get_para_font(para, theme_resolver) -> str:
    def _from_rFonts(el):
        if el is None:
            return ""
        for attr in (qn("w:ascii"), qn("w:hAnsi"), qn("w:cs")):
            v = el.get(attr, "")
            if v:
                return v
        for attr in (qn("w:asciiTheme"), qn("w:hAnsiTheme"), qn("w:cstheme")):
            v = el.get(attr, "")
            if v:
                resolved = theme_resolver(v)
                if resolved:
                    return resolved
        return ""

    try:
        run_fonts = []
        for run in para.runs:
            if not (run.text or "").strip():
                continue
            rPr = run._r.find(qn("w:rPr"))
            if rPr is not None:
                f = _from_rFonts(rPr.find(qn("w:rFonts")))
                if f:
                    run_fonts.append(f)
            if not run_fonts and run.font and run.font.name:
                run_fonts.append(run.font.name)
        if run_fonts:
            return Counter(run_fonts).most_common(1)[0][0]

        pPr = para._element.find(qn("w:pPr"))
        if pPr is not None:
            rPr = pPr.find(qn("w:rPr"))
            if rPr is not None:
                f = _from_rFonts(rPr.find(qn("w:rFonts")))
                if f:
                    return f

        style = para.style
        while style is not None:
            try:
                if style.font and style.font.name:
                    return style.font.name
                s_el = style.element
                s_pPr = s_el.find(qn("w:pPr"))
                if s_pPr is not None:
                    s_rPr = s_pPr.find(qn("w:rPr"))
                    if s_rPr is not None:
                        f = _from_rFonts(s_rPr.find(qn("w:rFonts")))
                style_elem  = style.element
                style_pPr   = style_elem.find(qn("w:pPr"))
                if style_pPr is not None:
                    style_rPr = style_pPr.find(qn("w:rPr"))
                    if style_rPr is not None:
                        f = _font_from_rFonts(style_rPr.find(qn("w:rFonts")))
                        if f:
                            return f
                s_rPr2 = s_el.find(qn("w:rPr"))
                if s_rPr2 is not None:
                    f = _from_rFonts(s_rPr2.find(qn("w:rFonts")))
                    if f:
                        return f
            except Exception:
                pass
            try:
                style = style.base_style
            except Exception:
                break
    except Exception:
        pass
    return ""

# =============================================================================
#  DOCUMENT ANALYSIS  (the main function called by /analyse)
# =============================================================================


# =============================================================================
#  DOCUMENT STRUCTURE ANALYSIS
# =============================================================================

def ft_analyze_document_structure(docx_path: str) -> dict:
    doc = Document(docx_path)
    doc            = Document(docx_path)
    footer_ids     = get_footer_paragraph_ids(doc)
    header_ids     = get_header_paragraph_ids(doc)

    elements:       list = []
    element_counts: dict = {}
    detected_types: set  = set()
    sample_texts:   dict = {}
    _font_votes:    dict = {}

    theme_resolver = _make_theme_resolver(doc)

    # Detect regions to ignore
    cover_end = _detect_cover_end(doc)
    toc_skip_indices = _detect_toc_ranges(doc)

    # Build set of paragraph element IDs inside table cells so they are not
    # double-classified in the main paragraph loop.
    table_para_ids: set = set()
    for tbl in doc.tables:
        for row in tbl.rows:
            seen_tc_ids: set = set()
            for cell in row.cells:
                tc_id = id(cell._tc)
                if tc_id in seen_tc_ids:
                    continue
                seen_tc_ids.add(tc_id)
                for cp in cell.paragraphs:
                    table_para_ids.add(id(cp._element))

    elements = []
    element_counts = {}
    detected_types = set()
    sample_texts = {}
    _font_votes = {}

    # Body paragraphs
    detected_table_bg_colors = set()

    theme_resolver    = _make_theme_resolver(doc)
    cover_end         = _ft_detect_cover_page(doc)
    toc_start, toc_end, has_toc = _ft_detect_toc_section(doc)
    first_content_idx = next(
        (i for i, p in enumerate(doc.paragraphs) if (p.text or "").strip()), 0)

    # ── body-position maps ────────────────────────────────────────────────────
    body_children    = list(doc.element.body)
    body_elem_to_pos = {id(child): pos for pos, child in enumerate(body_children)}
    para_body_pos:   dict = {}
    para_seq = 0
    for child in body_children:
        if child.tag == qn("w:p"):
            para_body_pos[para_seq] = body_elem_to_pos[id(child)]
            para_seq += 1
    cover_body_threshold = -1
    if cover_end >= 0 and cover_end in para_body_pos:
        cover_body_threshold = para_body_pos[cover_end]

    # ── Step 1: classify every top-level paragraph ────────────────────────────
    for idx, para in enumerate(doc.paragraphs):
        if is_in_footer_or_header(para):
            continue
        if not (para.text or "").strip():
            continue
        if id(para._element) in table_para_ids:
            continue
        if cover_end >= 0 and idx <= cover_end:
            continue
        if idx in toc_skip_indices:
            continue

        sb = _style_base(para)
        raw_text = (para.text or "").strip()
        ptype = None

        # 1. Prefer numbered heading FIRST
        lvl, _, _ = detect_numbered_heading(raw_text)
        if lvl:
            ptype = f"HEADING_{lvl}"

        # 2. Fallback to style
        elif sb.startswith("Heading"):
            try:
                level = max(1, min(6, int(sb.split()[-1])))
                ptype = f"HEADING_{level}"
            except Exception:
                ptype = "HEADING_1"

        if not ptype:
            if get_list_type(para):
                ptype = "LIST_ITEM"

        if not ptype:
            ptype = "PARAGRAPH"

        style_name = para.style.name if para.style is not None else ""
        style_base = style_name.split(" Char")[0].strip()
        ptype      = None
        in_toc     = has_toc and toc_start <= idx <= toc_end
        in_cover   = cover_end >= 0 and idx <= cover_end

        if style_base == "Title":
            ptype = "TITLE"
        elif style_base == "TOC Heading":
            ptype = "TOC_TITLE"
        elif style_base in ("Subtitle", "Document Map", "Author", "Date", "Company",
                            "Abstract", "Document Label", "Revision", "Version"):
            ptype = "COVER_PAGE"
        elif re.match(r"^TOC \d+$", style_base):
            try:
                level = int(style_base.split()[-1])
            except Exception:
                level = 1
            ptype = f"TOC_HEADING_{level}"

        if ptype is None and in_toc:
            text_lower = para.text.strip().lower()
            if (_ft_get_style_base(para) == "TOC Heading" or
                    re.match(r"^\s*(table\s+of\s+contents|contents)\s*$", text_lower)):
                ptype = "TOC_TITLE"
            else:
                lvl = (_ft_toc_level_from_style(para) or
                       _ft_toc_level_from_heuristic((para.text or "").strip()))
                ptype = f"TOC_HEADING_{lvl}"

        if ptype is None and in_cover and not in_toc:
            ptype = "TITLE" if idx == first_content_idx else "COVER_PAGE"

        if ptype is None:
            raw_text = (para.text or "").strip()
            if style_base.startswith("Heading"):
                try:
                    level = max(1, min(6, int(style_base.split()[-1])))
                    ptype = f"HEADING_{level}"
                except Exception:
                    ptype = "HEADING_1"
            if not ptype:
                lvl, num, title = detect_numbered_heading(raw_text)
                if lvl:
                    ptype = f"HEADING_{lvl}"
            if not ptype:
                lt = get_list_type(para)
                if lt:
                    ptype = lt
            if not ptype:
                ptype = "PARAGRAPH"

        elements.append({
            "type": ptype,
            "para_idx": idx,
            "indent": get_paragraph_indentation(para),
        })
        item = {
            "type":     ptype,
            "para_idx": idx,
            "indent":   get_paragraph_indentation(para),
            "in_cover": in_cover,   # ← store cover flag by para index
        }
        elements.append(item)
        element_counts[ptype] = element_counts.get(ptype, 0) + 1
        detected_types.add(ptype)

        if ptype not in sample_texts:
            sample_texts[ptype] = raw_text[:250]

        font = _get_para_font(para, theme_resolver)
        if font:
            _font_votes.setdefault(ptype, []).append(font)

    # Tables
    body_children = list(doc.element.body)
    body_elem_to_pos = {id(child): pos for pos,
                        child in enumerate(body_children)}

    para_body_pos: dict = {}
    para_seq = 0
    for child in body_children:
        if child.tag == qn("w:p"):
            para_body_pos[para_seq] = body_elem_to_pos[id(child)]
            para_seq += 1

    cover_body_threshold = -1
    if cover_end >= 0 and cover_end in para_body_pos:
        cover_body_threshold = para_body_pos[cover_end]

    for i, table in enumerate(doc.tables):
    # ── Step 2: classify every table AND its cell paragraphs ──────────────────
    for tbl_idx, table in enumerate(doc.tables):
        tbl_body_pos = body_elem_to_pos.get(id(table._tbl), -1)
        on_cover = (cover_body_threshold >= 0 and 0 <=
                    tbl_body_pos <= cover_body_threshold)

        elements.append(
            {"type": "TABLE", "table_idx": i, "on_cover": on_cover})
        on_cover     = (cover_body_threshold >= 0 and
                        0 <= tbl_body_pos <= cover_body_threshold)

        # --- Detect header colors from BODY tables only (skip cover) ---
        if not on_cover:
            try:
                for row_idx, row in enumerate(table.rows):
                    if _is_table_header_row(table, row_idx):
                        for cell in row.cells:
                            tcPr = cell._tc.find(qn("w:tcPr"))
                            if tcPr is not None:
                                shd = tcPr.find(qn("w:shd"))
                                if shd is not None:
                                    fill = shd.get(qn("w:fill"))
                                    if fill and fill.upper() not in ("AUTO", "FFFFFF", "000000"):
                                        detected_table_bg_colors.add(f"#{fill.upper()}")
            except Exception:
                pass

        total_rows   = len(table.rows)

        # One summary entry for the table itself (used by set_table_heading_bg)
        elements.append({
            "type":      "TABLE",
            "table_idx": tbl_idx,
            "on_cover":  on_cover,
        })
        element_counts["TABLE"] = element_counts.get("TABLE", 0) + 1

        for row_idx, row in enumerate(table.rows):
            seen_tc_ids: set = set()
            for cell_idx, cell in enumerate(row.cells):
                tc_id = id(cell._tc)
                if tc_id in seen_tc_ids:
                    continue
                seen_tc_ids.add(tc_id)

                for para_idx_in_cell, cp in enumerate(cell.paragraphs):
                    if not (cp.text or "").strip():
                        continue
                    if on_cover:
                        continue

                    elements.append({
                        "type": "TABLE_TEXT",
                        "table_idx": i,
                        "row_idx": row_idx,
                        "cell_idx": cell_idx,
                        "para_idx_in_cell": para_idx_in_cell,
                        "indent": get_paragraph_indentation(cp),
                    })
                    element_counts["TABLE_TEXT"] = element_counts.get(
                        "TABLE_TEXT", 0) + 1
                    detected_types.add("TABLE_TEXT")

                    if "TABLE_TEXT" not in sample_texts:
                        sample_texts["TABLE_TEXT"] = (
                            cp.text or "").strip()[:250]

                    font = _get_para_font(cp, theme_resolver)
                    if font:
                        _font_votes.setdefault("TABLE_TEXT", []).append(font)

    _USER_ORDER = {t: i for i, t in enumerate(USER_TYPES)}
        detected_types.add("TABLE")

        if "TABLE" not in sample_texts:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        sample_texts["TABLE"] = cell.text.strip()[:250]
                        break

        # Per-cell-paragraph entries
        seen_cell_ids: set = set()
        for row_idx, row in enumerate(table.rows):
            if _is_table_header_row(table, row_idx):
                row_role = "TABLE_HEADER"
            elif _is_table_footer_row(table, row_idx, total_rows):
                row_role = "TABLE_FOOTER"
            else:
                row_role = "TABLE_BODY"

            for col_idx, cell in enumerate(row.cells):
                cid = id(cell._tc)
                if cid in seen_cell_ids:
                    continue
                seen_cell_ids.add(cid)

                for cell_para_idx, para in enumerate(cell.paragraphs):
                    nested = False
                    node   = para._element.getparent()
                    while node is not None and node is not cell._tc:
                        if node.tag == qn("w:tbl"):
                            nested = True
                            break
                        node = node.getparent()
                    if nested:
                        continue

                    text = (para.text or "").strip()

                    cell_elem = {
                        "type":          row_role,
                        "table_idx":     tbl_idx,
                        "row_idx":       row_idx,
                        "col_idx":       col_idx,
                        "cell_para_idx": cell_para_idx,
                        "on_cover":      on_cover,
                        "_para_elem_id": id(para._element),
                    }
                    elements.append(cell_elem)

                    element_counts[row_role] = element_counts.get(row_role, 0) + 1
                    detected_types.add(row_role)

                    if text and row_role not in sample_texts:
                        sample_texts[row_role] = text[:250]

                    font_name = _ft_get_para_font(para, theme_resolver)
                    if font_name:
                        _font_votes.setdefault(row_role, []).append(font_name)

    # ── Sort key for detected_elements display list ───────────────────────────
    def _sort_key(x):
        return _USER_ORDER.get(x, 99)

    visible_types = [t for t in detected_types if t in _USER_ORDER]

        if x == "TITLE":           return (0, 0)
        if x == "COVER_PAGE":      return (1, 0)
        if x == "TOC_TITLE":       return (2, 0)
        if x.startswith("TOC_HEADING_"):
            try:    return (3, int(x.split("_")[-1]))
            except: return (3, 99)
        if x.startswith("HEADING_"):
            try:    return (4, int(x.split("_")[-1]))
            except: return (4, 99)
        if x == "LIST_ITEM":       return (5, 0)
        if x == "PARAGRAPH":       return (6, 0)
        if x == "TABLE":           return (7, 0)
        if x == "TABLE_HEADER":    return (7, 1)
        if x == "TABLE_BODY":      return (7, 2)
        if x == "TABLE_FOOTER":    return (7, 3)
        return (8, 0)

    from collections import Counter
    current_fonts = {
        pt: Counter(votes).most_common(1)[0][0]
        for pt, votes in _font_votes.items() if votes and pt in _USER_ORDER
    }

    return {
        "elements": elements,
        "detected_elements": sorted(visible_types, key=_sort_key),
        "element_counts": {k: v for k, v in element_counts.items() if k in _USER_ORDER},
        "sample_texts": sample_texts,
        "current_fonts": current_fonts,
    }
# =============================================================================
#  FORMATTING HELPERS
# =============================================================================


def _get_all_runs(para):
    """Return all <w:r> elements directly in para, excluding nested table runs."""
        "elements":          elements,
        "detected_elements": sorted(detected_types, key=_sort_key),
        "element_counts":    element_counts,
        "sample_texts":      sample_texts,
        "current_fonts":     current_fonts,
        "toc_detected":      has_toc,
        "toc_range":         [toc_start, toc_end] if has_toc else [],
        "detected_table_bgs": sorted(list(detected_table_bg_colors))
    }


# =============================================================================
#  STANDARD TABLE HEADER COLORS (COMMON DEFAULTS)
# =============================================================================

STANDARD_TABLE_BG_COLORS = [
    "#4472C4",  # Blue
    "#70AD47",  # Green
    "#FFC000",  # Gold
    "#ED7D31",  # Orange
    "#A5A5A5",  # Gray
    "#44546A",  # Dark Gray
    "#BDD7EE",  # Light Blue
    "#C6EFCE",  # Light Green
    "#FFE699",  # Light Yellow
    "#F4B084",  # Light Orange
]


# =============================================================================
#  FORMATTING HELPERS
# =============================================================================

def _ft_get_config_for_type(ptype: str, config: dict):
    """Return (font_name, font_size_pt) for a given element type."""
    font_name     = config.get(ptype.lower() + "_font")
    font_size_raw = config.get(ptype.lower() + "_size")

    if not font_name:
        if ptype == "COVER_PAGE":
            font_name = config.get("title_font") or config.get("paragraph_font")
        elif ptype == "LIST_ITEM":
            font_name = config.get("list_item_font") or config.get("paragraph_font")
        elif ptype == "TOC_TITLE":
            font_name = (config.get("toc_title_font")
                         or config.get("title_font")
                         or config.get("paragraph_font"))
        elif ptype.startswith("TOC_HEADING_"):
            font_name = config.get("toc_heading_font") or config.get("paragraph_font")
        elif ptype == "TABLE_HEADER":
            font_name = (config.get("table_header_font")
                         or config.get("table_font")
                         or config.get("paragraph_font"))
        elif ptype in ("TABLE_BODY", "TABLE_FOOTER"):
            font_name = (config.get("table_body_font")
                         or config.get("table_font")
                         or config.get("paragraph_font"))
        else:
            font_name = config.get("paragraph_font")

    if not font_size_raw:
        if ptype == "COVER_PAGE":
            font_size_raw = config.get("title_size") or 12
        elif ptype == "LIST_ITEM":
            font_size_raw = (config.get("list_item_size")
                             or config.get("paragraph_size") or 12)
        elif ptype == "TOC_TITLE":
            font_size_raw = config.get("toc_title_size") or 14
        elif ptype.startswith("TOC_HEADING_"):
            try:
                lvl = int(ptype.split("_")[-1])
            except Exception:
                lvl = 1
            font_size_raw = config.get("toc_heading_size") or max(9, 13 - lvl)
        elif ptype == "TABLE_HEADER":
            font_size_raw = (config.get("table_header_size")
                             or config.get("table_size")
                             or config.get("paragraph_size") or 11)
        elif ptype in ("TABLE_BODY", "TABLE_FOOTER"):
            font_size_raw = (config.get("table_body_size")
                             or config.get("table_size")
                             or config.get("paragraph_size") or 11)
        else:
            font_size_raw = config.get("paragraph_size") or 12

    try:
        font_size = int(font_size_raw)
    except Exception:
        font_size = 12
    return font_name, font_size


def _ft_get_all_runs(para):
    """All <w:r> elements inside para, excluding those inside nested tables."""
    all_r = para._element.xpath(".//w:r")
    nested = set()
    for tbl in para._element.xpath(".//w:tbl"):
        for r in tbl.xpath(".//w:r"):
            nested.add(id(r))
    return [r for r in all_r if id(r) not in nested]


def _apply_para_rpr(para, font_name, font_size_pt, bold, highlight_name=None):
    """Apply font/size/bold/highlight to the paragraph-level default run properties."""
def _ft_apply_para_direct_format(para, font_name, font_size_pt, bold,
                                  highlight_color_name=None):
    pPr = para._element.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    _set_rpr_font_size_bold(rPr, font_name, font_size_pt, bold, highlight_name)


def _apply_numbering_rpr(para, font_name, font_size, bold, highlight_name=None):
    """Apply font/size to the numbering run properties so bullet glyphs match."""
    if font_name:
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        for attr_name, value in (("w:ascii", font_name), ("w:hAnsi", font_name),
                                  ("w:cs", font_name), ("w:eastAsia", font_name)):
            rFonts.set(qn(attr_name), value)
        for attr in (qn("w:asciiTheme"), qn("w:hAnsiTheme"), qn("w:cstheme")):
            if rFonts.get(attr) is not None:
                del rFonts.attrib[attr]
        half_pts = str(int(font_size_pt * 2))
        for tag_name in ("w:sz", "w:szCs"):
            el = rPr.find(qn(tag_name))
            if el is None:
                el = OxmlElement(tag_name)
                rPr.append(el)
            el.set(qn("w:val"), half_pts)
    for tag_name in ("w:b", "w:bCs"):
        el = rPr.find(qn(tag_name))
        if el is None:
            el = OxmlElement(tag_name)
            rPr.append(el)
        el.set(qn("w:val"), "1" if bold else "0")
    if highlight_color_name:
        hl = rPr.find(qn("w:highlight"))
        if hl is None:
            hl = OxmlElement("w:highlight")
            rPr.append(hl)
        hl.set(qn("w:val"), highlight_color_name)


def _ft_apply_run_format(r_elem, font_name, font_size, bold, highlight_name=None):
    """Apply font / size / bold / highlight to a single <w:r> element."""
    rPr = r_elem.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        r_elem.insert(0, rPr)
    if font_name:
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        for attr_name, value in (("w:ascii", font_name), ("w:hAnsi", font_name),
                                  ("w:cs", font_name), ("w:eastAsia", font_name)):
            rFonts.set(qn(attr_name), value)
        for attr in (qn("w:asciiTheme"), qn("w:hAnsiTheme"), qn("w:cstheme")):
            if rFonts.get(attr) is not None:
                del rFonts.attrib[attr]
        half_pts = str(int(font_size) * 2)
        for tag_name in ("w:sz", "w:szCs"):
            el = rPr.find(qn(tag_name))
            if el is None:
                el = OxmlElement(tag_name)
                rPr.append(el)
            el.set(qn("w:val"), half_pts)
    for tag_name in ("w:b", "w:bCs"):
        el = rPr.find(qn(tag_name))
        if el is None:
            el = OxmlElement(tag_name)
            rPr.append(el)
        el.set(qn("w:val"), "1" if bold else "0")
    if highlight_name:
        hl = rPr.find(qn("w:highlight"))
        if hl is None:
            hl = OxmlElement("w:highlight")
            rPr.append(hl)
        hl.set(qn("w:val"), highlight_name)


def _ft_apply_numbering_runprops(para, font_name, font_size, bold=False,
                                  highlight_val=None):
    if not para._element.xpath(".//w:numPr"):
        return
    pPr = para._element.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    _set_rpr_font_size_bold(rPr, font_name, font_size, bold, highlight_name)


def _set_rpr_font_size_bold(rPr, font_name, font_size_pt, bold, highlight_name=None):
    if font_name:
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
            rPr.insert(0, rFonts)
        for attr, val in (("w:ascii", font_name), ("w:hAnsi", font_name),
                          ("w:cs", font_name), ("w:eastAsia", font_name)):
            rFonts.set(qn(attr), val)
        for attr in (qn("w:asciiTheme"), qn("w:hAnsiTheme"), qn("w:cstheme")):
            if rFonts.get(attr) is not None:
                del rFonts.attrib[attr]
        half_pts = str(int(font_size_pt) * 2)
            rPr.append(rFonts)
        for attr_name, value in (("w:ascii", font_name), ("w:hAnsi", font_name),
                                  ("w:cs", font_name), ("w:eastAsia", font_name)):
            rFonts.set(qn(attr_name), value)
        half_pts = str(int(font_size) * 2)
        for tag in ("w:sz", "w:szCs"):
            el = rPr.find(qn(tag))
            if el is None:
                el = OxmlElement(tag)
                rPr.append(el)
            el.set(qn("w:val"), half_pts)

    for tag in ("w:b", "w:bCs"):
        el = rPr.find(qn(tag))
        if el is None:
            el = OxmlElement(tag)
            rPr.append(el)
        el.set(qn("w:val"), "1" if bold else "0")

    if highlight_name:
        hl = rPr.find(qn("w:highlight"))
        if hl is None:
            hl = OxmlElement("w:highlight")
            rPr.append(hl)
        hl.set(qn("w:val"), highlight_name)


def _hex_to_rgb(hex_color: str) -> str:
    return hex_color.lstrip("#").upper()


def _is_header_row(row) -> bool:
    """
    Return True if this table row should be treated as a header.
    Checks signals in order:
      1. Word explicit tblHeader marker (<w:tblHeader/> in w:trPr).
      2. Every unique cell in the row already has a non-white / non-auto
         background fill — i.e. the author manually shaded the header row.
         This is the most common case in documents that do NOT use Word's
         built-in 'Repeat Header Rows' feature.
    """
    trPr = row._tr.find(qn("w:trPr"))
    if trPr is not None:
        if trPr.find(qn("w:tblHeader")) is not None:
            return True

    # Signal 2: every (unique) cell carries a visible fill
    unique_tcs = []
    seen_tc_ids: set = set()
    for cell in row.cells:
        tc_id = id(cell._tc)
        if tc_id not in seen_tc_ids:
            seen_tc_ids.add(tc_id)
            unique_tcs.append(cell._tc)

    if not unique_tcs:
        return False

    shaded_count = 0
    for tc in unique_tcs:
        tcPr = tc.find(qn("w:tcPr"))
        if tcPr is not None:
            shd = tcPr.find(qn("w:shd"))
            if shd is not None:
                fill = (shd.get(qn("w:fill"), "") or "").upper().strip()
                if fill and fill not in ("AUTO", "FFFFFF", ""):
                    shaded_count += 1

    # All unique cells must be shaded (partial shading = data row, not header)
    if shaded_count > 0 and shaded_count == len(unique_tcs):
        return True

    return False


def _table_has_style_first_row(table) -> bool:
    """
    Return True if the table uses a tblStyle AND tblLook has firstRow=1,
    meaning the style applies special conditional formatting to row 0.
    This lets us detect header rows whose shading comes from the style
    rather than explicit w:tcPr/w:shd.
    """
    tblPr = table._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        return False
    style_el = tblPr.find(qn("w:tblStyle"))
    if style_el is None:
        return False
    tbl_look = tblPr.find(qn("w:tblLook"))
    if tbl_look is None:
        # No tblLook means firstRow is on by default in most styles
        return True
    val = tbl_look.get(qn("w:firstRow"), "1")
    return val not in ("0", "false", "off")


def _get_header_rows(table) -> list:
    """
    Return the list of rows that constitute the header of a table.
    Detection priority:
      1. Explicit w:tblHeader markers on rows.
      2. All cells in a row have a non-white explicit fill (manual shading).
      3. Table uses a tblStyle with firstRow conditional formatting — treat
         row 0 as the header.
      4. Positional fallback: row 0.
    """
    if not table.rows:
        return []
    header_rows = [row for row in table.rows if _is_header_row(row)]
    if header_rows:
        return header_rows
    # Signal 3: style-driven first-row band
    if _table_has_style_first_row(table):
        return [table.rows[0]]
    # Positional fallback
    return [table.rows[0]]


def _apply_bg_to_row(row, rgb: str):
    """Write w:shd fill to every unique <w:tc> in a row."""
    seen_tc_ids: set = set()
    for cell in row.cells:
        tc = cell._tc
        tc_id = id(tc)
        if tc_id in seen_tc_ids:
            continue
        seen_tc_ids.add(tc_id)
        tcPr = tc.find(qn("w:tcPr"))
        if tcPr is None:
            tcPr = OxmlElement("w:tcPr")
            tc.insert(0, tcPr)
        shd = tcPr.find(qn("w:shd"))
        if shd is None:
            shd = OxmlElement("w:shd")
            tcPr.append(shd)
        shd.set(qn("w:val"),   "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"),  rgb)


def _disable_tbl_look_first_row(table):
    """
    Clear the tblLook firstRow / lastRow / firstColumn / lastColumn
    conditional-format flags so that a table style's 'Header Row' band
    cannot override the direct w:tcPr/w:shd we have just written.
    Also removes any table-level w:shd that could bleed through.
    """
    tblPr = table._tbl.find(qn("w:tblPr"))
    if tblPr is None:
        tblPr = OxmlElement("w:tblPr")
        table._tbl.insert(0, tblPr)

    # Remove table-level shading (can bleed through cell shading in some renderers)
    tbl_shd = tblPr.find(qn("w:shd"))
    if tbl_shd is not None:
        tblPr.remove(tbl_shd)

    # Neutralise tblLook conditional formatting so the style's firstRow
    # band does not override our explicit cell fill.
    tbl_look = tblPr.find(qn("w:tblLook"))
    if tbl_look is None:
        tbl_look = OxmlElement("w:tblLook")
        tblPr.append(tbl_look)
    # Keep header-row band OFF, keep the rest as-is
    tbl_look.set(qn("w:firstRow"),    "0")
    tbl_look.set(qn("w:lastRow"),     "0")
    tbl_look.set(qn("w:firstColumn"), "0")
    tbl_look.set(qn("w:lastColumn"),  "0")
    tbl_look.set(qn("w:noHBand"),     "0")
    tbl_look.set(qn("w:noVBand"),     "0")
    # val attribute is a bitmask — zero it out completely
    tbl_look.set(qn("w:val"), "0000")


def set_table_heading_bg(table, hex_color: str):
    """Apply background colour to all header rows of a table.
    Header rows are detected via tblHeader marker or existing shading first,
    with positional row-0 fallback.  Also disables tblLook conditional
    formatting so a table style cannot override the explicit cell fill."""
    if not hex_color:
        return
    rgb = _hex_to_rgb(hex_color)
    if rgb.upper() == "FFFFFF":
        return
    # Disable table-style conditional formatting BEFORE writing cell fills
    _disable_tbl_look_first_row(table)
    for row in _get_header_rows(table):
        _apply_bg_to_row(row, rgb)

# =============================================================================
#  FORMATTER
# =============================================================================

# =============================================================================
#  TOC HIGHLIGHT PASS
# =============================================================================

def ft_highlight_toc(doc, elements: list, config: dict):
    """
    Dedicated pass: stamp highlight colours onto every TOC paragraph
    including HYPERLINK field runs that python-docx .runs skips.
    """
    if not config.get("highlight"):
        return

    para_map = {e["para_idx"]: e for e in elements if "para_idx" in e}

    for idx, para in enumerate(doc.paragraphs):
        if is_in_footer_or_header(para, None, None):
            continue
        elem_info = para_map.get(idx)
        if not elem_info:
            continue
        ptype = elem_info["type"]
        if ptype != "TOC_TITLE" and not ptype.startswith("TOC_HEADING_"):
            continue

        highlight_idx  = HIGHLIGHT_COLORS.get(ptype)
        highlight_name = _COLOR_MAP.get(highlight_idx) if highlight_idx else None
        if not highlight_name:
            continue

        font_name, font_size = _ft_get_config_for_type(ptype, config)
        bold = (config.get("bold_toc", True)
                if ptype == "TOC_TITLE"
                else config.get("bold_toc_entries", False))

        _ft_apply_para_direct_format(para, font_name, font_size, bold, highlight_name)
        for r_elem in _ft_get_all_runs(para):
            _ft_apply_run_format(r_elem, font_name, font_size, bold, highlight_name)


# =============================================================================
#  TABLE CELL CONTENT FORMAT PASS
# =============================================================================

def ft_format_table_cells(doc, elements: list, config: dict):
    """
    Walk every table in the document and apply font / size / bold / highlight
    to every cell paragraph, classified by its row role.
    Skips tables on the cover page.
    """
    do_highlight = config.get("highlight", False)

    # Build a fast lookup: table_idx → on_cover flag
    tbl_cover_map: dict = {}
    for e in elements:
        if e.get("type") == "TABLE" and "table_idx" in e:
            tbl_cover_map[e["table_idx"]] = e.get("on_cover", False)

    for tbl_idx, table in enumerate(doc.tables):
        # ── FIX 3: skip cover-page tables entirely ────────────────────────────
        if tbl_cover_map.get(tbl_idx, False):
            continue

        total_rows    = len(table.rows)
        seen_cell_ids: set = set()

        for row_idx, row in enumerate(table.rows):
            if _is_table_header_row(table, row_idx):
                row_role = "TABLE_HEADER"
            elif _is_table_footer_row(table, row_idx, total_rows):
                row_role = "TABLE_FOOTER"
            else:
                row_role = "TABLE_BODY"

            font_name, font_size = _ft_get_config_for_type(row_role, config)
            bold = (row_role == "TABLE_HEADER" and
                    config.get("bold_table_header", True))

            highlight_name = None
            if do_highlight:
                hi             = HIGHLIGHT_COLORS.get(row_role)
                highlight_name = _COLOR_MAP.get(hi) if hi else None

            for col_idx, cell in enumerate(row.cells):
                cid = id(cell._tc)
                if cid in seen_cell_ids:
                    continue
                seen_cell_ids.add(cid)

                for para in cell.paragraphs:
                    nested = False
                    node   = para._element.getparent()
                    while node is not None and node is not cell._tc:
                        if node.tag == qn("w:tbl"):
                            nested = True
                            break
                        node = node.getparent()
                    if nested:
                        continue

                    _ft_apply_para_direct_format(
                        para, font_name, font_size, bold, highlight_name)
                    for r_elem in _ft_get_all_runs(para):
                        _ft_apply_run_format(
                            r_elem, font_name, font_size, bold, highlight_name)


# =============================================================================
#  MAIN FORMAT FUNCTION
# =============================================================================

def ft_format_docx(input_path: str, elements: list, output_path: str, config: dict):
    doc = Document(input_path)

    # Body paragraph lookup
    para_map = {}
    for e in elements:
        if e.get("type") == "TABLE_TEXT":
    doc        = Document(input_path)
    footer_ids = get_footer_paragraph_ids(doc)
    header_ids = get_header_paragraph_ids(doc)

    # ── Recompute cover-page body threshold for THIS Document instance ────────
    # (elements list was built from a different Document open; id()-based
    #  on_cover flags for TABLE entries are stale.  We re-derive it here.)
    cover_body_threshold = _detect_cover_page_body_threshold(doc)
    body_children    = list(doc.element.body)
    body_elem_to_pos = {id(child): pos for pos, child in enumerate(body_children)}

    para_map = {e["para_idx"]: e for e in elements if "para_idx" in e}

    # ── Pass 1: top-level (non-table) paragraphs ──────────────────────────────
    for idx, para in enumerate(doc.paragraphs):
        if is_in_footer_or_header(para, footer_ids, header_ids):
            continue
        if "para_idx" in e and e["para_idx"] >= 0:
            para_map[e["para_idx"]] = e

    # Stable table-text lookup
    table_text_map: dict = {}
    for e in elements:
        if e.get("type") == "TABLE_TEXT":
            key = (
                e.get("table_idx"),
                e.get("row_idx"),
                e.get("cell_idx"),
                e.get("para_idx_in_cell"),
            )
            table_text_map[key] = e

    highlight_on = config.get("highlight", False)

    def _apply_formatting(para, elem_info):
        ptype = elem_info["type"]
        indent = elem_info.get("indent", 0.0)

        key = ptype.lower()
        font_name = config.get(f"{key}_font") or config.get(
            "global_font") or None
        size_raw = config.get(f"{key}_size") or config.get("global_size") or 12
        try:
            font_size = int(size_raw)
        except Exception:
            font_size = 12

        is_heading = ptype.startswith("HEADING_")
        bold = (
            (is_heading and config.get("bold_headings", True)) or
            (ptype == "LIST_ITEM" and config.get("bold_lists", False))
        elem_info = para_map.get(idx)
        if not elem_info:
            continue

        # ── FIX 3: skip cover-page paragraphs entirely ────────────────────────
        if elem_info.get("in_cover", False):
            continue

        ptype           = elem_info["type"]
        original_indent = elem_info.get("indent", 0.0)
        font_name, font_size = _ft_get_config_for_type(ptype, config)

        is_heading   = ptype.startswith("HEADING_")
        is_title     = ptype == "TITLE"
        is_toc_title = ptype == "TOC_TITLE"
        is_toc_entry = ptype.startswith("TOC_HEADING_")
        is_cover     = ptype == "COVER_PAGE"
        is_list      = ptype == "LIST_ITEM"

        bold_run = (
            (is_heading   and config.get("bold_headings",    True))  or
            (is_title     and config.get("bold_titles",      True))  or
            (is_cover     and config.get("bold_titles",      True))  or
            (is_toc_title and config.get("bold_toc",         True))  or
            (is_toc_entry and config.get("bold_toc_entries", False)) or
            (is_list      and config.get("bold_lists",       False))
        )

        hl_idx = HIGHLIGHT_COLORS.get(ptype) if highlight_on else None
        hl_name = _COLOR_MAP.get(hl_idx) if hl_idx else None

        # TOC highlight is handled by ft_highlight_toc(); skip here
        if is_toc_title or is_toc_entry:
            highlight_name = None
        else:
            hi             = HIGHLIGHT_COLORS.get(ptype) if config.get("highlight") else None
            highlight_name = _COLOR_MAP.get(hi) if hi else None

        _apply_para_rpr(para, font_name, font_size, bold, hl_name)
        _ft_apply_para_direct_format(para, font_name, font_size, bold_run, highlight_name)

        if ptype == "LIST_ITEM":
            _apply_numbering_rpr(para, font_name, font_size, bold, hl_name)
        if is_list:
            _ft_apply_numbering_runprops(para, font_name, font_size, bold_run, highlight_name)

        for r_elem in _get_all_runs(para):
            rPr = r_elem.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                r_elem.insert(0, rPr)
            _set_rpr_font_size_bold(rPr, font_name, font_size, bold, hl_name)
        for r_elem in _ft_get_all_runs(para):
            _ft_apply_run_format(r_elem, font_name, font_size, bold_run,
                                 highlight_name if not (is_toc_title or is_toc_entry) else None)

        if indent > 0:
            try:
                para.paragraph_format.left_indent = Inches(indent)
            except Exception:
                pass

    # Pass 1: body paragraphs
    for idx, para in enumerate(doc.paragraphs):
        if is_in_footer_or_header(para):
            continue
        elem_info = para_map.get(idx)
        if not elem_info:
            continue
        _apply_formatting(para, elem_info)

    # Pass 2: table cell paragraphs
    for table_idx, table in enumerate(doc.tables):
        for row_idx, row in enumerate(table.rows):
            seen_tc_ids: set = set()
            for cell_idx, cell in enumerate(row.cells):
                tc_id = id(cell._tc)
                if tc_id in seen_tc_ids:
                    continue
                seen_tc_ids.add(tc_id)

                for para_idx_in_cell, cp in enumerate(cell.paragraphs):
                    if not (cp.text or "").strip():
                        continue

                    elem_info = table_text_map.get((
                        table_idx,
                        row_idx,
                        cell_idx,
                        para_idx_in_cell,
                    ))
                    if not elem_info:
                        continue

                    _apply_formatting(cp, elem_info)

    # Table header background
    tbl_header_bg = config.get("table_heading_bg")
    if tbl_header_bg:
        # Build set of cover-table indices recorded during analysis.
        # These use the same enumerate(doc.tables) order so index == position.
        cover_table_indices: set = {
            e["table_idx"]
            for e in elements
            if e.get("type") == "TABLE" and e.get("on_cover", False)
        }
        for i, table in enumerate(doc.tables):
            if i not in cover_table_indices:
                set_table_heading_bg(table, tbl_header_bg)
    # ── Pass 2: table heading background colour ────────────────────────────────
    table_heading_bg = config.get("table_heading_bg")
    if table_heading_bg:
        rgb = table_heading_bg.lstrip("#").upper()
        for tbl_idx, table in enumerate(doc.tables):
            # ── FIX 2 + FIX 3: recompute on_cover using THIS doc's element ids ─
            tbl_body_pos = body_elem_to_pos.get(id(table._tbl), -1)
            on_cover = (cover_body_threshold >= 0 and
                        0 <= tbl_body_pos <= cover_body_threshold)
            if on_cover:
                continue  # never touch cover-page tables

            # Apply background to ALL header rows (not just row 0)
            try:
                for row_idx, row in enumerate(table.rows):
                    if not _is_table_header_row(table, row_idx):
                        continue
                    for cell in row.cells:
                        tc   = cell._tc
                        tcPr = tc.find(qn("w:tcPr"))
                        if tcPr is None:
                            tcPr = OxmlElement("w:tcPr")
                            tc.insert(0, tcPr)
                        shd = tcPr.find(qn("w:shd"))
                        if shd is None:
                            shd = OxmlElement("w:shd")
                            tcPr.append(shd)
                        shd.set(qn("w:val"),   "clear")
                        shd.set(qn("w:color"), "auto")
                        shd.set(qn("w:fill"),  rgb)
            except Exception:
                pass

    # ── Pass 3: format every cell paragraph (font / size / bold / highlight) ──
    ft_format_table_cells(doc, elements, config)

    # ── Pass 4: TOC highlight (after main loop so field runs are covered) ──────
    ft_highlight_toc(doc, elements, config)

    doc.save(output_path)

# =============================================================================
#  HTML PREVIEW  (renders docx to HTML for the right-pane iframe)
#  HTML PREVIEW
# =============================================================================


def docx_to_html_preview(docx_path: str, elements: list = None) -> str:
    """
    Render a .docx as an HTML page for browser preview.
    When *elements* is supplied, TOC and table rows are colour-coded.
    """
    import html as _html_mod
    doc = Document(docx_path)
    doc        = Document(docx_path)
    footer_ids = get_footer_paragraph_ids(doc)
    header_ids = get_header_paragraph_ids(doc)

    css = (
    para_type_map: dict = {}
    if elements:
        for e in elements:
            if "para_idx" in e:
                para_type_map[e["para_idx"]] = e["type"]

    page_css = (
        "<!DOCTYPE html><html><head><meta charset='utf-8'><style>"
        "body{font-family:Arial,sans-serif;margin:0;padding:0;background:#e8e8e8;}"
        ".page{background:white;width:794px;min-height:1123px;margin:30px auto;"
        "padding:72px 80px;box-shadow:0 2px 12px rgba(0,0,0,0.18);box-sizing:border-box;}"
        "table{border-collapse:collapse;width:100%;margin:12px 0;}"
        "th,td{border:1px solid #ccc;padding:6px 10px;text-align:left;font-size:13px;}"
        "th{font-weight:600;}"
        "h1,h2,h3,h4,h5,h6{margin:0.5em 0 0.25em;}"
        "p{margin:0.35em 0;line-height:1.55;}"
        "ul,ol{margin:0.5em 0 0.5em 1.8em;padding:0;}"
        "li{margin:0.2em 0;line-height:1.5;}"
        "hr.page-break{border:none;border-top:2px dashed #bbb;margin:24px 0;}"
        ".toc-title{background:#ffe0e0;font-weight:700;padding:4px 8px;"
        "border-left:4px solid #d44;margin:2px 0;}"
        ".toc-1{background:#f0f0f0;padding-left:0px;margin:1px 0;}"
        ".toc-2{background:#e4e4e4;padding-left:18px;margin:1px 0;}"
        ".toc-3{background:#d4e2f5;padding-left:36px;margin:1px 0;}"
        ".toc-4{background:#b8cce8;padding-left:54px;margin:1px 0;}"
        ".toc-5{background:#9fd5d5;padding-left:72px;margin:1px 0;}"
        ".toc-6{background:#7ecece;padding-left:90px;margin:1px 0;}"
        ".tbl-header-row th,.tbl-header-row td"
        "{background:#c6efce;font-weight:700;color:#276221;}"
        ".tbl-footer-row td{background:#ffeb9c;font-style:italic;color:#7a5c00;}"
        ".tbl-body-row td{background:#ffffff;}"
        "</style></head><body><div class='page'>"
    )
    parts = [css]
    para_map = {id(p._element): p for p in doc.paragraphs}
    table_map = {id(t._tbl): t for t in doc.tables}
    parts = [page_css]

    para_map       = {id(p._element): p for p in doc.paragraphs}
    table_map      = {id(t._tbl):     t for t in doc.tables}
    para_seq_index = {id(p._element): i for i, p in enumerate(doc.paragraphs)}

    def _render_runs(para) -> str:
        out = []
        for run in para.runs:
            raw = run.text or ""
            t = _html_mod.escape(str(raw))
            raw = run.text or ""
            t   = _html_mod.escape(str(raw))
            rPr = run._r.find(qn("w:rPr"))
            bold = ital = uline = False
            color = ""
            if rPr is not None:
                bold = rPr.find(qn("w:b")) is not None
                ital = rPr.find(qn("w:i")) is not None
                uline = rPr.find(qn("w:u")) is not None
                bold      = rPr.find(qn("w:b")) is not None
                ital      = rPr.find(qn("w:i")) is not None
                underline = rPr.find(qn("w:u")) is not None
                c = rPr.find(qn("w:color"))
                if c is not None:
                    val = c.get(qn("w:val"), "")
                    if val and val.upper() not in ("AUTO", ""):
                        color = f"color:#{val};"
            if bold:
                t = f"<strong>{t}</strong>"
            if ital:
                t = f"<em>{t}</em>"
            if uline:
                t = f"<u>{t}</u>"
            if color:
                t = f'<span style="{color}">{t}</span>'
            if bold:      t = f"<strong>{t}</strong>"
            if ital:      t = f"<em>{t}</em>"
            if underline: t = f"<u>{t}</u>"
            if color:     t = f'<span style="{color}">{t}</span>'
            out.append(t)
        return "".join(out)

    def _style_tag(sn: str) -> str:
        sn = sn.lower()
        for i in range(1, 7):
            if f"heading {i}" in sn:
                return f"h{i}"
        return "p"

    def _para_css(para) -> str:
        css = []
        if para.alignment == WD_ALIGN_PARAGRAPH.CENTER:
            css.append("text-align:center")
        elif para.alignment == WD_ALIGN_PARAGRAPH.RIGHT:
            css.append("text-align:right")
        elif para.alignment == WD_ALIGN_PARAGRAPH.JUSTIFY:
            css.append("text-align:justify")
        if para.paragraph_format.left_indent:
            try:
                css.append(
                    f"padding-left:{para.paragraph_format.left_indent.pt:.0f}pt")
            except Exception:
                pass
        return ";".join(css)

    def _toc_html_class(ptype: str) -> str:
        if ptype == "TOC_TITLE":
            return "toc-title"
        if ptype.startswith("TOC_HEADING_"):
            try:
                lvl = int(ptype.split("_")[-1])
                return f"toc-{lvl}"
            except Exception:
                return "toc-1"
        return ""

    def _render_table(table) -> str:
        rows = ["<table>"]
        header_row_ids = {id(r._tr) for r in _get_header_rows(table)}
        for row in table.rows:
            rows.append("<tr>")
            is_header = id(row._tr) in header_row_ids
            seen_tc_ids: set = set()
            for cell in row.cells:
                tc_id = id(cell._tc)
                if tc_id in seen_tc_ids:
                    continue
                seen_tc_ids.add(tc_id)
                parts = []
                for cp in cell.paragraphs:
                    if is_in_footer_or_header(cp):
                        continue
                    r = _render_runs(cp)
                    if r is not None:
                        parts.append(r)
                content = "<br>".join(p for p in parts if p) or "&nbsp;"
                tag = "th" if is_header else "td"
                cell_sty = ""
        total_rows = len(table.rows)
        rows_html  = ["<table>"]
        seen_cell_ids: set = set()

        for row_idx, row in enumerate(table.rows):
            if _is_table_header_row(table, row_idx):
                tr_class = "tbl-header-row"
            elif _is_table_footer_row(table, row_idx, total_rows):
                tr_class = "tbl-footer-row"
            else:
                tr_class = "tbl-body-row"

            rows_html.append(f'<tr class="{tr_class}">')

            for col_idx, cell in enumerate(row.cells):
                cid = id(cell._tc)
                if cid in seen_cell_ids:
                    continue
                seen_cell_ids.add(cid)

                cell_parts = []
                for cpara in cell.paragraphs:
                    if is_in_footer_or_header(cpara, footer_ids, header_ids):
                        continue
                    rendered = _render_runs(cpara)
                    cell_parts.append(rendered)
                cell_content = "<br>".join(p for p in cell_parts if p) or "&nbsp;"

                tag_name   = "th" if tr_class == "tbl-header-row" else "td"
                cell_style = ""
                tcPr = cell._tc.find(qn("w:tcPr"))
                if tcPr is not None:
                    shd  = tcPr.find(qn("w:shd"))
                    if shd is not None:
                        fill = shd.get(qn("w:fill"), "") or ""
                        if fill and fill.upper() not in ("", "AUTO", "FFFFFF"):
                            cell_sty = f' style="background:#{fill};color:white;"'
                rows.append(f"<{tag}{cell_sty}>{content}</{tag}>")
            rows.append("</tr>")
        rows.append("</table>")
        return "".join(rows)
                            cell_style = f' style="background:#{fill};"'

                rows_html.append(f"<{tag_name}{cell_style}>{cell_content}</{tag_name}>")

            rows_html.append("</tr>")

        rows_html.append("</table>")
        return "".join(rows_html)

    in_list = in_num = False
    in_list     = False
    in_numbered = False

    def close_list():
        nonlocal in_list, in_num
        if in_list:
            parts.append("</ol>" if in_num else "</ul>")
            in_list = in_num = False
            parts.append("</ol>" if in_numbered else "</ul>")
            in_list     = False
            in_numbered = False

    for child in doc.element.body:
        tag = child.tag
        if tag == qn("w:p"):
            para = para_map.get(id(child))
            if para is None or is_in_footer_or_header(para):
                continue
            text   = para.text or ""
            has_pb = has_page_break_before(para)
            if has_pb:
                close_list()
                parts.append('<hr class="page-break">')
            if not str(text).strip():
                if not has_pb:
                    close_list()
                    parts.append("<p>&nbsp;</p>")
                continue
            sn = (para.style.name if para.style else None) or ""
            is_list_para = bool(para._element.xpath(".//w:numPr"))
            fmt = _resolve_numFmt(para) if is_list_para else ""
            is_num_list = fmt in _NUMBERED_FMTS if fmt else False
            htag = _style_tag(sn)
            block_css = _para_css(para)
            style_attr = f' style="{block_css}"' if block_css else ""
            inner = _render_runs(para)
            if is_list_para:

            style_name   = (para.style.name if para.style else None) or ""
            seq_idx      = para_seq_index.get(id(para._element), -1)
            ptype        = para_type_map.get(seq_idx, "")
            is_para_list = bool(para._element.xpath(".//w:numPr"))
            fmt          = _resolve_numFmt(para) if is_para_list else ""
            is_num       = fmt in _NUMBERED_FMTS if fmt else False
            toc_cls      = _toc_html_class(ptype)
            htag         = _style_tag(style_name)
            block_css    = _para_css(para)
            style_attr   = f' style="{block_css}"' if block_css else ""
            inner        = _render_runs(para)

            if toc_cls:
                close_list()
                parts.append(f'<p class="{toc_cls}"{style_attr}>{inner}</p>')
            elif is_para_list:
                if not in_list:
                    parts.append("<ol>" if is_num_list else "<ul>")
                    in_list = True
                    in_num = is_num_list
                elif is_num_list != in_num:
                    parts.append("</ol>" if in_num else "</ul>")
                    parts.append("<ol>" if is_num_list else "<ul>")
                    in_num = is_num_list
                    parts.append("<ol>" if is_num else "<ul>")
                    in_list     = True
                    in_numbered = is_num
                elif is_num != in_numbered:
                    parts.append("</ol>" if in_numbered else "</ul>")
                    parts.append("<ol>" if is_num else "<ul>")
                    in_numbered = is_num
                parts.append(f"<li{style_attr}>{inner}</li>")
            else:
                close_list()
                parts.append(f"<{htag}{style_attr}>{inner}</{htag}>")

        elif tag == qn("w:tbl"):
            close_list()
            table = table_map.get(id(child))
            if table is not None:
                parts.append(_render_table(table))

    close_list()
    parts.append("</div></body></html>")
    return "".join(parts)

# =============================================================================
#  FLASK ROUTES
#  FLASK ROUTES
# =============================================================================


@font_bp.route("/")
def home():
    return render_template("fontUI.html")


@font_bp.route("/analyse", methods=["POST"])
def ft_analyse():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file in session. Please upload a file first."}), 400
    try:
        data = ft_analyze_document_structure(path)
        data["standard_table_bgs"] = STANDARD_TABLE_BG_COLORS
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@font_bp.route("/original", methods=["POST"])
def ft_original():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file in session."}), 400
    try:
        if IS_WINDOWS and WORD_AVAILABLE:
            pdf_path = os.path.join(
                _session_out_dir(), f"orig_{uuid.uuid4().hex}.pdf")
            pythoncom.CoInitialize()
            try:
                convert(path, pdf_path)
            finally:
                pythoncom.CoUninitialize()
            return send_file(pdf_path, mimetype="application/pdf")
        html = docx_to_html_preview(path)
        return html, 200, {"Content-Type": "text/html; charset=utf-8"}

        html_content = docx_to_html_preview(path)
        return html_content, 200, {"Content-Type": "text/html; charset=utf-8"}
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@font_bp.route("/preview", methods=["POST"])
def ft_preview():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file in session."}), 400
        return jsonify({"error": "No file uploaded yet. Please upload from the home page."}), 400
    try:
        cfg      = json.loads(request.form.get("config", "{}"))
        out_docx = os.path.join(
            _session_out_dir(), f"ft_preview_{uuid.uuid4().hex}.docx")
        data = ft_analyze_document_structure(path)
        ft_format_docx(path, data["elements"], out_docx, cfg)
        if IS_WINDOWS and WORD_AVAILABLE:
            pdf_path = os.path.join(
                _session_out_dir(), f"ft_prev_{uuid.uuid4().hex}.pdf")
            pythoncom.CoInitialize()
            try:
                convert(out_docx, pdf_path)
            finally:
                pythoncom.CoUninitialize()
            return send_file(pdf_path, mimetype="application/pdf")
        html = docx_to_html_preview(out_docx)
        return html, 200, {"Content-Type": "text/html; charset=utf-8"}

        html_content = docx_to_html_preview(out_docx, elements=data["elements"])
        return html_content, 200, {"Content-Type": "text/html; charset=utf-8"}
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@font_bp.route("/format", methods=["POST"])
def ft_format():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file in session."}), 400
        return jsonify({"error": "No file uploaded yet. Please upload from the home page."}), 400
    try:
        cfg = json.loads(request.form.get("config", "{}"))

        # Do not persist highlight in the saved/formatted document
        cfg["highlight"] = False

        out_path = os.path.join(
            _session_out_dir(), f"font_formatted_{uuid.uuid4().hex}.docx")
        data = ft_analyze_document_structure(path)
        cfg      = json.loads(request.form.get("config", "{}"))
        out_name = f"font_formatted_{uuid.uuid4().hex}.docx"
        out_path = os.path.join(_session_out_dir(), out_name)
        data     = ft_analyze_document_structure(path)
        ft_format_docx(path, data["elements"], out_path, cfg)
        _set_path(out_path)
        return send_file(out_path, as_attachment=True, download_name="font_formatted.docx")
    except Exception as e:
        return jsonify({"error": str(e)}), 500


# =============================================================================
#  STANDALONE RUNNER
# =============================================================================

if __name__ == "__main__":
    app = Flask(__name__, template_folder="templates")
    app.secret_key = "dev"
    app.register_blueprint(font_bp)

    @app.route("/")
    def root():
        from flask import redirect
        return redirect("/font/")

    print("Font Formatter → http://127.0.0.1:5001/font/")
    app.run(debug=True, port=5001, threaded=False)