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
    try:
        return (para.style.name if para.style else "").split(" Char")[0].strip()
    except Exception:
        return ""


def _is_toc_entry_by_style(para) -> bool:
    return bool(re.match(r"^TOC \d+$", _style_base(para)))


_TOC_ENTRY_RE = re.compile(
    r"^\s*(\d+(?:\.\d+)*)?\s*(.+?)[\s\.…\-]{2,}(\d{1,4})\s*$", re.UNICODE)
_TOC_TAB_RE = re.compile(r"^.+\t\d{1,4}\s*$")


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


def ft_analyze_document_structure(docx_path: str) -> dict:
    doc = Document(docx_path)

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

        elements.append({
            "type": ptype,
            "para_idx": idx,
            "indent": get_paragraph_indentation(para),
        })
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
        tbl_body_pos = body_elem_to_pos.get(id(table._tbl), -1)
        on_cover = (cover_body_threshold >= 0 and 0 <=
                    tbl_body_pos <= cover_body_threshold)

        elements.append(
            {"type": "TABLE", "table_idx": i, "on_cover": on_cover})
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

    def _sort_key(x):
        return _USER_ORDER.get(x, 99)

    visible_types = [t for t in detected_types if t in _USER_ORDER]

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
    all_r = para._element.xpath(".//w:r")
    nested = set()
    for tbl in para._element.xpath(".//w:tbl"):
        for r in tbl.xpath(".//w:r"):
            nested.add(id(r))
    return [r for r in all_r if id(r) not in nested]


def _apply_para_rpr(para, font_name, font_size_pt, bold, highlight_name=None):
    """Apply font/size/bold/highlight to the paragraph-level default run properties."""
    pPr = para._element.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    _set_rpr_font_size_bold(rPr, font_name, font_size_pt, bold, highlight_name)


def _apply_numbering_rpr(para, font_name, font_size, bold, highlight_name=None):
    """Apply font/size to the numbering run properties so bullet glyphs match."""
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


def ft_format_docx(input_path: str, elements: list, output_path: str, config: dict):
    doc = Document(input_path)

    # Body paragraph lookup
    para_map = {}
    for e in elements:
        if e.get("type") == "TABLE_TEXT":
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
        )

        hl_idx = HIGHLIGHT_COLORS.get(ptype) if highlight_on else None
        hl_name = _COLOR_MAP.get(hl_idx) if hl_idx else None

        _apply_para_rpr(para, font_name, font_size, bold, hl_name)

        if ptype == "LIST_ITEM":
            _apply_numbering_rpr(para, font_name, font_size, bold, hl_name)

        for r_elem in _get_all_runs(para):
            rPr = r_elem.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                r_elem.insert(0, rPr)
            _set_rpr_font_size_bold(rPr, font_name, font_size, bold, hl_name)

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

    doc.save(output_path)

# =============================================================================
#  HTML PREVIEW  (renders docx to HTML for the right-pane iframe)
# =============================================================================


def docx_to_html_preview(docx_path: str) -> str:
    import html as _html_mod
    doc = Document(docx_path)

    css = (
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
        "</style></head><body><div class='page'>"
    )
    parts = [css]
    para_map = {id(p._element): p for p in doc.paragraphs}
    table_map = {id(t._tbl): t for t in doc.tables}

    def _render_runs(para) -> str:
        out = []
        for run in para.runs:
            raw = run.text or ""
            t = _html_mod.escape(str(raw))
            rPr = run._r.find(qn("w:rPr"))
            bold = ital = uline = False
            color = ""
            if rPr is not None:
                bold = rPr.find(qn("w:b")) is not None
                ital = rPr.find(qn("w:i")) is not None
                uline = rPr.find(qn("w:u")) is not None
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
                tcPr = cell._tc.find(qn("w:tcPr"))
                if tcPr is not None:
                    shd = tcPr.find(qn("w:shd"))
                    if shd is not None:
                        fill = shd.get(qn("w:fill"), "") or ""
                        if fill and fill.upper() not in ("", "AUTO", "FFFFFF"):
                            cell_sty = f' style="background:#{fill};color:white;"'
                rows.append(f"<{tag}{cell_sty}>{content}</{tag}>")
            rows.append("</tr>")
        rows.append("</table>")
        return "".join(rows)

    in_list = in_num = False

    def close_list():
        nonlocal in_list, in_num
        if in_list:
            parts.append("</ol>" if in_num else "</ul>")
            in_list = in_num = False

    for child in doc.element.body:
        tag = child.tag
        if tag == qn("w:p"):
            para = para_map.get(id(child))
            if para is None or is_in_footer_or_header(para):
                continue
            text = para.text or ""
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
                if not in_list:
                    parts.append("<ol>" if is_num_list else "<ul>")
                    in_list = True
                    in_num = is_num_list
                elif is_num_list != in_num:
                    parts.append("</ol>" if in_num else "</ul>")
                    parts.append("<ol>" if is_num_list else "<ul>")
                    in_num = is_num_list
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
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@font_bp.route("/preview", methods=["POST"])
def ft_preview():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file in session."}), 400
    try:
        cfg = json.loads(request.form.get("config", "{}"))
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
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@font_bp.route("/format", methods=["POST"])
def ft_format():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file in session."}), 400
    try:
        cfg = json.loads(request.form.get("config", "{}"))

        # Do not persist highlight in the saved/formatted document
        cfg["highlight"] = False

        out_path = os.path.join(
            _session_out_dir(), f"font_formatted_{uuid.uuid4().hex}.docx")
        data = ft_analyze_document_structure(path)
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
