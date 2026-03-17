import os
import re
import json
import uuid
import platform

from flask import (Blueprint, Flask, request, jsonify,
                   send_file, render_template, session, current_app)
from werkzeug.utils import secure_filename

from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_COLOR_INDEX, WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# ── Platform / optional deps ─────────────────────────────────────────────────
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
font_bp = Blueprint(
    "font",
    __name__,
    url_prefix="/font",
    template_folder="templates"
)

# ── Folders ───────────────────────────────────────────────────────────────────
OUTPUT = "outputs/font"
os.makedirs(OUTPUT, exist_ok=True)


# ── Session store helpers ─────────────────────────────────────────────────────

def _get_path():
    """Return the current working-file path for this session, or None."""
    sid = session.get("sid")
    if not sid:
        return None
    return current_app.config["GET_WORKING_PATH"](sid)


def _set_path(path: str):
    """Update the working-file path for this session."""
    sid = session.get("sid")
    if sid:
        current_app.config["SET_WORKING_PATH"](sid, path)


def _session_out_dir():
    sid = session.get("sid", "default")
    d = os.path.join(OUTPUT, sid)
    os.makedirs(d, exist_ok=True)
    return d


# =============================================================================
#  ALL UTILITY FUNCTIONS (unchanged from original font1.py)
#  — copied verbatim so this file is self-contained.
# =============================================================================

def is_in_footer_or_header(para, *args) -> bool:
    _FH_TAGS = {qn("w:hdr"), qn("w:ftr")}
    node = para._element.getparent()
    while node is not None:
        if node.tag in _FH_TAGS:
            return True
        node = node.getparent()
    return False


def get_footer_paragraph_ids(doc) -> set:
    return set()


def get_header_paragraph_ids(doc) -> set:
    return set()


def get_paragraph_indentation(para) -> float:
    try:
        if para.paragraph_format.left_indent:
            return para.paragraph_format.left_indent.inches
        return 0.0
    except Exception:
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


NUMBERED_HEADING_RE = re.compile(r"^\s*(\d+(?:\.\d+)+)\s+(.*\S.*)$")


def detect_numbered_heading(text: str):
    if not text:
        return None, None, None
    m = NUMBERED_HEADING_RE.match(text)
    if not m:
        return None, None, None
    num = m.group(1)
    title = m.group(2).strip()
    if len(num.split(".")) < 2:
        return None, None, None
    level = min(6, len(num.split(".")))
    return level, num, title


def clear_paragraph_runs(para):
    p = para._p
    for r in list(para.runs):
        try:
            p.remove(r._r)
        except Exception:
            pass


_NUMBERED_FMTS = frozenset({
    "decimal", "lowerLetter", "upperLetter", "lowerRoman",
    "upperRoman", "ordinal", "cardinalText", "decimalZero",
    "decimalEnclosedCircle", "decimalFullWidth", "decimalHalfWidth",
    "japaneseKounting", "koreanDigital", "russianLower", "russianUpper",
})


def _resolve_numFmt(para) -> str:
    try:
        numbering_part = para.part.numbering_part
    except Exception:
        numbering_part = None

    def _fmt_from_numId_ilvl(numId_val: str, ilvl_val: str) -> str:
        if numId_val == "0":
            return ""
        if numbering_part is None:
            return "bullet"
        nbr = numbering_part._element
        abs_id_nodes = nbr.xpath(
            f".//w:num[@w:numId='{numId_val}']/w:abstractNumId")
        if not abs_id_nodes:
            return "bullet"
        abstract_id = abs_id_nodes[0].get(qn("w:val"), "0")
        fmt_nodes = nbr.xpath(
            f".//w:abstractNum[@w:abstractNumId='{abstract_id}']"
            f"/w:lvl[@w:ilvl='{ilvl_val}']/w:numFmt")
        if fmt_nodes:
            return fmt_nodes[0].get(qn("w:val"), "")
        return "bullet"

    pPr = para._element.find(qn("w:pPr"))
    if pPr is not None:
        numPr = pPr.find(qn("w:numPr"))
        if numPr is not None:
            numId_node = numPr.find(qn("w:numId"))
            ilvl_node = numPr.find(qn("w:ilvl"))
            numId_val = numId_node.get(
                qn("w:val"), "0") if numId_node is not None else "0"
            ilvl_val = ilvl_node.get(
                qn("w:val"),  "0") if ilvl_node is not None else "0"
            if numId_val == "0":
                return ""
            return _fmt_from_numId_ilvl(numId_val, ilvl_val)

    style = para.style if para.style else None
    while style is not None:
        try:
            style_elem = style.element
            style_pPr = style_elem.find(qn("w:pPr"))
            if style_pPr is not None:
                style_numPr = style_pPr.find(qn("w:numPr"))
                if style_numPr is not None:
                    numId_node = style_numPr.find(qn("w:numId"))
                    ilvl_node = style_numPr.find(qn("w:ilvl"))
                    numId_val = numId_node.get(
                        qn("w:val"), "0") if numId_node is not None else "0"
                    ilvl_val = ilvl_node.get(
                        qn("w:val"),  "0") if ilvl_node is not None else "0"
                    if numId_val == "0":
                        return ""
                    return _fmt_from_numId_ilvl(numId_val, ilvl_val)
        except Exception:
            pass
        try:
            style = style.base_style
        except Exception:
            break
    return ""


def get_list_type(para) -> str:
    try:
        style_name = para.style.name if para.style else ""
    except Exception:
        style_name = ""
    if "List Bullet" in style_name:
        return "LIST_ITEM"
    if "List Number" in style_name:
        return "LIST_ITEM"
    if style_name.startswith("List ") or style_name == "List":
        return "LIST_ITEM"
    fmt = _resolve_numFmt(para)
    if not fmt:
        return None
    return "LIST_ITEM"


HIGHLIGHT_COLORS = {
    "TITLE":         WD_COLOR_INDEX.YELLOW,
    "COVER_PAGE":    WD_COLOR_INDEX.DARK_YELLOW,
    "TOC_TITLE":     WD_COLOR_INDEX.RED,
    "TOC_HEADING_1": WD_COLOR_INDEX.GRAY_25,
    "TOC_HEADING_2": WD_COLOR_INDEX.GRAY_50,
    "TOC_HEADING_3": WD_COLOR_INDEX.BLUE,
    "TOC_HEADING_4": WD_COLOR_INDEX.DARK_BLUE,
    "HEADING_1":     WD_COLOR_INDEX.BRIGHT_GREEN,
    "HEADING_2":     WD_COLOR_INDEX.TURQUOISE,
    "HEADING_3":     WD_COLOR_INDEX.PINK,
    "HEADING_4":     WD_COLOR_INDEX.VIOLET,
    "HEADING_5":     WD_COLOR_INDEX.TEAL,
    "HEADING_6":     WD_COLOR_INDEX.DARK_BLUE,
    "PARAGRAPH":     WD_COLOR_INDEX.WHITE,
    "LIST_ITEM":     WD_COLOR_INDEX.RED,
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


def _ft_get_style_base(para) -> str:
    try:
        name = para.style.name if para.style else ""
        return name.split(" Char")[0].strip()
    except Exception:
        return ""


def _ft_is_toc_entry_by_style(para) -> bool:
    return bool(re.match(r"^TOC \d+$", _ft_get_style_base(para)))


def _ft_toc_level_from_style(para) -> int:
    m = re.match(r"^TOC (\d+)$", _ft_get_style_base(para))
    return int(m.group(1)) if m else 0


_FT_TOC_ENTRY_RE = re.compile(
    r"^\s*(\d+(?:\.\d+)*)?\s*(.+?)[\s\.…\-]{2,}(\d{1,4})\s*$", re.UNICODE)
_FT_TOC_TAB_RE = re.compile(r"^.+\t\d{1,4}\s*$")


def _ft_is_toc_entry_by_heuristic(para) -> bool:
    text = (para.text or "").strip()
    if not text:
        return False
    if _FT_TOC_ENTRY_RE.match(text) or _FT_TOC_TAB_RE.match(text):
        return True
    xml = para._element.xml
    return "PAGEREF" in xml or ("w:instr" in xml and "TOC" in xml)


def _ft_toc_level_from_heuristic(text: str) -> int:
    m = re.match(r"^\s*(\d+(?:\.\d+)*)\s+", text.strip())
    return min(len(m.group(1).split(".")), 6) if m else 1


def _ft_detect_toc_section(doc):
    paragraphs = doc.paragraphs
    first_toc_style_idx = last_toc_style_idx = -1
    for i, para in enumerate(paragraphs):
        if _ft_is_toc_entry_by_style(para):
            if first_toc_style_idx == -1:
                first_toc_style_idx = i
            last_toc_style_idx = i
        elif first_toc_style_idx != -1 and (i - last_toc_style_idx) > 5:
            break
    if first_toc_style_idx != -1:
        toc_start = first_toc_style_idx
        for j in range(first_toc_style_idx - 1, max(-1, first_toc_style_idx - 5), -1):
            text = (paragraphs[j].text or "").strip().lower()
            sb = _ft_get_style_base(paragraphs[j])
            if sb == "TOC Heading" or "table of contents" in text or text == "contents":
                toc_start = j
                break
            if text:
                break
        return toc_start, last_toc_style_idx, True

    TOC_TITLE_RE = re.compile(
        r"^\s*(table\s+of\s+contents|contents)\s*$", re.IGNORECASE)
    toc_start = -1
    for i, para in enumerate(paragraphs[:100]):
        text = (para.text or "").strip()
        sb = _ft_get_style_base(para)
        if sb == "TOC Heading" or TOC_TITLE_RE.match(text):
            toc_start = i
            break
        if "TOC" in para._element.xml and "w:fldChar" in para._element.xml:
            toc_start = i
            break
    if toc_start == -1:
        return -1, -1, False

    toc_end = toc_start
    consec = 0
    for i in range(toc_start + 1, min(len(paragraphs), toc_start + 300)):
        text = (paragraphs[i].text or "").strip()
        if not text:
            continue
        if (_ft_is_toc_entry_by_heuristic(paragraphs[i]) or
                _ft_is_toc_entry_by_style(paragraphs[i])):
            toc_end = i
            consec = 0
        else:
            consec += 1
            if consec >= 3:
                break
    return toc_start, toc_end, True


def _ft_detect_cover_page(doc) -> int:
    paragraphs = doc.paragraphs
    for i, para in enumerate(paragraphs[:60]):
        for run in para.runs:
            for br in run._element.xpath(".//w:br"):
                if br.get(qn("w:type"), "") == "page":
                    return i
        if i > 0 and has_page_break_before(para):
            return i - 1
    for i, para in enumerate(paragraphs[:40]):
        sn = para.style.name if para.style else ""
        if sn.startswith("Heading") and i > 0:
            return i - 1
    COVER_KW = re.compile(
        r"\b(prepared\s+(by|for)|author|version|copyright|confidential|restricted|"
        r"january|february|march|april|may|june|july|august|september|october|november|december|\d{4})\b",
        re.IGNORECASE)
    non_empty = [(i, p) for i, p in enumerate(
        paragraphs) if (p.text or "").strip()]
    cover_candidate = -1
    for rank, (i, para) in enumerate(non_empty[:25]):
        text = (para.text or "").strip()
        if (len(text) < 200 or para.alignment == WD_ALIGN_PARAGRAPH.CENTER or
                bool(COVER_KW.search(text)) or
                (para.style.name if para.style else "") in ("Title", "Subtitle")):
            cover_candidate = i
        elif rank > 2:
            break
    return cover_candidate


def _make_theme_resolver(doc):
    _cache = {}
    try:
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        theme_part = doc.part.part_related_by(RT.THEME)
        theme_root = theme_part._element
        A = "http://schemas.openxmlformats.org/drawingml/2006/main"

        def _find(path_tags):
            el = theme_root
            for tag in path_tags:
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

    def resolver(theme_val: str) -> str:
        return _cache.get(theme_val, "")
    return resolver


def _ft_get_para_font(para, theme_resolver) -> str:
    from collections import Counter

    def _font_from_rFonts(rFonts_el) -> str:
        if rFonts_el is None:
            return ""
        for attr in (qn("w:ascii"), qn("w:hAnsi"), qn("w:cs")):
            v = rFonts_el.get(attr, "")
            if v:
                return v
        for attr in (qn("w:asciiTheme"), qn("w:hAnsiTheme"), qn("w:cstheme")):
            v = rFonts_el.get(attr, "")
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
                rFonts = rPr.find(qn("w:rFonts"))
                f = _font_from_rFonts(rFonts)
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
                f = _font_from_rFonts(rPr.find(qn("w:rFonts")))
                if f:
                    return f

        style = para.style
        while style is not None:
            try:
                if style.font and style.font.name:
                    return style.font.name
                style_elem = style.element
                style_pPr = style_elem.find(qn("w:pPr"))
                if style_pPr is not None:
                    style_rPr = style_pPr.find(qn("w:rPr"))
                    if style_rPr is not None:
                        f = _font_from_rFonts(style_rPr.find(qn("w:rFonts")))
                        if f:
                            return f
                style_rPr2 = style_elem.find(qn("w:rPr"))
                if style_rPr2 is not None:
                    f = _font_from_rFonts(style_rPr2.find(qn("w:rFonts")))
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


def ft_analyze_document_structure(docx_path: str) -> dict:
    doc = Document(docx_path)
    footer_ids = get_footer_paragraph_ids(doc)
    header_ids = get_header_paragraph_ids(doc)

    elements = []
    element_counts: dict = {}
    detected_types: set = set()
    sample_texts:   dict = {}
    _font_votes:    dict = {}

    theme_resolver = _make_theme_resolver(doc)
    cover_end = _ft_detect_cover_page(doc)
    toc_start, toc_end, has_toc = _ft_detect_toc_section(doc)
    first_content_idx = next(
        (i for i, p in enumerate(doc.paragraphs) if (p.text or "").strip()), 0)

    for idx, para in enumerate(doc.paragraphs):
        if is_in_footer_or_header(para, footer_ids, header_ids):
            continue
        if not (para.text or "").strip():
            continue

        style_name = para.style.name if para.style is not None else ""
        style_base = style_name.split(" Char")[0].strip()
        ptype = None
        in_toc = has_toc and toc_start <= idx <= toc_end
        in_cover = cover_end >= 0 and idx <= cover_end

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
            except:
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

        item = {"type": ptype, "para_idx": idx,
                "indent": get_paragraph_indentation(para)}
        elements.append(item)
        element_counts[ptype] = element_counts.get(ptype, 0) + 1
        detected_types.add(ptype)
        if ptype not in sample_texts:
            t = (para.text or "").strip()
            if t:
                sample_texts[ptype] = t[:250]
        font_name = _ft_get_para_font(para, theme_resolver)
        if font_name:
            _font_votes.setdefault(ptype, []).append(font_name)

    body_children = list(doc.element.body)
    body_elem_to_pos = {id(child): pos for pos,
                        child in enumerate(body_children)}
    para_body_pos:   dict = {}
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
        on_cover = cover_body_threshold >= 0 and 0 <= tbl_body_pos <= cover_body_threshold
        elements.append(
            {"type": "TABLE", "table_idx": i, "on_cover": on_cover})
        element_counts["TABLE"] = element_counts.get("TABLE", 0) + 1
        detected_types.add("TABLE")
        if "TABLE" not in sample_texts:
            for row in table.rows:
                for cell in row.cells:
                    if cell.text.strip():
                        sample_texts["TABLE"] = cell.text.strip()[:250]
                        break

    def _sort_key(x):
        if x == "TITLE":
            return (0, 0)
        if x == "COVER_PAGE":
            return (1, 0)
        if x == "TOC_TITLE":
            return (2, 0)
        if x.startswith("TOC_HEADING_"):
            try:
                return (3, int(x.split("_")[-1]))
            except:
                return (3, 99)
        if x.startswith("HEADING_"):
            try:
                return (4, int(x.split("_")[-1]))
            except:
                return (4, 99)
        if x == "LIST_ITEM":
            return (5, 0)
        if x == "PARAGRAPH":
            return (6, 0)
        return (7, 0)

    from collections import Counter
    current_fonts = {
        ptype: Counter(votes).most_common(1)[0][0]
        for ptype, votes in _font_votes.items() if votes
    }

    return {
        "elements":          elements,
        "detected_elements": sorted(detected_types, key=_sort_key),
        "element_counts":    element_counts,
        "sample_texts":      sample_texts,
        "current_fonts":     current_fonts,
    }


def _ft_get_config_for_type(ptype: str, config: dict):
    font_name = config.get(ptype.lower() + "_font")
    font_size_raw = config.get(ptype.lower() + "_size")
    if not font_name:
        if ptype == "COVER_PAGE":
            font_name = config.get(
                "title_font") or config.get("paragraph_font")
        elif ptype == "LIST_ITEM":
            font_name = config.get(
                "list_item_font") or config.get("paragraph_font")
        else:
            font_name = config.get("paragraph_font")
    if not font_size_raw:
        if ptype == "COVER_PAGE":
            font_size_raw = config.get("title_size") or 12
        elif ptype == "LIST_ITEM":
            font_size_raw = config.get(
                "list_item_size") or config.get("paragraph_size") or 12
        else:
            font_size_raw = config.get("paragraph_size") or 12
    try:
        font_size = int(font_size_raw)
    except:
        font_size = 12
    return font_name, font_size


def _ft_get_all_runs(para):
    all_r = para._element.xpath(".//w:r")
    nested_tbl_runs = set()
    for tbl in para._element.xpath(".//w:tbl"):
        for r in tbl.xpath(".//w:r"):
            nested_tbl_runs.add(id(r))
    return [r for r in all_r if id(r) not in nested_tbl_runs]


def _ft_apply_para_direct_format(para, font_name, font_size_pt, bold,
                                 highlight_color_name=None):
    pPr = para._element.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
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


def _ft_apply_numbering_runprops(para, font_name, font_size, bold=False,
                                 highlight_val=None):
    if not para._element.xpath(".//w:numPr"):
        return
    pPr = para._element.get_or_add_pPr()
    rPr = pPr.find(qn("w:rPr"))
    if rPr is None:
        rPr = OxmlElement("w:rPr")
        pPr.append(rPr)
    # Only touch font if one was actually chosen — never pass None to lxml .set()
    if font_name:
        rFonts = rPr.find(qn("w:rFonts"))
        if rFonts is None:
            rFonts = OxmlElement("w:rFonts")
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
    b_elem = rPr.find(qn("w:b"))
    if bold:
        if b_elem is None:
            b_elem = OxmlElement("w:b")
            rPr.append(b_elem)
    else:
        if b_elem is not None:
            rPr.remove(b_elem)
    if highlight_val:
        hl = rPr.find(qn("w:highlight"))
        if hl is None:
            hl = OxmlElement("w:highlight")
            rPr.append(hl)
        hl.set(qn("w:val"), highlight_val)


def _hex_to_rgb_str(hex_color: str) -> str:
    return hex_color.lstrip("#").upper()


def set_table_heading_bg(table, hex_color: str):
    if not hex_color:
        return
    rgb = _hex_to_rgb_str(hex_color)
    if rgb.upper() == "FFFFFF":
        return
    try:
        first_row = table.rows[0]
    except IndexError:
        return
    for cell in first_row.cells:
        tc = cell._tc
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


def ft_format_docx(input_path: str, elements: list, output_path: str, config: dict):
    doc = Document(input_path)
    footer_ids = get_footer_paragraph_ids(doc)
    header_ids = get_header_paragraph_ids(doc)

    para_map = {e["para_idx"]: e for e in elements if "para_idx" in e}

    for idx, para in enumerate(doc.paragraphs):
        if is_in_footer_or_header(para, footer_ids, header_ids):
            continue
        elem_info = para_map.get(idx)
        if not elem_info:
            continue
        ptype = elem_info["type"]
        original_indent = elem_info.get("indent", 0.0)
        font_name, font_size = _ft_get_config_for_type(ptype, config)

        is_heading = ptype.startswith("HEADING_")
        is_title = ptype == "TITLE"
        is_toc_title = ptype == "TOC_TITLE"
        is_cover = ptype == "COVER_PAGE"
        is_list = ptype == "LIST_ITEM"

        bold_run = (
            (is_heading and config.get("bold_headings", True)) or
            (is_title and config.get("bold_titles",   True)) or
            (is_cover and config.get("bold_titles",   True)) or
            (is_toc_title and config.get("bold_toc",      True)) or
            (is_list and config.get("bold_lists",    False))
        )
        highlight_idx = HIGHLIGHT_COLORS.get(
            ptype) if config.get("highlight") else None
        highlight_name = _COLOR_MAP.get(
            highlight_idx) if highlight_idx else None

        _ft_apply_para_direct_format(
            para, font_name, font_size, bold_run, highlight_name)

        if is_list:
            _ft_apply_numbering_runprops(
                para, font_name, font_size, bold_run, highlight_name)

        all_runs = _ft_get_all_runs(para)
        for r_elem in all_runs:
            rPr = r_elem.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                r_elem.insert(0, rPr)
            if font_name:
                rFonts = rPr.find(qn("w:rFonts"))
                if rFonts is None:
                    rFonts = OxmlElement("w:rFonts")
                    rPr.insert(0, rFonts)
                for attr_name, value in (
                    ("w:ascii", font_name), ("w:hAnsi", font_name),
                        ("w:cs", font_name),    ("w:eastAsia", font_name)):
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
                el.set(qn("w:val"), "1" if bold_run else "0")
            if highlight_name:
                hl = rPr.find(qn("w:highlight"))
                if hl is None:
                    hl = OxmlElement("w:highlight")
                    rPr.append(hl)
                hl.set(qn("w:val"), highlight_name)

        if config.get("preserve_indentation") and original_indent > 0:
            try:
                para.paragraph_format.left_indent = Inches(original_indent)
            except Exception:
                pass

    table_heading_bg = config.get("table_heading_bg")
    for i, table in enumerate(doc.tables):
        tbl_elem = {e.get("table_idx"): e for e in elements if e.get("type") == "TABLE"}
        elem_info = tbl_elem.get(i)
        on_cover = elem_info.get("on_cover", False) if elem_info else False
        if not on_cover and table_heading_bg:
            set_table_heading_bg(table, table_heading_bg)

    doc.save(output_path)


# =============================================================================
#  HTML PREVIEW  (unchanged from original)
# =============================================================================

def docx_to_html_preview(docx_path: str) -> str:
    import html as _html_mod
    doc = Document(docx_path)
    footer_ids = get_footer_paragraph_ids(doc)
    header_ids = get_header_paragraph_ids(doc)

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
        "</style></head><body><div class='page'>"
    )
    parts = [page_css]
    para_map = {id(p._element): p for p in doc.paragraphs}
    table_map = {id(t._tbl): t for t in doc.tables}

    def _render_runs(para) -> str:
        out = []
        for run in para.runs:
            raw = run.text
            if raw is None:
                raw = ""
            t = _html_mod.escape(str(raw))
            rPr = run._r.find(qn("w:rPr"))
            bold = ital = underline = False
            color = ""
            if rPr is not None:
                bold = rPr.find(qn("w:b")) is not None
                ital = rPr.find(qn("w:i")) is not None
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
            if underline:
                t = f"<u>{t}</u>"
            if color:
                t = f'<span style="{color}">{t}</span>'
            out.append(t)
        return "".join(out)

    def _style_tag(style_name: str) -> str:
        sn = style_name.lower()
        for i in range(1, 7):
            if f"heading {i}" in sn:
                return f"h{i}"
        if "title" in sn:
            return "h1"
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
        rows_html = ["<table>"]
        for row_idx, row in enumerate(table.rows):
            rows_html.append("<tr>")
            for cell in row.cells:
                cell_parts = []
                for cpara in cell.paragraphs:
                    if is_in_footer_or_header(cpara, footer_ids, header_ids):
                        continue
                    rendered = _render_runs(cpara)
                    if rendered is not None:
                        cell_parts.append(rendered)
                cell_content = "<br>".join(
                    p for p in cell_parts if p) or "&nbsp;"
                tag_name = "th" if row_idx == 0 else "td"
                cell_style = ""
                tcPr = cell._tc.find(qn("w:tcPr"))
                if tcPr is not None:
                    shd = tcPr.find(qn("w:shd"))
                    if shd is not None:
                        fill = shd.get(qn("w:fill"), "") or ""
                        if fill and fill.upper() not in ("", "AUTO", "FFFFFF"):
                            cell_style = f' style="background:#{fill};color:white;"'
                rows_html.append(
                    f"<{tag_name}{cell_style}>{cell_content}</{tag_name}>")
            rows_html.append("</tr>")
        rows_html.append("</table>")
        return "".join(rows_html)

    in_list = False
    in_numbered = False

    def close_list():
        nonlocal in_list, in_numbered
        if in_list:
            parts.append("</ol>" if in_numbered else "</ul>")
            in_list = False
            in_numbered = False

    for child in doc.element.body:
        tag = child.tag
        if tag == qn("w:p"):
            para = para_map.get(id(child))
            if para is None or is_in_footer_or_header(para, footer_ids, header_ids):
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
            style_name = (para.style.name if para.style else None) or ""
            is_para_list = bool(para._element.xpath(".//w:numPr"))
            fmt = _resolve_numFmt(para) if is_para_list else ""
            is_num = fmt in _NUMBERED_FMTS if fmt else False
            htag = _style_tag(style_name)
            block_css = _para_css(para)
            style_attr = f' style="{block_css}"' if block_css else ""
            inner = _render_runs(para)
            if is_para_list:
                if not in_list:
                    parts.append("<ol>" if is_num else "<ul>")
                    in_list = True
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
#  FLASK ROUTES  — session-aware, no file upload
# =============================================================================

@font_bp.route("/")
def home():
    return render_template("fontUI.html")


@font_bp.route("/analyse", methods=["POST"])
def ft_analyse():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file uploaded yet. Please upload from the home page."}), 400
    try:
        data = ft_analyze_document_structure(path)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@font_bp.route("/original", methods=["POST"])
def ft_original():
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
        html_content = docx_to_html_preview(path)
        return html_content, 200, {"Content-Type": "text/html; charset=utf-8"}
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@font_bp.route("/preview", methods=["POST"])
def ft_preview():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file uploaded yet. Please upload from the home page."}), 400

    try:
        cfg = json.loads(request.form.get("config", "{}"))
        out_docx = os.path.join(
            _session_out_dir(), f"ft_preview_{uuid.uuid4().hex}.docx")
        data = ft_analyze_document_structure(path)
        ft_format_docx(path, data["elements"], out_docx, cfg)

        if IS_WINDOWS and WORD_AVAILABLE:
            try:
                pythoncom.CoInitialize()
                pdf_path = os.path.join(
                    _session_out_dir(), f"ft_preview_{uuid.uuid4().hex}.pdf")
                convert(out_docx, pdf_path)
                pythoncom.CoUninitialize()
                return send_file(pdf_path, mimetype="application/pdf")
            except Exception:
                try:
                    pythoncom.CoUninitialize()
                except Exception:
                    pass

        html_content = docx_to_html_preview(out_docx)
        return html_content, 200, {"Content-Type": "text/html; charset=utf-8"}
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@font_bp.route("/format", methods=["POST"])
def ft_format():
    path = _get_path()
    if not path or not os.path.exists(path):
        return jsonify({"error": "No file uploaded yet. Please upload from the home page."}), 400

    try:
        cfg = json.loads(request.form.get("config", "{}"))
        out_name = f"font_formatted_{uuid.uuid4().hex}.docx"
        out_path = os.path.join(_session_out_dir(), out_name)
        data = ft_analyze_document_structure(path)
        ft_format_docx(path, data["elements"], out_path, cfg)

        # Update working file so the next tool in the pipeline uses this output.
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

    print("Font Formatter running → http://127.0.0.1:5001/font/")
    app.run(debug=True, port=5001, threaded=False)
