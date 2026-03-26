#font1.py
import os
import re
import json
import uuid
import platform

from flask import Blueprint, Flask, request, jsonify, send_file, render_template
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
UPLOAD = "uploads/font"
OUTPUT = "outputs/font"
os.makedirs(UPLOAD, exist_ok=True)
os.makedirs(OUTPUT, exist_ok=True)


# =============================================================================
#  SHARED UTILITIES
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


def is_numbered_list(para) -> bool:
    try:
        style_name = para.style.name if para.style else ""
        if "List Number" in style_name:
            return True
        numPr_list = para._element.xpath(".//w:numPr")
        if not numPr_list:
            return False
        numId_nodes = para._element.xpath(".//w:numPr/w:numId")
        if not numId_nodes:
            return False
        try:
            numbering_part = para.part.numbering_part
            if numbering_part is None:
                return False
            numId_val = numId_nodes[0].get(qn("w:val"), "0")
            ilvl_nodes = para._element.xpath(".//w:numPr/w:ilvl")
            ilvl_val = ilvl_nodes[0].get(
                qn("w:val"), "0") if ilvl_nodes else "0"
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            num_elements = numbering_part._element.xpath(
                f".//w:num[@w:numId='{numId_val}']", namespaces=ns)
            if not num_elements:
                return False
            abstractNumId_nodes = num_elements[0].xpath(
                ".//w:abstractNumId", namespaces=ns)
            if not abstractNumId_nodes:
                return False
            abstract_id = abstractNumId_nodes[0].get(qn("w:val"), "0")
            abstract_num = numbering_part._element.xpath(
                f".//w:abstractNum[@w:abstractNumId='{abstract_id}']", namespaces=ns)
            if not abstract_num:
                return False
            lvl_elements = abstract_num[0].xpath(
                f".//w:lvl[@w:ilvl='{ilvl_val}']/w:numFmt", namespaces=ns)
            if lvl_elements:
                fmt = lvl_elements[0].get(qn("w:val"), "")
                if fmt in ("decimal", "lowerLetter", "upperLetter", "lowerRoman",
                           "upperRoman", "ordinal", "cardinalText", "decimalZero"):
                    return True
        except Exception:
            pass
    except Exception:
        pass
    return False


def _is_bullet_list(para) -> bool:
    try:
        style_name = para.style.name if para.style else ""
        if "List Bullet" in style_name:
            return True
        numPr_list = para._element.xpath(".//w:numPr")
        if not numPr_list:
            return False
        numId_nodes = para._element.xpath(".//w:numPr/w:numId")
        if not numId_nodes:
            return False
        try:
            numbering_part = para.part.numbering_part
            if numbering_part is None:
                return True
            numId_val = numId_nodes[0].get(qn("w:val"), "0")
            ilvl_nodes = para._element.xpath(".//w:numPr/w:ilvl")
            ilvl_val = ilvl_nodes[0].get(
                qn("w:val"), "0") if ilvl_nodes else "0"
            ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
            num_elements = numbering_part._element.xpath(
                f".//w:num[@w:numId='{numId_val}']", namespaces=ns)
            if not num_elements:
                return False
            abstractNumId_nodes = num_elements[0].xpath(
                ".//w:abstractNumId", namespaces=ns)
            if not abstractNumId_nodes:
                return False
            abstract_id = abstractNumId_nodes[0].get(qn("w:val"), "0")
            abstract_num = numbering_part._element.xpath(
                f".//w:abstractNum[@w:abstractNumId='{abstract_id}']", namespaces=ns)
            if not abstract_num:
                return False
            lvl_elements = abstract_num[0].xpath(
                f".//w:lvl[@w:ilvl='{ilvl_val}']/w:numFmt", namespaces=ns)
            if lvl_elements:
                fmt = lvl_elements[0].get(qn("w:val"), "")
                if fmt in ("bullet", "none"):
                    return True
                if fmt in ("decimal", "lowerLetter", "upperLetter", "lowerRoman",
                           "upperRoman", "ordinal", "cardinalText", "decimalZero"):
                    return False
                return True
        except Exception:
            return False
    except Exception:
        pass
    return False


def is_any_list_paragraph(para) -> bool:
    try:
        if para._element.xpath(".//w:numPr"):
            return True
    except Exception:
        pass
    try:
        style_name = para.style.name if para.style is not None else ""
        if "List" in style_name:
            return True
    except Exception:
        pass
    return False


def get_list_type(para) -> str:
    if _is_bullet_list(para):
        return "BULLET_ITEM"
    if is_numbered_list(para):
        return "NUMBERED_ITEM"
    if is_any_list_paragraph(para):
        return "LIST_ITEM"
    return None


# =============================================================================
#  HIGHLIGHT COLOURS
# =============================================================================

HIGHLIGHT_COLORS = {
    "TITLE":         WD_COLOR_INDEX.YELLOW,
    "COVER_PAGE":    WD_COLOR_INDEX.DARK_YELLOW,
    "TOC_TITLE":     WD_COLOR_INDEX.RED,
    "TOC_HEADING_1": WD_COLOR_INDEX.GRAY_25,
    "TOC_HEADING_2": WD_COLOR_INDEX.GRAY_50,
    "TOC_HEADING_3": WD_COLOR_INDEX.BLUE,
    "TOC_HEADING_4": WD_COLOR_INDEX.GREEN,
    "HEADING_1":     WD_COLOR_INDEX.BRIGHT_GREEN,
    "HEADING_2":     WD_COLOR_INDEX.TURQUOISE,
    "HEADING_3":     WD_COLOR_INDEX.PINK,
    "HEADING_4":     WD_COLOR_INDEX.VIOLET,
    "HEADING_5":     WD_COLOR_INDEX.TEAL,
    "HEADING_6":     WD_COLOR_INDEX.DARK_BLUE,
    "PARAGRAPH":     WD_COLOR_INDEX.WHITE,
    "LIST_ITEM":     WD_COLOR_INDEX.DARK_RED,
    "TABLE":         WD_COLOR_INDEX.AUTO,
}

_COLOR_MAP = {
    WD_COLOR_INDEX.YELLOW:       "yellow",
    WD_COLOR_INDEX.BRIGHT_GREEN: "green",
    WD_COLOR_INDEX.TURQUOISE:    "cyan",
    WD_COLOR_INDEX.PINK:         "magenta",
    WD_COLOR_INDEX.VIOLET:       "magenta",
    WD_COLOR_INDEX.TEAL:         "cyan",
    WD_COLOR_INDEX.DARK_BLUE:    "darkBlue",
    WD_COLOR_INDEX.GRAY_25:      "lightGray",
    WD_COLOR_INDEX.GRAY_50:      "darkGray",
    WD_COLOR_INDEX.DARK_YELLOW:  "darkYellow",
}


# =============================================================================
#  TOC / COVER-PAGE DETECTION
# =============================================================================

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


# =============================================================================
#  DOCUMENT STRUCTURE ANALYSIS
# =============================================================================

def ft_analyze_document_structure(docx_path: str) -> dict:
    doc = Document(docx_path)
    footer_ids = get_footer_paragraph_ids(doc)
    header_ids = get_header_paragraph_ids(doc)

    elements = []
    element_counts = {}
    detected_types = set()
    sample_texts = {}

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
                lt = get_list_type(para)
                if lt:
                    ptype = lt
            if not ptype:
                lvl, num, title = detect_numbered_heading(raw_text)
                if lvl:
                    ptype = f"HEADING_{lvl}"
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

    body_children = list(doc.element.body)
    body_elem_to_pos = {id(child): pos for pos,
                        child in enumerate(body_children)}
    para_body_pos = {}
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
            return 0
        if x == "COVER_PAGE":
            return 1
        if x == "TOC_TITLE":
            return 2
        if x.startswith("TOC_HEADING_"):
            return 3
        if x.startswith("HEADING_"):
            return 4
        if x in ("BULLET_ITEM", "NUMBERED_ITEM", "LIST_ITEM"):
            return 5
        if x == "PARAGRAPH":
            return 6
        return 7

    return {
        "elements":          elements,
        "detected_elements": sorted(detected_types, key=_sort_key),
        "element_counts":    element_counts,
        "sample_texts":      sample_texts,
    }


# =============================================================================
#  FONT FORMATTING HELPERS
# =============================================================================

def _ft_get_config_for_type(ptype: str, config: dict):
    font_name = config.get(ptype.lower() + "_font")
    font_size_raw = config.get(ptype.lower() + "_size")
    if not font_name:
        if ptype == "COVER_PAGE":
            font_name = config.get("title_font") or config.get(
                "paragraph_font") or "Arial"
        elif ptype in ("BULLET_ITEM", "NUMBERED_ITEM"):
            font_name = config.get("list_item_font") or config.get(
                "paragraph_font") or "Arial"
        else:
            font_name = config.get("paragraph_font") or "Arial"
    if not font_size_raw:
        if ptype == "COVER_PAGE":
            font_size_raw = config.get("title_size") or 12
        elif ptype in ("BULLET_ITEM", "NUMBERED_ITEM"):
            font_size_raw = config.get(
                "list_item_size") or config.get("paragraph_size") or 12
        else:
            font_size_raw = config.get("paragraph_size") or 12
    try:
        font_size = int(font_size_raw)
    except Exception:
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
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), rgb)


# =============================================================================
#  MAIN FORMATTING FUNCTION
# =============================================================================

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
        is_list = ptype in ("BULLET_ITEM", "NUMBERED_ITEM", "LIST_ITEM")

        bold_run = (
            (is_heading and config.get("bold_headings", True)) or
            (is_title and config.get("bold_titles",   True)) or
            (is_cover and config.get("bold_titles",   True)) or
            (is_toc_title and config.get("bold_toc",      True)) or
            (is_list and config.get("bold_lists",    False))
        )

        highlight_color_name = None
        if config.get("highlight", False):
            hc = HIGHLIGHT_COLORS.get(ptype, WD_COLOR_INDEX.GRAY_25)
            highlight_color_name = _COLOR_MAP.get(hc)

        _ft_apply_para_direct_format(para, font_name, font_size, bold_run,
                                     highlight_color_name)

        if is_heading:
            text = (para.text or "").strip()
            lvl, num, title = detect_numbered_heading(text)
            if num and title:
                clear_paragraph_runs(para)
                para.add_run(f"{num}\t")
                para.add_run(title)

        for r_el in _ft_get_all_runs(para):
            if r_el.xpath(".//w:drawing"):
                continue
            rPr = r_el.find(qn("w:rPr"))
            if rPr is None:
                rPr = OxmlElement("w:rPr")
                r_el.insert(0, rPr)
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
            half_pts = str(int(font_size * 2))
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
                el.set(qn("w:val"), "1" if bold_run else "0")
            if highlight_color_name:
                hl = rPr.find(qn("w:highlight"))
                if hl is None:
                    hl = OxmlElement("w:highlight")
                    rPr.append(hl)
                hl.set(qn("w:val"), highlight_color_name)

        if is_list:
            _ft_apply_numbering_runprops(para, font_name, font_size,
                                         bold_run, highlight_color_name)

        if (config.get("preserve_indentation", True) and
                original_indent and original_indent > 0):
            para.paragraph_format.left_indent = Inches(original_indent)

    for e in elements:
        if e["type"] != "TABLE":
            continue
        tbl_idx = e.get("table_idx", -1)
        if tbl_idx < 0 or tbl_idx >= len(doc.tables):
            continue
        table = doc.tables[tbl_idx]
        fn, fs = _ft_get_config_for_type("TABLE", config)
        half_pts = str(int(fs * 2))
        tbl_heading_bg = config.get("table_heading_bg", "#1e3a5f")
        if not e.get("on_cover", False):
            set_table_heading_bg(table, tbl_heading_bg)
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    if is_in_footer_or_header(para, footer_ids, header_ids):
                        continue
                    _ft_apply_para_direct_format(
                        para, fn, fs, False,
                        "cyan" if config.get("highlight") else None)
                    for r_el in _ft_get_all_runs(para):
                        if r_el.xpath(".//w:drawing"):
                            continue
                        rPr = r_el.find(qn("w:rPr"))
                        if rPr is None:
                            rPr = OxmlElement("w:rPr")
                            r_el.insert(0, rPr)
                        rFonts = rPr.find(qn("w:rFonts"))
                        if rFonts is None:
                            rFonts = OxmlElement("w:rFonts")
                            rPr.insert(0, rFonts)
                        for attr_name, value in (
                                ("w:ascii", fn), ("w:hAnsi", fn),
                                ("w:cs", fn), ("w:eastAsia", fn)):
                            rFonts.set(qn(attr_name), value)
                        for tag in ("w:sz", "w:szCs"):
                            el = rPr.find(qn(tag))
                            if el is None:
                                el = OxmlElement(tag)
                                rPr.append(el)
                            el.set(qn("w:val"), half_pts)
                        if config.get("highlight", False):
                            hl = rPr.find(qn("w:highlight"))
                            if hl is None:
                                hl = OxmlElement("w:highlight")
                                rPr.append(hl)
                            hl.set(qn("w:val"), "cyan")

    doc.save(output_path)


# =============================================================================
#  HTML PREVIEW  (fallback for non-Windows / no Word installed)
# =============================================================================

def _run_font_css(run):
    styles = []
    try:
        if run.font.name:
            styles.append("font-family:'{}',serif".format(run.font.name))
        if run.font.size:
            styles.append("font-size:{}pt".format(run.font.size.pt))
        if run.bold:
            styles.append("font-weight:bold")
        if run.italic:
            styles.append("font-style:italic")
        if run.underline:
            styles.append("text-decoration:underline")
        try:
            if run.font.color and run.font.color.type is not None:
                styles.append("color:#{}".format(run.font.color.rgb))
        except Exception:
            pass
        if run.font.highlight_color:
            _HL_CSS = {
                WD_COLOR_INDEX.YELLOW:       "#fff59d",
                WD_COLOR_INDEX.BRIGHT_GREEN: "#b9f6ca",
                WD_COLOR_INDEX.TURQUOISE:    "#b2ebf2",
                WD_COLOR_INDEX.PINK:         "#f8bbd0",
                WD_COLOR_INDEX.VIOLET:       "#e1bee7",
                WD_COLOR_INDEX.TEAL:         "#b2dfdb",
                WD_COLOR_INDEX.DARK_BLUE:    "#1a237e",
                WD_COLOR_INDEX.GRAY_25:      "#f5f5f5",
                WD_COLOR_INDEX.GRAY_50:      "#e0e0e0",
                WD_COLOR_INDEX.DARK_YELLOW:  "#fff9c4",
            }
            bg = _HL_CSS.get(run.font.highlight_color)
            if bg:
                styles.append("background:{}".format(bg))
    except Exception:
        pass
    return ";".join(styles)


def _para_css(para):
    styles = []
    try:
        _ALIGN = {
            WD_ALIGN_PARAGRAPH.LEFT:    "left",
            WD_ALIGN_PARAGRAPH.CENTER:  "center",
            WD_ALIGN_PARAGRAPH.RIGHT:   "right",
            WD_ALIGN_PARAGRAPH.JUSTIFY: "justify",
        }
        if para.alignment in _ALIGN:
            styles.append("text-align:{}".format(_ALIGN[para.alignment]))
        pf = para.paragraph_format
        if pf.left_indent:
            try:
                styles.append(
                    "margin-left:{}px".format(int(pf.left_indent.inches * 96)))
            except Exception:
                pass
        if pf.space_before:
            try:
                styles.append("margin-top:{}pt".format(pf.space_before.pt))
            except Exception:
                pass
        if pf.space_after:
            try:
                styles.append("margin-bottom:{}pt".format(pf.space_after.pt))
            except Exception:
                pass
    except Exception:
        pass
    return ";".join(styles)


def _style_tag(style_name):
    if not style_name:
        return "p"
    sn = style_name.lower()
    for i in range(1, 7):
        if "heading {}".format(i) in sn:
            return "h{}".format(i)
    if "title" in sn:
        return "h1"
    return "p"


def docx_to_html_preview(docx_path: str) -> str:
    import html as _html
    doc = Document(docx_path)
    footer_ids = get_footer_paragraph_ids(doc)
    header_ids = get_header_paragraph_ids(doc)

    para_map = {id(p._element): p for p in doc.paragraphs}
    table_map = {id(t._tbl):     t for t in doc.tables}

    page_css = (
        '<!DOCTYPE html><html><head><meta charset="utf-8"><style>'
        'body{font-family:Arial,sans-serif;margin:0;padding:0;background:#e0e0e0;}'
        '.page{background:white;width:794px;min-height:1123px;margin:30px auto;'
        'padding:72px 80px;box-shadow:0 2px 12px rgba(0,0,0,0.18);box-sizing:border-box;}'
        'h1,h2,h3,h4,h5,h6{margin:0.4em 0 0.2em;}'
        'p{margin:0.25em 0;line-height:1.4;}'
        'table{border-collapse:collapse;width:100%;margin:10px 0;}'
        'td,th{border:1px solid #ccc;padding:5px 8px;vertical-align:top;}'
        'ul,ol{margin:0.5em 0 0.5em 1.8em;padding:0;}'
        'li{margin:0.2em 0;line-height:1.4;}'
        'hr.page-break{border:none;border-top:2px dashed #ccc;margin:24px 0;}'
        '</style></head><body><div class="page">'
    )
    parts = [page_css]

    def _render_runs(para):
        inner = ""
        if para.runs:
            for run in para.runs:
                run_text = _html.escape(run.text)
                if not run_text:
                    continue
                rcss = _run_font_css(run)
                inner += ('<span style="{}">{}</span>'.format(rcss, run_text)
                          if rcss else run_text)
        else:
            inner = _html.escape(para.text or "")
        return inner

    def _render_table(table):
        rows_html = ["<table>"]
        for row_idx, row in enumerate(table.rows):
            rows_html.append("<tr>")
            seen_cells = set()
            for cell in row.cells:
                cell_key = id(cell._tc)
                if cell_key in seen_cells:
                    rows_html.append("<td>&nbsp;</td>")
                    continue
                seen_cells.add(cell_key)
                cell_parts = []
                for cpara in cell.paragraphs:
                    if is_in_footer_or_header(cpara, footer_ids, header_ids):
                        continue
                    cell_parts.append(_render_runs(cpara))
                cell_content = "<br>".join(
                    p for p in cell_parts if p) or "&nbsp;"
                tag_name = "th" if row_idx == 0 else "td"
                cell_style = ""
                tcPr = cell._tc.find(qn("w:tcPr"))
                if tcPr is not None:
                    shd = tcPr.find(qn("w:shd"))
                    if shd is not None:
                        fill = shd.get(qn("w:fill"), "")
                        if fill and fill.upper() not in ("", "AUTO", "FFFFFF"):
                            cell_style = ' style="background:#{};color:white;"'.format(
                                fill)
                rows_html.append(
                    "<{0}{1}>{2}</{0}>".format(tag_name, cell_style, cell_content))
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
            if not text.strip():
                if not has_pb:
                    close_list()
                    parts.append("<p>&nbsp;</p>")
                continue
            style_name = para.style.name if para.style else ""
            is_para_list = bool(para._element.xpath(".//w:numPr"))
            is_num = is_numbered_list(para)
            htag = _style_tag(style_name)
            block_css = _para_css(para)
            style_attr = ' style="{}"'.format(block_css) if block_css else ""
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
                parts.append("<li{}>{}</li>".format(style_attr, inner))
            else:
                close_list()
                parts.append(
                    "<{0}{1}>{2}</{0}>".format(htag, style_attr, inner))
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
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400
    ext = os.path.splitext(secure_filename(f.filename))[1].lower()
    if ext != ".docx":
        return jsonify({"error": "Only .docx files are supported"}), 400
    path = os.path.join(UPLOAD, f"{uuid.uuid4()}{ext}")
    f.save(path)
    try:
        data = ft_analyze_document_structure(path)
        return jsonify(data)
    except Exception as e:
        return jsonify({"error": str(e)}), 500


@font_bp.route("/preview", methods=["POST"])
def ft_preview():
    """
    Preview logic mirrors alignment.py:
      - On Windows with Word installed → convert docx → PDF → stream PDF
      - Otherwise → convert docx → HTML → stream HTML
    """
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400
    cfg = json.loads(request.form.get("config", "{}"))
    ext = os.path.splitext(secure_filename(f.filename))[1].lower()
    in_path = os.path.join(UPLOAD, f"{uuid.uuid4()}{ext}")
    f.save(in_path)

    out_docx = os.path.join(OUTPUT, f"ft_preview_{uuid.uuid4()}.docx")
    data = ft_analyze_document_structure(in_path)
    ft_format_docx(in_path, data["elements"], out_docx, cfg)

    # ── Windows + Word: return PDF (same as alignment.py) ─────────────────
    if IS_WINDOWS and WORD_AVAILABLE:
        try:
            pythoncom.CoInitialize()
            pdf_path = os.path.join(OUTPUT, f"ft_preview_{uuid.uuid4()}.pdf")
            convert(out_docx, pdf_path)
            pythoncom.CoUninitialize()
            return send_file(pdf_path, mimetype="application/pdf")
        except Exception as e:
            try:
                pythoncom.CoUninitialize()
            except Exception:
                pass
            # fall through to HTML preview on failure

    # ── Fallback: return HTML preview ─────────────────────────────────────
    html_content = docx_to_html_preview(out_docx)
    return html_content, 200, {"Content-Type": "text/html; charset=utf-8"}


@font_bp.route("/format", methods=["POST"])
def ft_format():
    f = request.files.get("file")
    if not f:
        return jsonify({"error": "No file uploaded"}), 400
    cfg = json.loads(request.form.get("config", "{}"))
    ext = os.path.splitext(secure_filename(f.filename))[1].lower()
    in_path = os.path.join(UPLOAD, f"{uuid.uuid4()}{ext}")
    out_name = f"{os.path.splitext(secure_filename(f.filename))[0]}_font_formatted.docx"
    out_path = os.path.join(OUTPUT, f"{uuid.uuid4()}_{out_name}")
    f.save(in_path)
    data = ft_analyze_document_structure(in_path)
    ft_format_docx(in_path, data["elements"], out_path, cfg)
    return send_file(out_path, as_attachment=True, download_name=out_name)


# =============================================================================
#  STANDALONE RUNNER  — python font1.py
# =============================================================================

if __name__ == "__main__":
    app = Flask(__name__, template_folder="templates")
    app.register_blueprint(font_bp)

    @app.route("/")
    def root():
        from flask import redirect
        return redirect("/font/")

    print("Font Formatter running → http://127.0.0.1:5001/font/")
    app.run(debug=True, port=5001, threaded=False)
