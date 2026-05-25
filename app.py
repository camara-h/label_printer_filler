import copy
import io
import json
import re
import uuid
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
try:
    import streamlit as st
except ModuleNotFoundError:
    st = None
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor

APP_DIR = Path(__file__).parent
PREFERRED_TEMPLATE = APP_DIR / "CryoSTUCK_labels.docx"
FALLBACK_TEMPLATE = APP_DIR / "Letter-125-NO0424.docx"
DEFAULT_TEMPLATE = PREFERRED_TEMPLATE if PREFERRED_TEMPLATE.exists() else FALLBACK_TEMPLATE
ROWS_PER_SHEET = 20
LABELS_PER_ROW_GROUP = 5
TOTAL_LABELS_PER_SHEET = ROWS_PER_SHEET * LABELS_PER_ROW_GROUP

ALIGNMENTS = {
    "Left": WD_ALIGN_PARAGRAPH.LEFT,
    "Center": WD_ALIGN_PARAGRAPH.CENTER,
    "Right": WD_ALIGN_PARAGRAPH.RIGHT,
}

DISPLAY_ALIGNMENTS = list(ALIGNMENTS.keys())

MIN_FONT_SIZE = 4.0
MAX_FONT_SIZE = 7.0
MIN_RECOMMENDED_FONT_SIZE = 5.0
MAX_CIRCLE_LINES = 3
MAX_RECTANGLE_LINES = 6
RECOMMENDED_RECTANGLE_LINES = 5
LABEL_LINE_SPACING_MULTIPLE = 0.8
LABEL_LINE_SPACING_TWIPS = int(240 * LABEL_LINE_SPACING_MULTIPLE)


def label_to_table_columns(label_column: int) -> Tuple[int, int]:
    if not 1 <= int(label_column) <= LABELS_PER_ROW_GROUP:
        raise ValueError("Label column must be between 1 and 5.")
    circle_col = (int(label_column) - 1) * 3
    rectangle_col = circle_col + 1
    return circle_col, rectangle_col


def serialize_text(text: str, offset: int, enabled: bool) -> str:
    if not enabled:
        return text
    match = re.search(r"(\d+)(?!.*\d)", text or "")
    if not match:
        return text or ""
    number = match.group(1)
    value = int(number) + offset
    return (text or "")[: match.start()] + str(value).zfill(len(number)) + (text or "")[match.end() :]


def clear_cell(cell):
    for paragraph in list(cell.paragraphs):
        p = paragraph._element
        p.getparent().remove(p)
    for tbl in list(cell.tables):
        t = tbl._element
        t.getparent().remove(t)


def set_cell_padding(cell, top="0", start="0", bottom="0", end="0"):
    tc = cell._tc
    tc_pr = tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for m, v in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = tc_mar.find(qn(f"w:{m}"))
        if node is None:
            node = OxmlElement(f"w:{m}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(v))
        node.set(qn("w:type"), "dxa")


def set_line_spacing(paragraph):
    p_pr = paragraph._p.get_or_add_pPr()
    spacing = p_pr.find(qn("w:spacing"))
    if spacing is None:
        spacing = OxmlElement("w:spacing")
        p_pr.append(spacing)
    spacing.set(qn("w:before"), "0")
    spacing.set(qn("w:after"), "0")
    spacing.set(qn("w:line"), str(LABEL_LINE_SPACING_TWIPS))
    spacing.set(qn("w:lineRule"), "auto")


def add_tab_stop_right(paragraph, position_twips: int = 1200):
    p_pr = paragraph._p.get_or_add_pPr()
    tabs = p_pr.find(qn("w:tabs"))
    if tabs is None:
        tabs = OxmlElement("w:tabs")
        p_pr.append(tabs)
    tab = OxmlElement("w:tab")
    tab.set(qn("w:val"), "right")
    tab.set(qn("w:pos"), str(position_twips))
    tabs.append(tab)


def normalize_hex_color(value: Any, default: str = "#000000") -> str:
    text = str(value or default).strip()
    if not text.startswith("#"):
        text = "#" + text
    if re.fullmatch(r"#[0-9A-Fa-f]{6}", text):
        return text.upper()
    return default


def add_formatted_run(paragraph, text: str, font_size: float, bold: bool, color: str = "#000000"):
    run = paragraph.add_run(text or "")
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    run.font.size = Pt(float(font_size))
    run.bold = bool(bold)
    color = normalize_hex_color(color)
    run.font.color.rgb = RGBColor.from_string(color.replace("#", ""))
    return run


def write_cell_from_lines(cell, lines: List[Dict[str, Any]], label_offset: int = 0, override_left_texts: Optional[List[str]] = None, override_right_texts: Optional[List[str]] = None):
    clear_cell(cell)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_padding(cell, top="0", start="20", bottom="0", end="20")

    if not lines:
        cell.add_paragraph("")
        return

    for idx, line in enumerate(lines):
        paragraph = cell.add_paragraph()
        paragraph.alignment = ALIGNMENTS.get(line.get("align", "Center"), WD_ALIGN_PARAGRAPH.CENTER)
        set_line_spacing(paragraph)

        if override_left_texts is None:
            left_text = serialize_text(line.get("left_text", ""), label_offset, line.get("serialize_left", False))
        else:
            left_text = override_left_texts[idx] if idx < len(override_left_texts) else ""

        if override_right_texts is None:
            right_text = serialize_text(line.get("right_text", ""), label_offset, line.get("serialize_right", False))
        else:
            right_text = override_right_texts[idx] if idx < len(override_right_texts) else ""

        add_formatted_run(paragraph, left_text, line.get("font_size", 6.0), line.get("bold", False), line.get("color", "#000000"))
        if line.get("use_tab") and right_text:
            add_tab_stop_right(paragraph, int(line.get("tab_pos", 1200)))
            paragraph.add_run("\t")
            add_formatted_run(paragraph, right_text, line.get("font_size", 6.0), line.get("bold", False), line.get("color", "#000000"))


def cell_has_content(cell) -> bool:
    return bool(cell.text.strip())


def validate_template(doc: Document) -> List[str]:
    errors = []
    if len(doc.tables) < 1:
        errors.append("No table found in template.")
        return errors
    first_table = doc.tables[0]
    if len(first_table.rows) != ROWS_PER_SHEET:
        errors.append(f"Expected 20 rows in the first table, found {len(first_table.rows)}.")
    if len(first_table.columns) != 14:
        errors.append(f"Expected 14 columns in the first table, found {len(first_table.columns)}.")
    for i, table in enumerate(doc.tables, start=1):
        if len(table.rows) != ROWS_PER_SHEET or len(table.columns) != 14:
            errors.append(f"Table {i} does not match the expected 20 x 14 template structure.")
    return errors


def get_existing_occupied_positions(template_bytes: bytes) -> set:
    doc = Document(io.BytesIO(template_bytes))
    occupied = set()
    for sheet_idx, table in enumerate(doc.tables, start=1):
        if len(table.rows) != ROWS_PER_SHEET or len(table.columns) != 14:
            continue
        for row_idx in range(ROWS_PER_SHEET):
            for label_col in range(1, LABELS_PER_ROW_GROUP + 1):
                circle_col, rectangle_col = label_to_table_columns(label_col)
                if cell_has_content(table.cell(row_idx, circle_col)) or cell_has_content(table.cell(row_idx, rectangle_col)):
                    occupied.add((sheet_idx, row_idx + 1, label_col))
    return occupied


def make_blank_table_copy(table):
    new_tbl = copy.deepcopy(table._tbl)
    # Wrap the copied XML in a temporary document so python-docx can expose cells cleanly.
    # The copied XML is returned after clearing through a real table proxy in the target document.
    return new_tbl


def ensure_sheet_count(doc: Document, desired_sheets: int):
    if desired_sheets <= len(doc.tables):
        return
    if not doc.tables:
        raise ValueError("No table available to duplicate for additional pages.")

    source_table_xml = copy.deepcopy(doc.tables[0]._tbl)
    body = doc._body._element

    def append_before_section_properties(element):
        """Append body elements before w:sectPr so Word does not repair the DOCX."""
        sect_pr = body.find(qn("w:sectPr"))
        if sect_pr is not None:
            body.insert(body.index(sect_pr), element)
        else:
            body.append(element)

    while len(doc.tables) < desired_sheets:
        paragraph = OxmlElement("w:p")
        run = OxmlElement("w:r")
        br = OxmlElement("w:br")
        br.set(qn("w:type"), "page")
        run.append(br)
        paragraph.append(run)
        append_before_section_properties(paragraph)
        append_before_section_properties(copy.deepcopy(source_table_xml))

    # Newly duplicated pages must be blank, even if the source template was partially filled.
    for table in doc.tables[1:]:
        if len(table.rows) == ROWS_PER_SHEET and len(table.columns) == 14:
            for row_idx in range(ROWS_PER_SHEET):
                for col_idx in range(14):
                    clear_cell(table.cell(row_idx, col_idx))

def default_lines(kind: str) -> List[Dict[str, Any]]:
    if kind == "circle":
        return [
            {"left_text": "Tissue 1", "right_text": "", "use_tab": False, "font_size": 7.0, "bold": True, "align": "Center", "serialize_left": True, "serialize_right": False, "tab_pos": 700, "color": "#000000"},
            {"left_text": "EXP_ID", "right_text": "", "use_tab": False, "font_size": 5.0, "bold": False, "align": "Center", "serialize_left": False, "serialize_right": False, "tab_pos": 700, "color": "#000000"},
        ]
    return [
        {"left_text": "Tissue 1", "right_text": "", "use_tab": False, "font_size": 7.0, "bold": True, "align": "Center", "serialize_left": True, "serialize_right": False, "tab_pos": 1200, "color": "#000000"},
        {"left_text": "Tissue Biopsy", "right_text": "", "use_tab": False, "font_size": 6.0, "bold": False, "align": "Center", "serialize_left": False, "serialize_right": False, "tab_pos": 1200, "color": "#000000"},
        {"left_text": "EXP_ID", "right_text": "Exp_data", "use_tab": True, "font_size": 6.0, "bold": False, "align": "Right", "serialize_left": False, "serialize_right": False, "tab_pos": 1200, "color": "#000000"},
    ]


def new_label_set(name="Tissue", start_row=1, start_col=1, count=20) -> Dict[str, Any]:
    circle = default_lines("circle")
    rectangle = default_lines("rectangle")
    circle[0]["left_text"] = f"{name} 1"
    rectangle[0]["left_text"] = f"{name} 1"
    rectangle[1]["left_text"] = f"{name} Biopsy"
    return {
        "name": name,
        "start_sheet": 1,
        "start_row": start_row,
        "start_col": start_col,
        "count": count,
        "circle_lines": circle,
        "rectangle_lines": rectangle,
    }


def init_state():
    if "label_sets" not in st.session_state:
        st.session_state.label_sets = [new_label_set("Tissue", 1, 1, 20)]
    if "layout_df" not in st.session_state:
        st.session_state.layout_df = pd.DataFrame()
    if "generated_docx" not in st.session_state:
        st.session_state.generated_docx = None
    if "generated_inventory_xlsx" not in st.session_state:
        st.session_state.generated_inventory_xlsx = None


def normalize_lines(lines: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    clean = []
    for line in lines:
        clean.append({
            "uid": str(line.get("uid") or uuid.uuid4().hex),
            "left_text": str(line.get("left_text", "")),
            "right_text": str(line.get("right_text", "")),
            "use_tab": bool(line.get("use_tab", False)),
            "font_size": float(line.get("font_size", 6.0)),
            "bold": bool(line.get("bold", False)),
            "align": line.get("align", "Center") if line.get("align", "Center") in DISPLAY_ALIGNMENTS else "Center",
            "serialize_left": bool(line.get("serialize_left", False)),
            "serialize_right": bool(line.get("serialize_right", False)),
            "tab_pos": int(line.get("tab_pos", 1200)),
            "color": normalize_hex_color(line.get("color", "#000000")),
        })
    return clean


def line_widget_key(prefix: str, line: Dict[str, Any], field: str) -> str:
    """Return a widget key tied to a stable line UID instead of a row index.

    Index-based widget keys make Streamlit keep values attached to the visual
    position. After a move, those old position-based values can overwrite the
    reordered list, making the ↑/↓ buttons look like they did nothing.
    """
    uid = str(line.get("uid") or uuid.uuid4().hex)
    line["uid"] = uid
    return f"{prefix}_{uid}_{field}"


def sync_line_widget_state(prefix: str, lines: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    """Copy visible widget values back into their matching logical lines."""
    synced = normalize_lines(lines)
    for line in synced:
        line["left_text"] = st.session_state.get(line_widget_key(prefix, line, "left"), line.get("left_text", ""))
        line["serialize_left"] = st.session_state.get(line_widget_key(prefix, line, "ser_left"), line.get("serialize_left", False))
        line["use_tab"] = st.session_state.get(line_widget_key(prefix, line, "tab"), line.get("use_tab", False))
        line["right_text"] = st.session_state.get(line_widget_key(prefix, line, "right"), line.get("right_text", ""))
        line["serialize_right"] = st.session_state.get(line_widget_key(prefix, line, "ser_right"), line.get("serialize_right", False))
        line["tab_pos"] = st.session_state.get(line_widget_key(prefix, line, "tabpos"), line.get("tab_pos", 1200))
        line["font_size"] = st.session_state.get(line_widget_key(prefix, line, "size"), line.get("font_size", 6.0))
        line["bold"] = st.session_state.get(line_widget_key(prefix, line, "bold"), line.get("bold", False))
        line["align"] = st.session_state.get(line_widget_key(prefix, line, "align"), line.get("align", "Center"))
        line["color"] = st.session_state.get(line_widget_key(prefix, line, "color"), line.get("color", "#000000"))
    return normalize_lines(synced)


def save_and_rerun_line_editor(prefix: str, state_key: str, lines: List[Dict[str, Any]]) -> None:
    st.session_state[state_key] = normalize_lines(lines)
    st.rerun()

def line_defaults() -> Dict[str, Any]:
    return {
        "uid": uuid.uuid4().hex,
        "left_text": "",
        "right_text": "",
        "use_tab": False,
        "font_size": 6.0,
        "bold": False,
        "align": "Center",
        "serialize_left": False,
        "serialize_right": False,
        "tab_pos": 1000,
        "color": "#000000",
    }


def line_count_guidance(label: str, max_lines: Optional[int], recommended_lines: Optional[int]) -> None:
    if label.lower() == "circle":
        st.caption(
            "Circle layout: up to 3 lines. Suggested 3-line format: font 5 plain for experiment info, "
            "font 6 or 7 bold for the main label, and font 5 plain or bold for secondary information."
        )
    else:
        st.caption(
            "Rectangle layout: up to 6 lines. Labels usually print best with 5 lines or fewer. "
            "Suggested format: line 1 font 7 bold for main info, lines 2 to 3 font 6 plain for complementary info, "
            "and line 4 font 5 plain for Exp ID, date, or other reproducibility metadata."
        )

    st.caption(
        "Font guidance: 7 bold is ideal for main information. 6 plain is good for concentrations, storage method, "
        "solvent, or expiration date. 5 is good for Exp ID or date. 4 should only be used if space is very limited."
    )


def font_help_text(label: str) -> str:
    if label.lower() == "circle":
        fit_text = "Approximate circle fit: 7 pt works best for 6 to 8 characters, 6 pt for 8 to 10, 5 pt for 10 to 12, and 4 pt only for very short metadata."
    else:
        fit_text = "Approximate rectangle fit per line: 7 pt works best for 12 to 16 characters, 6 pt for 16 to 22, 5 pt for 22 to 28, and 4 pt only when space is very limited."
    return (
        "Allowed range: 4 to 7 pt. Recommended: 5 to 7 pt. "
        "7 bold is ideal for main information. 6 plain is good for important additional information. "
        "5 is good for reproducibility metadata. 4 should only be used if space is very limited. "
        + fit_text
    )


def line_editor(prefix: str, label: str, lines: List[Dict[str, Any]], max_lines: Optional[int] = None, recommended_lines: Optional[int] = None) -> List[Dict[str, Any]]:
    state_key = f"{prefix}_lines_state"

    if state_key not in st.session_state:
        st.session_state[state_key] = normalize_lines(lines)

    # Keep the canonical line list in sync with any visible widget edits before
    # handling add/remove/reorder buttons. This avoids stale values after Streamlit reruns.
    lines = sync_line_widget_state(prefix, normalize_lines(st.session_state[state_key]))
    st.session_state[state_key] = lines

    line_count_guidance(label, max_lines, recommended_lines)

    if max_lines is not None and len(lines) >= max_lines:
        if label.lower() == "circle":
            st.warning(f"Circle labels are limited to {max_lines} lines because additional lines usually do not print correctly.")
        else:
            st.warning(f"Maximum reached: rectangle labels are limited to {max_lines} lines.")
    elif recommended_lines is not None and len(lines) >= recommended_lines:
        st.warning(f"Rectangle labels usually print best with {recommended_lines} lines or fewer. You can use up to {max_lines} lines if needed.")

    if max_lines is not None and len(lines) > max_lines:
        st.warning(
            f"{label} has {len(lines)} lines, which is above the recommended app limit of {max_lines}. "
            "Remove extra lines one at a time. The app will not delete them automatically."
        )

    add_disabled = max_lines is not None and len(lines) >= max_lines
    add_clicked = st.button(f"Add line to {label.lower()}", key=f"add_{prefix}", disabled=add_disabled)
    if add_disabled:
        st.caption(f"Maximum reached: {label.lower()} labels are limited to {max_lines} lines.")
    if add_clicked:
        lines = sync_line_widget_state(prefix, normalize_lines(st.session_state[state_key]))
        if max_lines is None or len(lines) < max_lines:
            lines.append(line_defaults())
            save_and_rerun_line_editor(prefix, state_key, lines)

    for idx, line in enumerate(lines):
        line = normalize_lines([line])[0]
        line_uid = line["uid"]
        with st.expander(f"{label} line {idx + 1}", expanded=True):
            action_cols = st.columns([1, 1, 1, 5])
            with action_cols[0]:
                if st.button("↑", key=f"{prefix}_move_up_{line_uid}", disabled=idx == 0, help="Move this line up"):
                    lines = sync_line_widget_state(prefix, normalize_lines(st.session_state[state_key]))
                    lines[idx - 1], lines[idx] = lines[idx], lines[idx - 1]
                    save_and_rerun_line_editor(prefix, state_key, lines)
            with action_cols[1]:
                if st.button("↓", key=f"{prefix}_move_down_{line_uid}", disabled=idx >= len(lines) - 1, help="Move this line down"):
                    lines = sync_line_widget_state(prefix, normalize_lines(st.session_state[state_key]))
                    lines[idx + 1], lines[idx] = lines[idx], lines[idx + 1]
                    save_and_rerun_line_editor(prefix, state_key, lines)
            with action_cols[2]:
                if st.button("🗑️", key=f"{prefix}_delete_{line_uid}", help="Delete this line"):
                    lines = sync_line_widget_state(prefix, normalize_lines(st.session_state[state_key]))
                    if 0 <= idx < len(lines):
                        del lines[idx]
                    save_and_rerun_line_editor(prefix, state_key, lines)
            with action_cols[3]:
                st.caption("Reorder or delete this line")

            line["left_text"] = st.text_input("Text", value=line.get("left_text", ""), key=line_widget_key(prefix, line, "left"))
            line["serialize_left"] = st.checkbox("Serialize trailing number in this text", value=line.get("serialize_left", False), key=line_widget_key(prefix, line, "ser_left"))
            line["use_tab"] = st.checkbox("Add tab and right text on this same line", value=line.get("use_tab", False), key=line_widget_key(prefix, line, "tab"))
            if line["use_tab"]:
                line["right_text"] = st.text_input("Right text after tab", value=line.get("right_text", ""), key=line_widget_key(prefix, line, "right"))
                line["serialize_right"] = st.checkbox("Serialize trailing number in right text", value=line.get("serialize_right", False), key=line_widget_key(prefix, line, "ser_right"))
                line["tab_pos"] = st.number_input("Right tab position, twips", min_value=300, max_value=2200, value=int(line.get("tab_pos", 1200)), step=50, key=line_widget_key(prefix, line, "tabpos"))
            else:
                line["right_text"] = line.get("right_text", "")
                line["serialize_right"] = line.get("serialize_right", False)

            c1, c2, c3, c4 = st.columns(4)
            with c1:
                current_size = float(line.get("font_size", 6.0))
                current_size = min(MAX_FONT_SIZE, max(MIN_FONT_SIZE, current_size))
                line["font_size"] = st.number_input(
                    "Font size",
                    min_value=MIN_FONT_SIZE,
                    max_value=MAX_FONT_SIZE,
                    value=current_size,
                    step=0.5,
                    help=font_help_text(label),
                    key=line_widget_key(prefix, line, "size"),
                )
            with c2:
                line["bold"] = st.checkbox("Bold", value=bool(line.get("bold", False)), key=line_widget_key(prefix, line, "bold"))
            with c3:
                line["align"] = st.selectbox("Alignment", options=DISPLAY_ALIGNMENTS, index=DISPLAY_ALIGNMENTS.index(line.get("align", "Center")), key=line_widget_key(prefix, line, "align"))
            with c4:
                line["color"] = st.color_picker("Text color", value=normalize_hex_color(line.get("color", "#000000")), key=line_widget_key(prefix, line, "color"))

            lines[idx] = normalize_lines([line])[0]
    lines = normalize_lines(lines)
    st.session_state[state_key] = lines
    return lines

def line_texts_for_label(lines: List[Dict[str, Any]], offset: int) -> Tuple[List[str], List[str], str]:
    lefts = []
    rights = []
    display = []
    for line in lines:
        left = serialize_text(line.get("left_text", ""), offset, line.get("serialize_left", False))
        right = serialize_text(line.get("right_text", ""), offset, line.get("serialize_right", False))
        lefts.append(left)
        rights.append(right)
        display.append(left + ((" | " + right) if line.get("use_tab") and right else ""))
    return lefts, rights, " / ".join(display)


def next_position(sheet: int, row: int, col: int) -> Tuple[int, int, int]:
    row += 1
    if row > ROWS_PER_SHEET:
        row = 1
        col += 1
    if col > LABELS_PER_ROW_GROUP:
        col = 1
        sheet += 1
    return sheet, row, col


def first_available_position(label_sets: List[Dict[str, Any]], occupied: set, skip_occupied: bool) -> Tuple[int, int, int]:
    """Find the earliest open printed-label position after existing planned sets.

    Scan order is top to bottom within a printed label column, then left to
    right across label columns, then the next sheet. Example: if set 1 starts
    at R1 C1 and has 15 labels, the next set starts at R16 C1; if it has 23
    labels, the next set starts at R4 C2.
    """
    planned = set()
    if label_sets:
        try:
            current_layout = build_layout(label_sets, occupied, skip_occupied)
            for _, r in current_layout[current_layout.get("Use", True)].iterrows():
                planned.add((int(r["Sheet"]), int(r["Row"]), int(r["Label column"])))
        except Exception:
            for label_set in label_sets:
                sheet = int(label_set.get("start_sheet", 1))
                row = int(label_set.get("start_row", 1))
                col = int(label_set.get("start_col", 1))
                for _ in range(int(label_set.get("count", 1))):
                    planned.add((sheet, row, col))
                    sheet, row, col = next_position(sheet, row, col)

    sheet, row, col = 1, 1, 1
    guard = 0
    while guard < 50000:
        candidate = (sheet, row, col)
        if candidate not in planned and not (skip_occupied and candidate in occupied):
            return candidate
        sheet, row, col = next_position(sheet, row, col)
        guard += 1
    raise ValueError("Could not find an available starting position for the new label set.")


def build_layout(label_sets: List[Dict[str, Any]], occupied: set, skip_occupied: bool) -> pd.DataFrame:
    rows = []
    planned = set()
    global_label_num = 1
    for set_idx, label_set in enumerate(label_sets):
        sheet = int(label_set.get("start_sheet", 1))
        row = int(label_set.get("start_row", 1))
        col = int(label_set.get("start_col", 1))
        count = int(label_set.get("count", 1))
        written_for_set = 0
        guard = 0
        while written_for_set < count:
            guard += 1
            if guard > count + 10000:
                raise ValueError("Could not find enough available labels. Please check blocked spaces and starting position.")
            candidate = (sheet, row, col)
            unavailable = candidate in planned or (skip_occupied and candidate in occupied)
            if not unavailable:
                circle_lefts, circle_rights, circle_display = line_texts_for_label(label_set["circle_lines"], written_for_set)
                rect_lefts, rect_rights, rect_display = line_texts_for_label(label_set["rectangle_lines"], written_for_set)
                rows.append({
                    "Use": True,
                    "Global #": global_label_num,
                    "Set #": set_idx + 1,
                    "Set name": label_set.get("name", f"Set {set_idx + 1}"),
                    "Within set #": written_for_set + 1,
                    "Sheet": sheet,
                    "Row": row,
                    "Label column": col,
                    "Circle text": circle_display,
                    "Rectangle text": rect_display,
                    "Circle left JSON": json.dumps(circle_lefts, ensure_ascii=False),
                    "Circle right JSON": json.dumps(circle_rights, ensure_ascii=False),
                    "Rectangle left JSON": json.dumps(rect_lefts, ensure_ascii=False),
                    "Rectangle right JSON": json.dumps(rect_rights, ensure_ascii=False),
                })
                planned.add(candidate)
                global_label_num += 1
                written_for_set += 1
            sheet, row, col = next_position(sheet, row, col)
    return pd.DataFrame(rows)


def layout_warnings(layout_df: pd.DataFrame, occupied: set) -> List[str]:
    warnings = []
    if layout_df.empty:
        return ["No labels were generated."]
    bad = layout_df[(layout_df["Sheet"] < 1) | (layout_df["Row"] < 1) | (layout_df["Row"] > 20) | (layout_df["Label column"] < 1) | (layout_df["Label column"] > 5)]
    if not bad.empty:
        warnings.append("Some edited positions are outside the valid range. Sheet must be at least 1, row must be 1 to 20, and label column must be 1 to 5.")
    pos_counts = layout_df[layout_df.get("Use", True)].groupby(["Sheet", "Row", "Label column"]).size().reset_index(name="n")
    duplicated = pos_counts[pos_counts["n"] > 1]
    if not duplicated.empty:
        preview = ", ".join([f"sheet {int(r.Sheet)}, row {int(r.Row)}, column {int(r['Label column'])}" for _, r in duplicated.head(10).iterrows()])
        warnings.append(f"Duplicate target positions in the editable layout: {preview}.")
    hits = []
    for _, row in layout_df[layout_df.get("Use", True)].iterrows():
        candidate = (int(row["Sheet"]), int(row["Row"]), int(row["Label column"]))
        if candidate in occupied:
            hits.append(candidate)
    if hits:
        preview = ", ".join([f"sheet {s}, row {r}, column {c}" for s, r, c in hits[:10]])
        warnings.append(f"Some target labels already contain text in the uploaded template: {preview}.")
    return warnings


def parse_json_list(value: Any) -> List[str]:
    if isinstance(value, list):
        return [str(x) for x in value]
    try:
        parsed = json.loads(value)
        if isinstance(parsed, list):
            return [str(x) for x in parsed]
    except Exception:
        pass
    return []


def flatten_label_text(left_values: List[str], right_values: Optional[List[str]] = None) -> str:
    """Combine line-level label text into one inventory-table cell.

    Word labels can contain separate lines and tabbed right-side text. For the
    inventory export, these are flattened into a semicolon-separated string so
    the result is easy to sort, filter, and paste into freezer inventory sheets.
    """
    right_values = right_values or []
    pieces = []
    max_len = max(len(left_values), len(right_values))
    for idx in range(max_len):
        left = str(left_values[idx]) if idx < len(left_values) else ""
        right = str(right_values[idx]) if idx < len(right_values) else ""
        combined = "\t".join([part for part in [left, right] if part.strip()])
        combined = re.sub(r"[\r\n\t]+", "; ", combined)
        combined = re.sub(r"\s*;\s*", "; ", combined).strip(" ;")
        if combined:
            pieces.append(combined)
    return "; ".join(pieces)


def box_position(index_zero_based: int) -> Tuple[int, int, str]:
    """Return 10 x 10 box coordinates filled top-to-bottom, then left-to-right.

    Inventory/freezer boxes are commonly filled down one column first:
    1A, 1B, 1C ... 1J, then 2A, 2B ... 10J. The exported
    box_row column is numeric (1-10), while grid_id uses the familiar
    row-letter convention.
    """
    box_col = ((index_zero_based // 10) % 10) + 1
    box_row = (index_zero_based % 10) + 1
    row_letter = chr(ord("A") + box_row - 1)
    grid_id = f"{box_col}{row_letter}"
    return box_col, box_row, grid_id


def build_inventory_table(layout_df: pd.DataFrame, include_box_layout: bool = True) -> pd.DataFrame:
    if layout_df.empty:
        return pd.DataFrame(columns=["sample_id", "description"])

    active_df = layout_df[layout_df.get("Use", True)].copy()

    # Inventory export should follow the generated/sample order, not the physical
    # printed-template order. Sorting by Sheet/Row/Label column makes the export
    # look like Tissue 1, Tissue 21, Tissue 41 because the label sheet has five
    # vertical label columns. Global # preserves the intended serialization order,
    # including manual edits made in the editable layout step.
    if "Global #" in active_df.columns:
        active_df["_inventory_order"] = pd.to_numeric(active_df["Global #"], errors="coerce")
        active_df = active_df.sort_values(["_inventory_order"], kind="stable")
    else:
        active_df = active_df.reset_index(drop=True)

    rows = []
    for inventory_idx, (_, layout_row) in enumerate(active_df.iterrows()):
        circle_lefts = parse_json_list(layout_row.get("Circle left JSON", "[]"))
        circle_rights = parse_json_list(layout_row.get("Circle right JSON", "[]"))
        rect_lefts = parse_json_list(layout_row.get("Rectangle left JSON", "[]"))
        rect_rights = parse_json_list(layout_row.get("Rectangle right JSON", "[]"))

        entry = {
            "sample_id": flatten_label_text(circle_lefts, circle_rights),
            "description": flatten_label_text(rect_lefts, rect_rights),
        }
        if include_box_layout:
            box_col, box_row, grid_id = box_position(inventory_idx)
            entry.update({"box_column": box_col, "box_row": box_row, "grid_id": grid_id})
        rows.append(entry)
    return pd.DataFrame(rows)


def inventory_table_to_excel_bytes(inventory_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        inventory_df.to_excel(writer, index=False, sheet_name="Inventory")
        worksheet = writer.sheets["Inventory"]
        worksheet.freeze_panes = "A2"
        for column_cells in worksheet.columns:
            header = str(column_cells[0].value or "")
            max_len = len(header)
            for cell in column_cells[1:]:
                max_len = max(max_len, len(str(cell.value or "")))
            worksheet.column_dimensions[column_cells[0].column_letter].width = min(max(max_len + 2, 12), 60)
        worksheet.auto_filter.ref = worksheet.dimensions
    return output.getvalue()


def fill_from_layout(template_bytes: bytes, label_sets: List[Dict[str, Any]], layout_df: pd.DataFrame, allow_overwrite: bool) -> bytes:
    doc = Document(io.BytesIO(template_bytes))
    errors = validate_template(doc)
    if errors:
        raise ValueError("Template validation failed: " + " ".join(errors))

    active_df = layout_df[layout_df.get("Use", True)].copy()
    if active_df.empty:
        raise ValueError("No active labels to write.")

    warnings = layout_warnings(active_df, get_existing_occupied_positions(template_bytes))
    blocking = [w for w in warnings if "outside the valid range" in w or "Duplicate target positions" in w]
    if blocking:
        raise ValueError(" ".join(blocking))
    if not allow_overwrite:
        occupied_warnings = [w for w in warnings if "already contain text" in w]
        if occupied_warnings:
            raise ValueError(occupied_warnings[0] + " Enable overwrite to continue.")

    max_sheet = int(active_df["Sheet"].max())
    ensure_sheet_count(doc, max_sheet)

    for _, layout_row in active_df.iterrows():
        set_idx = int(layout_row["Set #"]) - 1
        if set_idx < 0 or set_idx >= len(label_sets):
            raise ValueError("Editable layout refers to a label set that no longer exists. Rebuild the layout preview.")
        label_set = label_sets[set_idx]
        sheet = int(layout_row["Sheet"])
        row_num = int(layout_row["Row"])
        label_col = int(layout_row["Label column"])
        if sheet < 1 or row_num < 1 or row_num > 20 or label_col < 1 or label_col > 5:
            raise ValueError("All edited positions must be inside the valid sheet, row, and label column ranges.")

        table = doc.tables[sheet - 1]
        circle_col, rectangle_col = label_to_table_columns(label_col)

        circle_lefts = parse_json_list(layout_row.get("Circle left JSON", "[]"))
        circle_rights = parse_json_list(layout_row.get("Circle right JSON", "[]"))
        rect_lefts = parse_json_list(layout_row.get("Rectangle left JSON", "[]"))
        rect_rights = parse_json_list(layout_row.get("Rectangle right JSON", "[]"))

        write_cell_from_lines(table.cell(row_num - 1, circle_col), label_set["circle_lines"], override_left_texts=circle_lefts, override_right_texts=circle_rights)
        write_cell_from_lines(table.cell(row_num - 1, rectangle_col), label_set["rectangle_lines"], override_left_texts=rect_lefts, override_right_texts=rect_rights)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def layout_grid_html(layout_df: pd.DataFrame, occupied: set) -> str:
    active = layout_df[layout_df.get("Use", True)].copy() if not layout_df.empty else pd.DataFrame()
    if active.empty:
        return "<p>No layout generated yet.</p>"
    max_sheet = max(1, int(active["Sheet"].max()))
    html_parts = ["<style>.sheetgrid{border-collapse:collapse;margin-bottom:24px}.sheetgrid td,.sheetgrid th{border:1px solid #ddd;padding:4px;text-align:center;font-size:12px}.sheetgrid td{width:110px;height:34px}.used{background:#f7f7f7}.occupied{background:#ffe5e5}.planned{background:#e7f3ff}.conflict{background:#ffd3a8}.small{font-size:11px;color:#555}</style>"]
    pos_to_text = {}
    duplicates = set()
    for _, r in active.iterrows():
        key = (int(r["Sheet"]), int(r["Row"]), int(r["Label column"]))
        if key in pos_to_text:
            duplicates.add(key)
            pos_to_text[key] += f"<br>⚠ {r['Set name']} #{int(r['Within set #'])}"
        else:
            pos_to_text[key] = f"{r['Set name']} #{int(r['Within set #'])}"
    for sheet in range(1, max_sheet + 1):
        html_parts.append(f"<h4>Sheet {sheet}</h4><table class='sheetgrid'><tr><th>Row</th>" + "".join([f"<th>Label col {c}</th>" for c in range(1, 6)]) + "</tr>")
        for row in range(1, 21):
            html_parts.append(f"<tr><th>{row}</th>")
            for col in range(1, 6):
                key = (sheet, row, col)
                cls = "used"
                text = ""
                if key in occupied:
                    cls = "occupied"
                    text = "Existing text"
                if key in pos_to_text:
                    cls = "conflict" if key in duplicates or key in occupied else "planned"
                    text = pos_to_text[key]
                html_parts.append(f"<td class='{cls}'>{text}</td>")
            html_parts.append("</tr>")
        html_parts.append("</table>")
    return "".join(html_parts)


def inject_custom_css():
    st.markdown(
        """
        <style>
        div[data-baseweb="tab-list"] {
            gap: 0.45rem;
            margin-top: 0.25rem;
            margin-bottom: 0.45rem;
        }
        div[data-baseweb="tab-list"] button[role="tab"] {
            background-color: #eeeeee;
            border: 1px solid #b8b8b8;
            border-radius: 0.55rem 0.55rem 0 0;
            padding: 0.55rem 0.9rem;
        }
        div[data-baseweb="tab-list"] button[role="tab"] p {
            color: #8A1538;
            font-weight: 800;
            font-size: 1.02rem;
        }
        div[data-baseweb="tab-list"] button[role="tab"][aria-selected="true"] {
            background-color: #4a4a4a;
            border-color: #4a4a4a;
        }
        div[data-baseweb="tab-list"] button[role="tab"][aria-selected="true"] p {
            color: #ffffff;
        }
        div[data-baseweb="tab-list"] button[role="tab"]:hover {
            background-color: #dcdcdc;
            border-color: #8A1538;
        }
        div[data-baseweb="tab-list"] button[role="tab"][aria-selected="true"]:hover {
            background-color: #3f3f3f;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def main():
    st.set_page_config(page_title="LabTAG LCS-125WH Label Filler", layout="wide")
    init_state()
    inject_custom_css()

    st.title("LabTAG LCS-125WH Label Filler")
    st.caption("Uses the Word template as the source of truth and only writes formatted text into label cells.")

    with st.sidebar:
        st.header("Template")
        uploaded_template = st.file_uploader("Upload official or partially used .docx template", type=["docx"])
        use_default = st.checkbox("Use included LCS-125WH template", value=True)
        if uploaded_template is not None:
            template_bytes = uploaded_template.read()
            st.success("Using uploaded template.")
        elif use_default and DEFAULT_TEMPLATE.exists():
            template_bytes = DEFAULT_TEMPLATE.read_bytes()
            st.success("Using included template.")
        else:
            st.error("Upload a .docx template or keep the included template selected.")
            st.stop()

        try:
            template_doc = Document(io.BytesIO(template_bytes))
            template_errors = validate_template(template_doc)
            if template_errors:
                for error in template_errors:
                    st.error(error)
            existing_occupied = get_existing_occupied_positions(template_bytes)
            st.caption(f"Detected {len(existing_occupied)} occupied label positions in the current template.")
        except Exception as exc:
            st.error(f"Could not read template: {exc}")
            st.stop()

        st.header("Output behavior")
        skip_occupied = st.checkbox("Skip labels that already contain text", value=True)
        allow_overwrite = st.checkbox("Allow overwrite if edited layout targets used labels", value=False)
        include_box_layout = st.checkbox("Add 10 x 10 box columns to inventory export", value=True)
        st.caption("If more labels are requested than fit on the existing page, the app adds another blank copy of the template page.")

    st.subheader("1. Build label ID sets")
    top_cols = st.columns([1, 1, 2])
    with top_cols[0]:
        if st.button("Add another Label ID Set", type="secondary"):
            previous = copy.deepcopy(st.session_state.label_sets[-1])
            previous["name"] = f"Set {len(st.session_state.label_sets) + 1}"
            next_sheet, next_row, next_col = first_available_position(st.session_state.label_sets, existing_occupied, skip_occupied)
            previous["start_sheet"] = next_sheet
            previous["start_row"] = next_row
            previous["start_col"] = next_col
            st.session_state.label_sets.append(previous)
            st.rerun()
    with top_cols[1]:
        if len(st.session_state.label_sets) > 1 and st.button("Remove last set"):
            st.session_state.label_sets.pop()
            st.rerun()

    for set_idx, label_set in enumerate(st.session_state.label_sets):
        with st.expander(f"Label ID Set {set_idx + 1}: {label_set.get('name', '')}", expanded=True):
            c1, c2, c3, c4, c5 = st.columns(5)
            with c1:
                label_set["name"] = st.text_input("Set name", value=label_set.get("name", f"Set {set_idx + 1}"), key=f"set_name_{set_idx}")
            with c2:
                label_set["start_sheet"] = st.number_input("Start sheet", min_value=1, max_value=50, value=int(label_set.get("start_sheet", 1)), step=1, key=f"set_sheet_{set_idx}")
            with c3:
                label_set["start_row"] = st.number_input("Start row", min_value=1, max_value=20, value=int(label_set.get("start_row", 1)), step=1, key=f"set_row_{set_idx}")
            with c4:
                label_set["start_col"] = st.number_input("Start label column", min_value=1, max_value=5, value=int(label_set.get("start_col", 1)), step=1, key=f"set_col_{set_idx}")
            with c5:
                label_set["count"] = st.number_input("Labels to fill", min_value=1, max_value=1000, value=int(label_set.get("count", 20)), step=1, key=f"set_count_{set_idx}")

            st.caption("Choose which part of the label to edit below.")
            ltab, rtab = st.tabs(["● Circle formatting", "▰ Rectangle formatting"])
            with ltab:
                label_set["circle_lines"] = line_editor(f"set{set_idx}_circle", "Circle", label_set.get("circle_lines", default_lines("circle")), max_lines=MAX_CIRCLE_LINES)
            with rtab:
                label_set["rectangle_lines"] = line_editor(f"set{set_idx}_rectangle", "Rectangle", label_set.get("rectangle_lines", default_lines("rectangle")), max_lines=MAX_RECTANGLE_LINES, recommended_lines=RECOMMENDED_RECTANGLE_LINES)

    st.divider()
    st.subheader("2. Build editable layout and preview")
    cbuild, cclear = st.columns([1, 3])
    with cbuild:
        if st.button("Build editable layout", type="primary"):
            try:
                st.session_state.layout_df = build_layout(st.session_state.label_sets, existing_occupied, skip_occupied)
                st.session_state.generated_docx = None
                st.session_state.generated_inventory_xlsx = None
                st.success("Editable layout generated.")
            except Exception as exc:
                st.error(str(exc))
    with cclear:
        st.caption("This creates the serialized rows first. Then you can manually move labels by editing Sheet, Row, and Label column, or fine tune the generated text JSON fields.")

    if not st.session_state.layout_df.empty:
        tab_preview, tab_grid, tab_advanced = st.tabs(["Editable layout", "Sheet map", "Advanced text editing"])
        with tab_preview:
            display_cols = ["Use", "Global #", "Set name", "Within set #", "Sheet", "Row", "Label column", "Circle text", "Rectangle text"]
            edited_display = st.data_editor(
                st.session_state.layout_df[display_cols],
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                column_config={
                    "Use": st.column_config.CheckboxColumn("Use"),
                    "Sheet": st.column_config.NumberColumn("Sheet", min_value=1, step=1),
                    "Row": st.column_config.NumberColumn("Row", min_value=1, max_value=20, step=1),
                    "Label column": st.column_config.NumberColumn("Label column", min_value=1, max_value=5, step=1),
                },
                disabled=["Global #", "Set name", "Within set #", "Circle text", "Rectangle text"],
                key="layout_display_editor",
            )
            for col in ["Use", "Sheet", "Row", "Label column"]:
                st.session_state.layout_df[col] = edited_display[col]

            warnings = layout_warnings(st.session_state.layout_df, existing_occupied)
            if warnings:
                for warning in warnings:
                    if "already contain text" in warning and allow_overwrite:
                        st.warning(warning + " Overwrite is enabled.")
                    else:
                        st.warning(warning)
            else:
                st.success("No layout conflicts detected.")

        with tab_grid:
            st.markdown(layout_grid_html(st.session_state.layout_df, existing_occupied), unsafe_allow_html=True)
            st.caption("Blue means planned labels. Red means existing text from the uploaded template. Orange means a conflict or duplicate.")

        with tab_advanced:
            st.caption("The display preview is not enough to preserve line-level formatting. Edit these JSON lists only when you need to fine tune the final text after serialization.")
            hidden_cols = ["Global #", "Circle left JSON", "Circle right JSON", "Rectangle left JSON", "Rectangle right JSON"]
            edited_json = st.data_editor(
                st.session_state.layout_df[hidden_cols],
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=["Global #"],
                key="json_layout_editor",
            )
            for col in hidden_cols[1:]:
                st.session_state.layout_df[col] = edited_json[col]

        st.divider()
        st.subheader("3. Generate files")
        if st.button("Generate filled DOCX and inventory table", type="primary"):
            try:
                output_bytes = fill_from_layout(
                    template_bytes=template_bytes,
                    label_sets=copy.deepcopy(st.session_state.label_sets),
                    layout_df=st.session_state.layout_df,
                    allow_overwrite=allow_overwrite,
                )
                inventory_df = build_inventory_table(st.session_state.layout_df, include_box_layout=include_box_layout)
                inventory_bytes = inventory_table_to_excel_bytes(inventory_df)
                st.session_state.generated_docx = output_bytes
                st.session_state.generated_inventory_xlsx = inventory_bytes
                st.success("DOCX and inventory-style table generated.")
            except Exception as exc:
                st.error(str(exc))

        if st.session_state.generated_docx is not None:
            c_docx, c_xlsx = st.columns(2)
            with c_docx:
                st.download_button(
                    label="Download filled template",
                    data=st.session_state.generated_docx,
                    file_name="filled_LCS_125WH_labels.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
            with c_xlsx:
                st.download_button(
                    label="Download Inventory-style Table",
                    data=st.session_state.generated_inventory_xlsx,
                    file_name="inventory_style_table.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
    else:
        st.info("Add one or more label sets, then click Build editable layout.")


if __name__ == "__main__":
    main()
