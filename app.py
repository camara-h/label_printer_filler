import copy
import io
import re
from pathlib import Path
from typing import Any, Dict, List, Tuple

try:
    import streamlit as st
except ModuleNotFoundError:
    st = None
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

APP_DIR = Path(__file__).parent
DEFAULT_TEMPLATE = APP_DIR / "Letter-125-NO0424.docx"
ROWS_PER_SHEET = 20
LABELS_PER_ROW_GROUP = 5
TOTAL_LABELS = 100

ALIGNMENTS = {
    "Left": WD_ALIGN_PARAGRAPH.LEFT,
    "Center": WD_ALIGN_PARAGRAPH.CENTER,
    "Right": WD_ALIGN_PARAGRAPH.RIGHT,
}


def label_to_table_columns(label_column: int) -> Tuple[int, int]:
    """Return zero-based table columns for a logical label column, 1 through 5."""
    if not 1 <= label_column <= 5:
        raise ValueError("Label column must be between 1 and 5")
    circle_col = (label_column - 1) * 3
    rectangle_col = circle_col + 1
    return circle_col, rectangle_col


def iter_label_positions(start_row: int, start_label_column: int, count: int, manual_jumps=None):
    """Yield zero-based row, logical label column for top-to-bottom then left-to-right filling.

    manual_jumps is a dict where the key is a one-based label number already written,
    and the value is the next one-based row and label column.
    """
    manual_jumps = manual_jumps or {}
    row = start_row - 1
    col = start_label_column
    for label_number in range(1, count + 1):
        if col > LABELS_PER_ROW_GROUP:
            raise ValueError("Not enough labels remain on this sheet for the requested count.")
        if not 0 <= row < ROWS_PER_SHEET:
            raise ValueError("Row cursor moved outside the 20 available rows.")
        yield row, col
        if label_number in manual_jumps:
            next_row, next_col = manual_jumps[label_number]
            row = int(next_row) - 1
            col = int(next_col)
        else:
            row += 1
            if row >= ROWS_PER_SHEET:
                row = 0
                col += 1


def serialize_text(text: str, offset: int, enabled: bool) -> str:
    """Increment the final integer in text when enabled. Preserves leading zeros."""
    if not enabled:
        return text
    match = re.search(r"(\d+)(?!.*\d)", text)
    if not match:
        return text
    number = match.group(1)
    value = int(number) + offset
    return text[: match.start()] + str(value).zfill(len(number)) + text[match.end() :]


def clear_cell(cell):
    for paragraph in cell.paragraphs:
        p = paragraph._element
        p.getparent().remove(p)
    for tbl in cell.tables:
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
    spacing.set(qn("w:line"), "120")
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


def add_formatted_run(paragraph, text: str, font_size: float, bold: bool):
    run = paragraph.add_run(text)
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    run.font.size = Pt(float(font_size))
    run.bold = bool(bold)
    return run


def write_cell(cell, lines: List[Dict[str, Any]], label_offset: int):
    clear_cell(cell)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_padding(cell, top="0", start="20", bottom="0", end="20")

    if not lines:
        cell.add_paragraph("")
        return

    for line in lines:
        paragraph = cell.add_paragraph()
        paragraph.alignment = ALIGNMENTS[line["align"]]
        set_line_spacing(paragraph)

        left_text = serialize_text(line["left_text"], label_offset, line.get("serialize_left", False))
        right_text = serialize_text(line.get("right_text", ""), label_offset, line.get("serialize_right", False))

        add_formatted_run(paragraph, left_text, line["font_size"], line["bold"])
        if line.get("use_tab") and right_text:
            add_tab_stop_right(paragraph, int(line.get("tab_pos", 1200)))
            paragraph.add_run("\t")
            add_formatted_run(paragraph, right_text, line["font_size"], line["bold"])


def cell_has_content(cell) -> bool:
    return bool(cell.text.strip())


def validate_template(doc: Document) -> List[str]:
    errors = []
    if len(doc.tables) < 1:
        errors.append("No table found in template.")
        return errors
    table = doc.tables[0]
    if len(table.rows) != 20:
        errors.append(f"Expected 20 rows, found {len(table.rows)}.")
    if len(table.columns) != 14:
        errors.append(f"Expected 14 columns, found {len(table.columns)}.")
    return errors


def fill_template(
    template_bytes: bytes,
    circle_lines: List[Dict[str, Any]],
    rectangle_lines: List[Dict[str, Any]],
    start_row: int,
    start_label_column: int,
    count: int,
    allow_overwrite: bool,
    manual_jumps=None,
) -> bytes:
    doc = Document(io.BytesIO(template_bytes))
    errors = validate_template(doc)
    if errors:
        raise ValueError("Template validation failed: " + " ".join(errors))

    table = doc.tables[0]
    positions = list(iter_label_positions(start_row, start_label_column, count, manual_jumps=manual_jumps))

    occupied = []
    for row_idx, label_col in positions:
        circle_col, rectangle_col = label_to_table_columns(label_col)
        circle_cell = table.cell(row_idx, circle_col)
        rectangle_cell = table.cell(row_idx, rectangle_col)
        if cell_has_content(circle_cell) or cell_has_content(rectangle_cell):
            occupied.append((row_idx + 1, label_col))

    if occupied and not allow_overwrite:
        preview = ", ".join([f"row {r}, label column {c}" for r, c in occupied[:10]])
        more = "" if len(occupied) <= 10 else f" and {len(occupied) - 10} more"
        raise ValueError(f"Some target labels already contain text: {preview}{more}. Enable overwrite to continue.")

    for label_offset, (row_idx, label_col) in enumerate(positions):
        circle_col, rectangle_col = label_to_table_columns(label_col)
        write_cell(table.cell(row_idx, circle_col), circle_lines, label_offset)
        write_cell(table.cell(row_idx, rectangle_col), rectangle_lines, label_offset)

    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def default_lines(kind: str) -> List[Dict[str, Any]]:
    if kind == "circle":
        return [
            {"left_text": "Tissue 1", "right_text": "", "use_tab": False, "font_size": 7.0, "bold": True, "align": "Center", "serialize_left": True, "serialize_right": False, "tab_pos": 700},
            {"left_text": "EXP_ID", "right_text": "", "use_tab": False, "font_size": 5.0, "bold": False, "align": "Center", "serialize_left": False, "serialize_right": False, "tab_pos": 700},
        ]
    return [
        {"left_text": "Tissue 1", "right_text": "", "use_tab": False, "font_size": 7.0, "bold": True, "align": "Center", "serialize_left": True, "serialize_right": False, "tab_pos": 1200},
        {"left_text": "Tissue Biopsy", "right_text": "", "use_tab": False, "font_size": 6.0, "bold": False, "align": "Center", "serialize_left": False, "serialize_right": False, "tab_pos": 1200},
        {"left_text": "EXP_ID", "right_text": "Exp_data", "use_tab": True, "font_size": 6.0, "bold": False, "align": "Right", "serialize_left": False, "serialize_right": False, "tab_pos": 1200},
    ]


def init_state():
    if "circle_lines" not in st.session_state:
        st.session_state.circle_lines = default_lines("circle")
    if "rectangle_lines" not in st.session_state:
        st.session_state.rectangle_lines = default_lines("rectangle")


def line_editor(prefix: str, label: str, key_name: str):
    st.subheader(label)
    lines = st.session_state[key_name]

    col_a, col_b = st.columns(2)
    with col_a:
        if st.button(f"Add line to {label.lower()}", key=f"add_{prefix}"):
            lines.append({"left_text": "", "right_text": "", "use_tab": False, "font_size": 6.0, "bold": False, "align": "Center", "serialize_left": False, "serialize_right": False, "tab_pos": 1000})
            st.rerun()
    with col_b:
        if len(lines) > 0 and st.button(f"Remove last {label.lower()} line", key=f"remove_{prefix}"):
            lines.pop()
            st.rerun()

    for idx, line in enumerate(lines):
        with st.expander(f"{label} line {idx + 1}", expanded=True):
            line["left_text"] = st.text_input("Text", value=line.get("left_text", ""), key=f"{prefix}_left_{idx}")
            line["serialize_left"] = st.checkbox("Serialize trailing number in this text", value=line.get("serialize_left", False), key=f"{prefix}_ser_left_{idx}")

            line["use_tab"] = st.checkbox("Add tab and right text on this same line", value=line.get("use_tab", False), key=f"{prefix}_tab_{idx}")
            if line["use_tab"]:
                line["right_text"] = st.text_input("Right text after tab", value=line.get("right_text", ""), key=f"{prefix}_right_{idx}")
                line["serialize_right"] = st.checkbox("Serialize trailing number in right text", value=line.get("serialize_right", False), key=f"{prefix}_ser_right_{idx}")
                line["tab_pos"] = st.number_input("Right tab position, twips", min_value=300, max_value=2000, value=int(line.get("tab_pos", 1200)), step=50, key=f"{prefix}_tabpos_{idx}")

            c1, c2, c3 = st.columns(3)
            with c1:
                line["font_size"] = st.number_input("Font size", min_value=3.0, max_value=14.0, value=float(line.get("font_size", 6.0)), step=0.5, key=f"{prefix}_size_{idx}")
            with c2:
                line["bold"] = st.checkbox("Bold", value=bool(line.get("bold", False)), key=f"{prefix}_bold_{idx}")
            with c3:
                line["align"] = st.selectbox("Alignment", options=list(ALIGNMENTS.keys()), index=list(ALIGNMENTS.keys()).index(line.get("align", "Center")), key=f"{prefix}_align_{idx}")


def main():
    st.set_page_config(page_title="LabTAG LCS-125WH Label Filler", layout="wide")
    init_state()

    st.title("LabTAG LCS-125WH Label Filler")
    st.caption("Fills the official Word template without rebuilding the table geometry.")

    with st.sidebar:
        st.header("Template")
        uploaded_template = st.file_uploader("Upload official .docx template", type=["docx"])
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

        st.header("Placement")
        start_row = st.number_input("Start row", min_value=1, max_value=20, value=1, step=1)
        start_col = st.number_input("Start label column", min_value=1, max_value=5, value=1, step=1)
        max_remaining = (5 - int(start_col)) * 20 + (21 - int(start_row))
        count = st.number_input("Number of labels to fill", min_value=1, max_value=max_remaining, value=min(20, max_remaining), step=1)
        allow_overwrite = st.checkbox("Allow overwrite if target labels already contain text", value=False)

        st.header("Manual jumps")
        st.caption("Optional. After a specific label number is filled, move the next label to a new row and label column.")
        jump_count = st.number_input("Number of manual jumps", min_value=0, max_value=20, value=0, step=1)
        manual_jumps = {}
        for j in range(int(jump_count)):
            st.markdown(f"Jump {j + 1}")
            jc1, jc2, jc3 = st.columns(3)
            with jc1:
                after_label = st.number_input("After label #", min_value=1, max_value=int(count), value=min(int(count), j + 1), key=f"jump_after_{j}")
            with jc2:
                jump_row = st.number_input("Next row", min_value=1, max_value=20, value=1, key=f"jump_row_{j}")
            with jc3:
                jump_col = st.number_input("Next label column", min_value=1, max_value=5, value=1, key=f"jump_col_{j}")
            manual_jumps[int(after_label)] = (int(jump_row), int(jump_col))

        st.header("Presets")
        if st.button("Load tissue example preset"):
            st.session_state.circle_lines = default_lines("circle")
            st.session_state.rectangle_lines = default_lines("rectangle")
            st.rerun()

    left, right = st.columns(2)
    with left:
        line_editor("circle", "Circle", "circle_lines")
    with right:
        line_editor("rectangle", "Rectangle", "rectangle_lines")

    st.divider()
    st.subheader("Preview of serialization")
    preview_rows = []
    for i in range(min(5, int(count))):
        circle_preview = " / ".join(
            [serialize_text(line["left_text"], i, line.get("serialize_left", False)) for line in st.session_state.circle_lines]
        )
        rectangle_preview = " / ".join(
            [
                serialize_text(line["left_text"], i, line.get("serialize_left", False))
                + ((" | " + serialize_text(line.get("right_text", ""), i, line.get("serialize_right", False))) if line.get("use_tab") else "")
                for line in st.session_state.rectangle_lines
            ]
        )
        preview_rows.append({"Label #": i + 1, "Circle": circle_preview, "Rectangle": rectangle_preview})
    st.table(preview_rows)

    if st.button("Generate filled DOCX", type="primary"):
        try:
            output_bytes = fill_template(
                template_bytes=template_bytes,
                circle_lines=copy.deepcopy(st.session_state.circle_lines),
                rectangle_lines=copy.deepcopy(st.session_state.rectangle_lines),
                start_row=int(start_row),
                start_label_column=int(start_col),
                count=int(count),
                allow_overwrite=allow_overwrite,
                manual_jumps=manual_jumps,
            )
            st.success("DOCX generated.")
            st.download_button(
                label="Download filled template",
                data=output_bytes,
                file_name="filled_LCS_125WH_labels.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
        except Exception as exc:
            st.error(str(exc))


if __name__ == "__main__":
    main()
