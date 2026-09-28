import copy
import io
import re
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pandas as pd
try:
    import qrcode
except ModuleNotFoundError:
    qrcode = None
try:
    import streamlit as st
except ModuleNotFoundError:
    st = None
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, RGBColor, Inches

APP_DIR = Path(__file__).parent
PREFERRED_TEMPLATE = APP_DIR / "CryoSTUCK_labels.docx"
FALLBACK_TEMPLATE = APP_DIR / "Letter-125-NO0424.docx"
DEFAULT_TEMPLATE = PREFERRED_TEMPLATE if PREFERRED_TEMPLATE.exists() else FALLBACK_TEMPLATE

ROWS_PER_SHEET = 20
LABELS_PER_ROW_GROUP = 5
TOTAL_LABELS_PER_SHEET = ROWS_PER_SHEET * LABELS_PER_ROW_GROUP
BOX_ROWS = 10
BOX_COLS = 10
LABEL_LINE_SPACING_MULTIPLE = 0.75
LABEL_LINE_SPACING_TWIPS = int(240 * LABEL_LINE_SPACING_MULTIPLE)
MAX_CIRCLE_LINES = 3
MAX_RECTANGLE_LINES = 6
DEFAULT_QR_SIZE_INCHES = 0.16
DEFAULT_RECTANGLE_QR_SIZE_INCHES = 0.16
DEFAULT_CIRCLE_QR_SMALL_INCHES = 0.15
DEFAULT_CIRCLE_QR_MAIN_INCHES = 0.20
DEFAULT_MAX_QR_PAYLOAD_CHARS = 60

DEFAULT_CHAR_LIMITS = {
    "Circle": {"7": 8, "6": 10, "5": 13, "4": 16},
    "Rectangle": {"7": 18, "6": 24, "5": 30, "4": 38},
}

ALIGNMENTS = {
    "Left": WD_ALIGN_PARAGRAPH.LEFT,
    "Center": WD_ALIGN_PARAGRAPH.CENTER,
    "Right": WD_ALIGN_PARAGRAPH.RIGHT,
}
DISPLAY_ALIGNMENTS = list(ALIGNMENTS.keys())
PRINT_PARTS = ["Circle", "Rectangle", "Ignore"]
SIDES = ["Left/new line", "Right/tab on same line"]
ALL_LABELS_SET_KEY = "__ALL_LABELS__"
UNASSIGNED_SET_KEY = "__UNASSIGNED__"


def label_to_table_columns(label_column: int) -> Tuple[int, int]:
    if not 1 <= int(label_column) <= LABELS_PER_ROW_GROUP:
        raise ValueError("Label column must be between 1 and 5.")
    circle_col = (int(label_column) - 1) * 3
    rectangle_col = circle_col + 1
    return circle_col, rectangle_col


def next_position(sheet: int, row: int, col: int) -> Tuple[int, int, int]:
    row += 1
    if row > ROWS_PER_SHEET:
        row = 1
        col += 1
    if col > LABELS_PER_ROW_GROUP:
        col = 1
        sheet += 1
    return sheet, row, col


def normalize_hex_color(value: Any, default: str = "#000000") -> str:
    text = str(value or default).strip()
    if not text.startswith("#"):
        text = "#" + text
    if re.fullmatch(r"#[0-9A-Fa-f]{6}", text):
        return text.upper()
    return default


def normalize_column_name(value: Any) -> str:
    return str(value or "").strip().lower().replace(" ", "_")


def load_char_limit_config(uploaded_json=None) -> Dict[str, Dict[str, int]]:
    config = copy.deepcopy(DEFAULT_CHAR_LIMITS)
    if uploaded_json is None:
        return config
    try:
        import json
        raw = json.load(uploaded_json)
        for part in ["Circle", "Rectangle"]:
            if isinstance(raw.get(part), dict):
                for size, limit in raw[part].items():
                    try:
                        config[part][str(int(float(size)))] = int(limit)
                    except Exception:
                        pass
    except Exception:
        pass
    return config


def printable_length(text: str) -> int:
    return len(str(text or "").replace("\t", " ").replace("\n", " "))




def line_text_for_qr(line: Dict[str, Any]) -> str:
    """Return one clean text value from a formatted label line for QR payloads."""
    left = str(line.get("left_text", "")).strip()
    right = str(line.get("right_text", "")).strip()
    return "	".join([part for part in [left, right] if part]).strip()


def truncate_qr_payload(parts: List[str], max_chars: int, separator: str = "	") -> str:
    """Join non-empty parts and cap total QR payload length.

    Keeps earlier fields first. UniqueID should be passed as the first part, so it is prioritized.
    """
    max_chars = max(1, int(max_chars or DEFAULT_MAX_QR_PAYLOAD_CHARS))
    cleaned = []
    for part in parts:
        text = re.sub(r"[\r\n\t]+", " ", str(part or "")).strip()
        text = re.sub(r"\s+", " ", text)
        if text:
            cleaned.append(text)
    if not cleaned:
        return ""

    payload = ""
    for part in cleaned:
        candidate = part if not payload else payload + separator + part
        if len(candidate) <= max_chars:
            payload = candidate
            continue
        remaining = max_chars - len(payload) - (len(separator) if payload else 0)
        if remaining > 0:
            payload = (payload + separator if payload else "") + part[:remaining].rstrip()
        break
    return payload[:max_chars].strip()


def build_qr_payload_from_layout_row(layout_row, max_chars: int) -> str:
    """Build a capped QR payload from available label information.

    Priority order:
    1) Unique ID, when present
    2) Circle line 2 main info
    3) Circle line 1
    4) Circle line 3, when present
    5) Rectangle line 1 main info
    6) Remaining rectangle lines, in order

    UniqueID is the safest database identifier and is always placed first when available,
    but it is not required. If no UniqueID exists, descriptive label text is encoded instead.
    """
    unique_id = str(layout_row.get("Unique ID", "")).strip()
    circle_lines = parse_lines_repr(layout_row.get("Circle lines JSON", "[]"))
    rect_lines = parse_lines_repr(layout_row.get("Rectangle lines JSON", "[]"))

    def find_line(lines, line_num):
        for line in lines:
            try:
                if int(line.get("line_num", 0)) == int(line_num):
                    return line_text_for_qr(line)
            except Exception:
                continue
        return ""

    parts = [
        unique_id,
        find_line(circle_lines, 2),
        find_line(circle_lines, 1),
        find_line(circle_lines, 3),
        find_line(rect_lines, 1),
    ]
    for line in sorted(rect_lines, key=lambda x: int(x.get("line_num", 999))):
        try:
            if int(line.get("line_num", 0)) <= 1:
                continue
        except Exception:
            pass
        parts.append(line_text_for_qr(line))
    return truncate_qr_payload(parts, max_chars=max_chars, separator="	")

def make_qr_image_bytes(value: str) -> Optional[io.BytesIO]:
    if not value or qrcode is None:
        return None
    qr = qrcode.QRCode(
        version=None,
        error_correction=qrcode.constants.ERROR_CORRECT_M,
        box_size=4,
        border=1,
    )
    qr.add_data(str(value))
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")
    bio = io.BytesIO()
    img.save(bio, format="PNG")
    bio.seek(0)
    return bio


def clear_cell(cell):
    for paragraph in list(cell.paragraphs):
        p = paragraph._element
        p.getparent().remove(p)
    for tbl in list(cell.tables):
        t = tbl._element
        t.getparent().remove(t)


def set_cell_padding(cell, top="0", start="20", bottom="0", end="20"):
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


def add_formatted_run(paragraph, text: str, font_size: float, bold: bool, color: str = "#000000"):
    run = paragraph.add_run(str(text or ""))
    run.font.name = "Calibri"
    run._element.rPr.rFonts.set(qn("w:eastAsia"), "Calibri")
    run.font.size = Pt(float(font_size))
    run.bold = bool(bold)
    color = normalize_hex_color(color)
    run.font.color.rgb = RGBColor.from_string(color.replace("#", ""))
    return run


def cell_has_content(cell) -> bool:
    return bool(cell.text.strip())


def validate_template(doc: Document) -> List[str]:
    errors = []
    if len(doc.tables) < 1:
        return ["No table found in template."]
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


def ensure_sheet_count(doc: Document, desired_sheets: int):
    if desired_sheets <= len(doc.tables):
        return
    if not doc.tables:
        raise ValueError("No table available to duplicate for additional pages.")

    source_table_xml = copy.deepcopy(doc.tables[0]._tbl)
    body = doc._body._element

    def append_before_section_properties(element):
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

    for table in doc.tables[1:]:
        if len(table.rows) == ROWS_PER_SHEET and len(table.columns) == 14:
            for row_idx in range(ROWS_PER_SHEET):
                for col_idx in range(14):
                    clear_cell(table.cell(row_idx, col_idx))


def clean_cell_value(value: Any) -> str:
    if pd.isna(value):
        return ""
    if hasattr(value, "strftime"):
        # Excel date cells usually arrive as Timestamp/datetime. Use a lab-friendly date format.
        try:
            return value.strftime("%d/%m/%Y")
        except Exception:
            pass
    text = str(value)
    if text.endswith(".0"):
        try:
            as_float = float(text)
            as_int = int(as_float)
            if as_float == as_int:
                return str(as_int)
        except Exception:
            pass
    return text.strip()


def flatten_label_text(lines: List[Dict[str, Any]]) -> str:
    pieces = []
    for line in lines:
        left = str(line.get("left_text", "")).strip()
        right = str(line.get("right_text", "")).strip()
        combined = "\t".join([part for part in [left, right] if part])
        combined = re.sub(r"[\r\n\t]+", "; ", combined)
        combined = re.sub(r"\s*;\s*", "; ", combined).strip(" ;")
        if combined:
            pieces.append(combined)
    return "; ".join(pieces)


def box_position(index_zero_based: int) -> Tuple[int, int, str]:
    box_col = ((index_zero_based // 10) % 10) + 1
    box_row = (index_zero_based % 10) + 1
    row_letter = chr(ord("A") + box_row - 1)
    return box_col, box_row, f"{box_col}{row_letter}"


def _line_has_text(line: Dict[str, Any]) -> bool:
    return bool(str(line.get("left_text", "")).strip() or str(line.get("right_text", "")).strip())


def add_qr_paragraph(cell, qr_text: str, qr_size_inches: float, alignment=WD_ALIGN_PARAGRAPH.CENTER):
    qr_image = make_qr_image_bytes(str(qr_text).strip())
    if qr_image is None:
        return
    paragraph = cell.add_paragraph()
    paragraph.alignment = alignment
    set_line_spacing(paragraph)
    run = paragraph.add_run()
    run.add_picture(qr_image, width=Inches(float(qr_size_inches)))


def write_text_lines(cell, lines: List[Dict[str, Any]], default_alignment: str = "Center"):
    for line in lines:
        if not _line_has_text(line):
            continue
        paragraph = cell.add_paragraph()
        paragraph.alignment = ALIGNMENTS.get(line.get("align", default_alignment), ALIGNMENTS.get(default_alignment, WD_ALIGN_PARAGRAPH.CENTER))
        set_line_spacing(paragraph)
        add_formatted_run(
            paragraph,
            line.get("left_text", ""),
            line.get("font_size", 6.0),
            line.get("bold", False),
            line.get("color", "#000000"),
        )
        if line.get("right_text", ""):
            add_tab_stop_right(paragraph, int(line.get("tab_pos", 1200)))
            paragraph.add_run("\t")
            add_formatted_run(
                paragraph,
                line.get("right_text", ""),
                line.get("font_size", 6.0),
                line.get("bold", False),
                line.get("color", "#000000"),
            )


def write_cell_from_lines(cell, lines: List[Dict[str, Any]], qr_text: str = "", qr_size_inches: float = DEFAULT_RECTANGLE_QR_SIZE_INCHES):
    clear_cell(cell)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_padding(cell, top="0", start="20", bottom="0", end="20")
    if not lines and not qr_text:
        cell.add_paragraph("")
        return
    write_text_lines(cell, lines, default_alignment="Left")
    if qr_text:
        add_qr_paragraph(cell, qr_text, qr_size_inches, alignment=WD_ALIGN_PARAGRAPH.RIGHT)


def write_circle_cell_from_lines(
    cell,
    lines: List[Dict[str, Any]],
    qr_text: str = "",
    add_circle_qr: bool = False,
    small_qr_size_inches: float = DEFAULT_CIRCLE_QR_SMALL_INCHES,
    main_qr_size_inches: float = DEFAULT_CIRCLE_QR_MAIN_INCHES,
):
    clear_cell(cell)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    set_cell_padding(cell, top="0", start="10", bottom="0", end="10")
    text_lines = [line for line in lines if _line_has_text(line)][:MAX_CIRCLE_LINES]

    if add_circle_qr and qr_text:
        # The third circle text line is intentionally reserved for the QR when lid QR is enabled.
        # Keep lines 1-2 only; CircleLine3 remains available when lid QR is disabled.
        qr_text_lines = [
            line for line in text_lines
            if int(line.get("line_num", 0) or 0) <= 2
        ]
        has_second_line = any(
            int(line.get("line_num", idx + 1)) == 2 and _line_has_text(line)
            for idx, line in enumerate(qr_text_lines)
        )
        if has_second_line:
            # With a second text line present, keep the QR small and place it first.
            add_qr_paragraph(cell, qr_text, small_qr_size_inches, alignment=WD_ALIGN_PARAGRAPH.CENTER)
            write_text_lines(cell, qr_text_lines, default_alignment="Center")
        else:
            # If line 2 is blank, use that space for the larger QR.
            write_text_lines(cell, qr_text_lines, default_alignment="Center")
            add_qr_paragraph(cell, qr_text, main_qr_size_inches, alignment=WD_ALIGN_PARAGRAPH.CENTER)
        return

    if text_lines:
        write_text_lines(cell, text_lines, default_alignment="Center")
    else:
        cell.add_paragraph("")


def build_input_template_excel_bytes() -> bytes:
    """Create a starter Excel file that users can download and fill in."""
    columns = [
        "CircleLine1",
        "CircleLine2MainInfo",
        "CircleLine3",
        "RectangleLine1MainInfo",
        "RectangleLine2",
        "RectangleLine3",
        "RectangleLine4",
        "RectangleLine5",
        "SetID",
        "UniqueID",
    ]
    example_rows = [
        [
            "ELN:",
            "Main Info 1",
            "(Optional)",
            "Main Info 1",
            "Detailed Info: concentration / solvent / condition",
            "Storage Info",
            "Expiration Date",
            "ELN: DD/MM/YYYY",
            "ExampleSet",
            "IT000001",
        ],
        [
            "ELN:",
            "Main Info 2",
            "(Optional)",
            "Main Info 2",
            "Detailed Info: concentration / solvent / condition",
            "Storage Info",
            "Expiration Date",
            "ELN: DD/MM/YYYY",
            "ExampleSet",
            "IT000002",
        ],
    ]
    df = pd.DataFrame(example_rows, columns=columns)
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Labels")
        ws = writer.sheets["Labels"]
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        widths = {
            "A": 16, "B": 20, "C": 16, "D": 24, "E": 42, "F": 22,
            "G": 22, "H": 22, "I": 16, "J": 16,
        }
        for col_letter, width in widths.items():
            ws.column_dimensions[col_letter].width = width
        for cell in ws[1]:
            cell.font = cell.font.copy(bold=True)
        note_sheet = writer.book.create_sheet("Instructions")
        note_rows = [
            ["How to use this template"],
            ["Fill one row per label. Keep column names unchanged for easiest mapping."],
            ["CircleLine1, CircleLine2MainInfo, and CircleLine3 can become the three circle/lid text lines. If lid QR is enabled, CircleLine3 is reserved for the QR and is not printed as text."],
            ["RectangleLine1MainInfo and RectangleLine2-5 become rectangle lines."],
            ["SetID is optional and can group labels by collection type, experiment, or batch."],
            ["UniqueID is optional. When provided, it is always encoded first because it is the safest database identifier. If it is absent, QR output can still encode descriptive label information up to the configured character limit."],
            ["Non-main info can be left blank. Examples marked (Optional) are placeholders only."],
        ]
        for r, values in enumerate(note_rows, start=1):
            for c, value in enumerate(values, start=1):
                note_sheet.cell(row=r, column=c).value = value
        note_sheet.column_dimensions["A"].width = 100
    output.seek(0)
    return output.getvalue()


def read_input_table(uploaded_file) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)
    df = df.dropna(axis=1, how="all")
    df = df.dropna(axis=0, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    df = df.loc[:, [not str(c).startswith("Unnamed") for c in df.columns]]
    return df.reset_index(drop=True)


def set_mapping_key(value: Any, setid_col: Optional[str]) -> str:
    """Return a stable formatting-map key for a SetID value."""
    if not setid_col:
        return ALL_LABELS_SET_KEY
    text = clean_cell_value(value).strip()
    return text if text else UNASSIGNED_SET_KEY


def detected_set_mapping_keys(df: pd.DataFrame, setid_col: Optional[str]) -> List[str]:
    """Return SetID keys in first-appearance order, including blank rows as Unassigned."""
    if not setid_col:
        return [ALL_LABELS_SET_KEY]
    keys: List[str] = []
    for value in df[setid_col].tolist():
        key = set_mapping_key(value, setid_col)
        if key not in keys:
            keys.append(key)
    return keys or [UNASSIGNED_SET_KEY]


def display_set_mapping_key(key: str) -> str:
    if key == ALL_LABELS_SET_KEY:
        return "All labels"
    if key == UNASSIGNED_SET_KEY:
        return "Unassigned"
    return str(key)


def mapping_columns_match(mapping_df: pd.DataFrame, expected_columns: List[str]) -> bool:
    if mapping_df is None or mapping_df.empty or "source_column" not in mapping_df.columns:
        return False
    return list(mapping_df["source_column"].astype(str)) == list(map(str, expected_columns))


def detect_position_columns(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    normalized = {str(c).strip().lower().replace(" ", "_"): c for c in df.columns}
    sheet_col = normalized.get("sheet") or normalized.get("print_sheet")
    row_col = normalized.get("row") or normalized.get("print_row")
    col_col = normalized.get("column") or normalized.get("label_column") or normalized.get("print_column")
    return {"Sheet": sheet_col, "Row": row_col, "Label column": col_col}


def default_mapping_for_columns(columns: List[str], setid_col: Optional[str] = None) -> pd.DataFrame:
    rows = []
    printable_cols = [c for c in columns if c != setid_col]
    position_names = {"sheet", "print_sheet", "row", "print_row", "column", "label_column", "print_column"}
    printable_cols = [c for c in printable_cols if str(c).strip().lower().replace(" ", "_") not in position_names]

    # Preferred defaults for the downloadable spreadsheet template.
    # These are intentionally name-aware so the template maps correctly even if
    # users rearrange the columns in Excel.
    named_defaults = {
        "circleline1": {"part": "Circle", "line": 1, "align": "Center", "font_size": 5.0, "bold": False},
        "circleline2maininfo": {"part": "Circle", "line": 2, "align": "Center", "font_size": 7.0, "bold": True},
        "circleline3": {"part": "Circle", "line": 3, "align": "Center", "font_size": 5.0, "bold": False},
        "rectangleline1maininfo": {"part": "Rectangle", "line": 1, "align": "Left", "font_size": 7.0, "bold": True},
        "rectangleline2": {"part": "Rectangle", "line": 2, "align": "Left", "font_size": 6.0, "bold": False},
        "rectangleline3": {"part": "Rectangle", "line": 3, "align": "Left", "font_size": 6.0, "bold": False},
        "rectangleline4": {"part": "Rectangle", "line": 4, "align": "Left", "font_size": 6.0, "bold": False},
        "rectangleline5": {"part": "Rectangle", "line": 5, "align": "Left", "font_size": 6.0, "bold": False},
    }

    for idx, col in enumerate(printable_cols):
        norm = normalize_column_name(col).replace("_", "")
        if norm in named_defaults:
            cfg = named_defaults[norm]
            part = cfg["part"]
            line = cfg["line"]
            align = cfg["align"]
            font_size = cfg["font_size"]
            bold = cfg["bold"]
        elif idx < MAX_CIRCLE_LINES:
            part = "Circle"
            line = idx + 1
            align = "Center"
            font_size = 7.0 if idx == 0 else 5.0
            bold = True if idx == 0 else False
        elif idx < MAX_CIRCLE_LINES + MAX_RECTANGLE_LINES:
            part = "Rectangle"
            line = idx - MAX_CIRCLE_LINES + 1
            align = "Left"
            font_size = 7.0 if line == 1 else 6.0
            bold = True if line == 1 else False
        else:
            part = "Ignore"
            line = 1
            align = "Left"
            font_size = 6.0
            bold = False
        rows.append({
            "source_column": col,
            "print_part": part,
            "line": line,
            "side": "Left/new line",
            "font_size": font_size,
            "bold": bold,
            "align": align,
            "color": "#000000",
            "tab_pos": 1200,
        })
    return pd.DataFrame(rows)


def normalize_mapping(mapping_df: pd.DataFrame) -> pd.DataFrame:
    mapping_df = mapping_df.copy()
    for col, default in [
        ("print_part", "Ignore"), ("line", 1), ("side", "Left/new line"),
        ("font_size", 6.0), ("bold", False), ("align", "Left"),
        ("color", "#000000"), ("tab_pos", 1200)
    ]:
        if col not in mapping_df.columns:
            mapping_df[col] = default
    mapping_df["print_part"] = mapping_df["print_part"].where(mapping_df["print_part"].isin(PRINT_PARTS), "Ignore")
    mapping_df["side"] = mapping_df["side"].where(mapping_df["side"].isin(SIDES), "Left/new line")
    mapping_df["line"] = pd.to_numeric(mapping_df["line"], errors="coerce").fillna(1).astype(int)
    mapping_df["font_size"] = pd.to_numeric(mapping_df["font_size"], errors="coerce").fillna(6.0)
    mapping_df["bold"] = mapping_df["bold"].fillna(False).astype(bool)
    mapping_df["align"] = mapping_df["align"].where(mapping_df["align"].isin(DISPLAY_ALIGNMENTS), "Left")
    mapping_df["color"] = mapping_df["color"].apply(normalize_hex_color)
    mapping_df["tab_pos"] = pd.to_numeric(mapping_df["tab_pos"], errors="coerce").fillna(1200).astype(int).clip(300, 2200)
    return mapping_df


def mapping_warnings(mapping_df: pd.DataFrame) -> List[str]:
    warnings = []
    active = mapping_df[mapping_df["print_part"] != "Ignore"].copy()
    circle_lines = set(active.loc[active["print_part"] == "Circle", "line"].astype(int).tolist())
    rect_lines = set(active.loc[active["print_part"] == "Rectangle", "line"].astype(int).tolist())
    if any(line < 1 or line > MAX_CIRCLE_LINES for line in circle_lines):
        warnings.append("Circle lines must be 1 to 3. Columns mapped outside that range will not be written correctly.")
    if any(line < 1 or line > MAX_RECTANGLE_LINES for line in rect_lines):
        warnings.append("Rectangle lines must be 1 to 6. Columns mapped outside that range will not be written correctly.")
    for part, max_lines in [("Circle", MAX_CIRCLE_LINES), ("Rectangle", MAX_RECTANGLE_LINES)]:
        part_df = active[active["print_part"] == part]
        for line in range(1, max_lines + 1):
            left_count = len(part_df[(part_df["line"] == line) & (part_df["side"] == "Left/new line")])
            right_count = len(part_df[(part_df["line"] == line) & (part_df["side"] == "Right/tab on same line")])
            if left_count > 1:
                warnings.append(f"{part} line {line} has more than one left/new-line column. Only the first one will be used.")
            if right_count > 1:
                warnings.append(f"{part} line {line} has more than one right/tab column. Only the first one will be used.")
    return warnings


def build_label_lines_for_row(row: pd.Series, mapping_df: pd.DataFrame, part: str) -> List[Dict[str, Any]]:
    max_lines = MAX_CIRCLE_LINES if part == "Circle" else MAX_RECTANGLE_LINES
    lines = []
    part_df = mapping_df[mapping_df["print_part"] == part].copy()
    for line_num in range(1, max_lines + 1):
        line_map = part_df[part_df["line"].astype(int) == line_num]
        left_maps = line_map[line_map["side"] == "Left/new line"]
        right_maps = line_map[line_map["side"] == "Right/tab on same line"]
        if left_maps.empty and right_maps.empty:
            continue

        style_source = left_maps.iloc[0] if not left_maps.empty else right_maps.iloc[0]
        left_text = clean_cell_value(row.get(style_source["source_column"], "")) if not left_maps.empty else ""
        right_text = clean_cell_value(row.get(right_maps.iloc[0]["source_column"], "")) if not right_maps.empty else ""
        if not left_text and not right_text:
            continue
        lines.append({
            "line_num": int(line_num),
            "left_text": left_text,
            "right_text": right_text,
            "font_size": float(style_source.get("font_size", 6.0)),
            "bold": bool(style_source.get("bold", False)),
            "align": style_source.get("align", "Center" if part == "Circle" else "Left"),
            "color": normalize_hex_color(style_source.get("color", "#000000")),
            "tab_pos": int(style_source.get("tab_pos", 1200)),
        })
    return lines


def positions_for_count(count: int, occupied: set, skip_occupied: bool, start_sheet: int, start_row: int, start_col: int) -> List[Tuple[int, int, int]]:
    positions = []
    sheet, row, col = int(start_sheet), int(start_row), int(start_col)
    guard = 0
    while len(positions) < count:
        guard += 1
        if guard > count + 50000:
            raise ValueError("Could not find enough available label positions.")
        candidate = (sheet, row, col)
        if not (skip_occupied and candidate in occupied):
            positions.append(candidate)
        sheet, row, col = next_position(sheet, row, col)
    return positions


def build_labels_from_input(df: pd.DataFrame, mapping_by_set: Dict[str, pd.DataFrame], setid_col: Optional[str], unique_id_col: Optional[str], occupied: set, skip_occupied: bool, start_sheet: int, start_row: int, start_col: int) -> pd.DataFrame:
    if not mapping_by_set:
        raise ValueError("No formatting map is available.")
    normalized_maps = {key: normalize_mapping(value) for key, value in mapping_by_set.items()}
    fallback_mapping = next(iter(normalized_maps.values()))
    positions = positions_for_count(len(df), occupied, skip_occupied, start_sheet, start_row, start_col)
    rows = []
    for i, (_, source_row) in enumerate(df.iterrows()):
        sheet, row, col = positions[i]
        set_id = clean_cell_value(source_row.get(setid_col, "")) if setid_col else ""
        map_key = set_mapping_key(source_row.get(setid_col, "") if setid_col else "", setid_col)
        mapping_df = normalized_maps.get(map_key, fallback_mapping)
        circle_lines = build_label_lines_for_row(source_row, mapping_df, "Circle")
        rect_lines = build_label_lines_for_row(source_row, mapping_df, "Rectangle")
        unique_id = clean_cell_value(source_row.get(unique_id_col, "")) if unique_id_col else ""
        rows.append({
            "Use": True,
            "Input row": i + 2,
            "Set ID": set_id,
            "Unique ID": unique_id,
            "Sheet": sheet,
            "Row": row,
            "Label column": col,
            "Circle text": flatten_label_text(circle_lines),
            "Rectangle text": flatten_label_text(rect_lines),
            "Circle lines JSON": repr(circle_lines),
            "Rectangle lines JSON": repr(rect_lines),
        })
    return pd.DataFrame(rows)

def parse_lines_repr(value: Any) -> List[Dict[str, Any]]:
    if isinstance(value, list):
        return value
    try:
        parsed = eval(str(value), {"__builtins__": {}})  # generated by this app only; no user-facing Python needed.
        if isinstance(parsed, list):
            return parsed
    except Exception:
        return []
    return []


def layout_warnings(layout_df: pd.DataFrame, occupied: set) -> List[str]:
    warnings = []
    if layout_df.empty:
        return ["No labels were generated."]
    active = layout_df[layout_df.get("Use", True)].copy()
    bad = active[(active["Sheet"] < 1) | (active["Row"] < 1) | (active["Row"] > 20) | (active["Label column"] < 1) | (active["Label column"] > 5)]
    if not bad.empty:
        warnings.append("Some positions are outside the valid range. Sheet must be at least 1, row must be 1 to 20, and label column must be 1 to 5.")
    duplicated = active.groupby(["Sheet", "Row", "Label column"]).size().reset_index(name="n")
    duplicated = duplicated[duplicated["n"] > 1]
    if not duplicated.empty:
        warnings.append("Some labels target the same printed position. Fix duplicates before generating.")
    hits = []
    for _, r in active.iterrows():
        candidate = (int(r["Sheet"]), int(r["Row"]), int(r["Label column"]))
        if candidate in occupied:
            hits.append(candidate)
    if hits:
        preview = ", ".join([f"sheet {s}, row {r}, column {c}" for s, r, c in hits[:10]])
        warnings.append(f"Some target labels already contain text in the uploaded template: {preview}.")
    return warnings



def character_fit_warnings(layout_df: pd.DataFrame, char_limits: Dict[str, Dict[str, int]]) -> List[str]:
    warnings = []
    if layout_df.empty:
        return warnings
    active = layout_df[layout_df.get("Use", True)].copy()
    for _, r in active.iterrows():
        label_pos = f"sheet {int(r.get('Sheet', 1))}, row {int(r.get('Row', 1))}, column {int(r.get('Label column', 1))}"
        for part, json_col in [("Circle", "Circle lines JSON"), ("Rectangle", "Rectangle lines JSON")]:
            for line_idx, line in enumerate(parse_lines_repr(r.get(json_col, "[]")), start=1):
                text = " ".join([str(line.get("left_text", "")), str(line.get("right_text", ""))]).strip()
                if not text:
                    continue
                size_key = str(int(round(float(line.get("font_size", 6.0)))))
                limit = int(char_limits.get(part, {}).get(size_key, 9999))
                length = printable_length(text)
                if length > limit:
                    warnings.append(f"{label_pos}: {part} line {line_idx} has about {length} characters at font {size_key}; suggested max is {limit}.")
    return warnings

def fill_from_layout(
    template_bytes: bytes,
    layout_df: pd.DataFrame,
    allow_overwrite: bool,
    add_qr_codes: bool = False,
    rectangle_qr_size_inches: float = DEFAULT_RECTANGLE_QR_SIZE_INCHES,
    add_circle_qr_codes: bool = False,
    circle_qr_small_inches: float = DEFAULT_CIRCLE_QR_SMALL_INCHES,
    circle_qr_main_inches: float = DEFAULT_CIRCLE_QR_MAIN_INCHES,
    max_qr_payload_chars: int = DEFAULT_MAX_QR_PAYLOAD_CHARS,
) -> bytes:
    doc = Document(io.BytesIO(template_bytes))
    errors = validate_template(doc)
    if errors:
        raise ValueError("Template validation failed: " + " ".join(errors))
    active = layout_df[layout_df.get("Use", True)].copy()
    if active.empty:
        raise ValueError("No active labels to write.")
    occupied = get_existing_occupied_positions(template_bytes)
    warnings = layout_warnings(active, occupied)
    blocking = [w for w in warnings if "outside the valid range" in w or "same printed position" in w]
    if blocking:
        raise ValueError(" ".join(blocking))
    if not allow_overwrite:
        occupied_warnings = [w for w in warnings if "already contain text" in w]
        if occupied_warnings:
            raise ValueError(occupied_warnings[0] + " Enable overwrite to continue.")
    ensure_sheet_count(doc, int(active["Sheet"].max()))
    for _, layout_row in active.iterrows():
        sheet = int(layout_row["Sheet"])
        row_num = int(layout_row["Row"])
        label_col = int(layout_row["Label column"])
        circle_lines = parse_lines_repr(layout_row.get("Circle lines JSON", "[]"))
        rect_lines = parse_lines_repr(layout_row.get("Rectangle lines JSON", "[]"))
        qr_payload = build_qr_payload_from_layout_row(layout_row, max_qr_payload_chars) if add_qr_codes else ""
        table = doc.tables[sheet - 1]
        circle_col, rectangle_col = label_to_table_columns(label_col)
        write_circle_cell_from_lines(
            table.cell(row_num - 1, circle_col),
            circle_lines,
            qr_text=qr_payload if add_qr_codes and add_circle_qr_codes and qr_payload else "",
            add_circle_qr=bool(add_qr_codes and add_circle_qr_codes and qr_payload),
            small_qr_size_inches=circle_qr_small_inches,
            main_qr_size_inches=circle_qr_main_inches,
        )
        write_cell_from_lines(
            table.cell(row_num - 1, rectangle_col),
            rect_lines,
            qr_text=qr_payload if add_qr_codes and qr_payload else "",
            qr_size_inches=rectangle_qr_size_inches,
        )
    output = io.BytesIO()
    doc.save(output)
    return output.getvalue()


def build_inventory_table(layout_df: pd.DataFrame, include_box_layout: bool = True) -> pd.DataFrame:
    active = layout_df[layout_df.get("Use", True)].copy() if not layout_df.empty else pd.DataFrame()
    rows = []
    for idx, (_, r) in enumerate(active.reset_index(drop=True).iterrows()):
        circle_lines = parse_lines_repr(r.get("Circle lines JSON", "[]"))
        rect_lines = parse_lines_repr(r.get("Rectangle lines JSON", "[]"))
        entry = {
            "sample_id": flatten_label_text(circle_lines),
            "description": flatten_label_text(rect_lines),
        }
        unique_id = str(r.get("Unique ID", "")).strip() if "Unique ID" in r else ""
        if unique_id:
            entry["uniqueID"] = unique_id
        if "Set ID" in r and str(r.get("Set ID", "")).strip():
            entry["setID"] = str(r.get("Set ID", ""))
        if include_box_layout:
            box_col, box_row, grid_id = box_position(idx)
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


def layout_grid_html(layout_df: pd.DataFrame, occupied: set) -> str:
    active = layout_df[layout_df.get("Use", True)].copy() if not layout_df.empty else pd.DataFrame()
    if active.empty:
        return "<p>No layout generated yet.</p>"
    max_sheet = max(1, int(active["Sheet"].max()))
    html_parts = ["<style>.sheetgrid{border-collapse:collapse;margin-bottom:24px}.sheetgrid td,.sheetgrid th{border:1px solid #ddd;padding:4px;text-align:center;font-size:12px}.sheetgrid td{width:110px;height:34px}.used{background:#f7f7f7}.occupied{background:#ffe5e5}.planned{background:#e7f3ff}.conflict{background:#ffd3a8}</style>"]
    pos_to_text = {}
    duplicates = set()
    for idx, r in active.iterrows():
        key = (int(r["Sheet"]), int(r["Row"]), int(r["Label column"]))
        text = str(r.get("Circle text", ""))[:28]
        if key in pos_to_text:
            duplicates.add(key)
            pos_to_text[key] += f"<br>⚠ {text}"
        else:
            pos_to_text[key] = text
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
        div[data-baseweb="tab-list"] button[role="tab"] {
            background-color: #eeeeee;
            border: 1px solid #b8b8b8;
            border-radius: 0.55rem 0.55rem 0 0;
            padding: 0.55rem 0.9rem;
        }
        div[data-baseweb="tab-list"] button[role="tab"] p {
            color: #8A1538;
            font-weight: 800;
        }
        div[data-baseweb="tab-list"] button[role="tab"][aria-selected="true"] {
            background-color: #4a4a4a;
            border-color: #4a4a4a;
        }
        div[data-baseweb="tab-list"] button[role="tab"][aria-selected="true"] p {
            color: #ffffff;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def init_state():
    if "mapping_df" not in st.session_state:
        st.session_state.mapping_df = pd.DataFrame()
    if "mapping_by_set" not in st.session_state:
        st.session_state.mapping_by_set = {}
    if "mapping_signature" not in st.session_state:
        st.session_state.mapping_signature = None
    if "layout_df" not in st.session_state:
        st.session_state.layout_df = pd.DataFrame()
    if "generated_docx" not in st.session_state:
        st.session_state.generated_docx = None
    if "generated_inventory_xlsx" not in st.session_state:
        st.session_state.generated_inventory_xlsx = None
    if "char_limits" not in st.session_state:
        st.session_state.char_limits = copy.deepcopy(DEFAULT_CHAR_LIMITS)


def main():
    st.set_page_config(page_title="Spreadsheet to LabTAG Label Filler", layout="wide")
    init_state()
    inject_custom_css()

    st.title("Spreadsheet to LabTAG Label Filler")
    st.caption("Use Excel or CSV for data entry, autofill, copying, and QC. This app maps spreadsheet columns to the Circle and Rectangle parts of the printable LabTAG template.")

    with st.sidebar:
        st.header("Template")
        uploaded_template = st.file_uploader("Upload CryoSTUCK_labels.docx or another compatible .docx template", type=["docx"])
        use_default = st.checkbox("Use included template if no upload", value=True)
        if uploaded_template is not None:
            template_bytes = uploaded_template.read()
            st.success("Using uploaded template.")
        elif use_default and DEFAULT_TEMPLATE.exists():
            template_bytes = DEFAULT_TEMPLATE.read_bytes()
            st.success(f"Using included template: {DEFAULT_TEMPLATE.name}")
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
            st.caption(f"Detected {len(existing_occupied)} occupied label positions in this template.")
        except Exception as exc:
            st.error(f"Could not read template: {exc}")
            st.stop()

        st.header("Placement")
        start_sheet = st.number_input("Start sheet", min_value=1, max_value=50, value=1, step=1)
        start_row = st.number_input("Start row", min_value=1, max_value=20, value=1, step=1)
        start_col = st.number_input("Start label column", min_value=1, max_value=5, value=1, step=1)
        skip_occupied = st.checkbox("Skip labels that already contain text", value=True)
        allow_overwrite = st.checkbox("Allow overwrite if layout targets used labels", value=False)
        include_box_layout = st.checkbox("Add 10 x 10 box columns to inventory export", value=True)

        st.header("QR codes")
        add_qr_codes = st.checkbox("Create QR code", value=True)
        add_circle_qr_codes = st.checkbox("Also add QR code to circle/lid", value=False, help="Optional. If enabled, the same capped QR payload is also added to the lid. This is off by default because most users only need the rectangle QR. If circle line 2 is blank, the lid QR can be larger.")
        max_qr_payload_chars = st.number_input("Maximum characters encoded in each QR", min_value=8, max_value=200, value=DEFAULT_MAX_QR_PAYLOAD_CHARS, step=5, help="The QR uses available label information up to this limit. UniqueID is always placed first when present, followed by CircleLine2MainInfo, CircleLine1, CircleLine3, RectangleLine1MainInfo, and remaining rectangle lines. Extra text is truncated.")
        st.caption("QR payload order: UniqueID (when present) → CircleLine2MainInfo → CircleLine1 → CircleLine3 → RectangleLine1MainInfo → remaining rectangle lines. Blank fields are skipped. If UniqueID is absent, the QR still uses the available sample information. Long payloads are truncated.")
        rectangle_qr_size_inches = st.number_input("QR size in rectangle, inches", value=DEFAULT_RECTANGLE_QR_SIZE_INCHES, step=0.01, format="%.2f", help="Default 0.16 in. You can enter another size if your printer/scanner setup works better. Values must be greater than 0.")
        circle_qr_small_inches = st.number_input("Circle QR size when two lid text lines are present", value=DEFAULT_CIRCLE_QR_SMALL_INCHES, step=0.01, format="%.2f", help="Default 0.15 in. Used when the circle already has text in line 2. You can enter another positive size if needed.")
        circle_qr_main_inches = st.number_input("Circle QR size when circle line 2 is blank", value=DEFAULT_CIRCLE_QR_MAIN_INCHES, step=0.01, format="%.2f", help="Default 0.20 in. Used when there is no circle line 2 text, so the QR can use more lid space. You can enter another positive size if needed.")
        if any(float(v) <= 0 for v in [rectangle_qr_size_inches, circle_qr_small_inches, circle_qr_main_inches]):
            st.error("QR code sizes must be greater than 0 inches.")
        if add_qr_codes and qrcode is None:
            st.error("QR code support requires the qrcode package. Add qrcode[pil] to requirements.txt.")

        st.header("Text length checks")
        uploaded_limits = st.file_uploader("Optional JSON character limit config", type=["json"], help="Optional. Use this to tune max character warnings by label part and font size.")
        st.session_state.char_limits = load_char_limit_config(uploaded_limits)
        st.caption("Character limits are warnings only. They do not block printing.")

    st.subheader("1. Upload Excel or CSV input")
    st.download_button(
        label="Download blank Excel input template",
        data=build_input_template_excel_bytes(),
        file_name="label_input_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        help="Download a starter spreadsheet with the recommended column names and example values.",
    )
    uploaded_data = st.file_uploader("Upload label data", type=["xlsx", "csv"])
    if uploaded_data is None:
        st.info("Download the starter Excel template above or upload your own Excel/CSV file. By default, the first 3 data columns become Circle lines and the next columns become Rectangle lines.")
        st.stop()

    try:
        input_df = read_input_table(uploaded_data)
    except Exception as exc:
        st.error(f"Could not read input file: {exc}")
        st.stop()

    if input_df.empty:
        st.error("The uploaded file does not contain any usable rows.")
        st.stop()

    st.dataframe(input_df.head(25), use_container_width=True)
    st.caption(f"Loaded {len(input_df)} rows and {len(input_df.columns)} columns.")

    st.subheader("2. Confirm set ID and column mapping")
    setid_candidates = [""] + list(input_df.columns)
    default_setid_index = 0
    for idx, col in enumerate(setid_candidates):
        if str(col).strip().lower() in ["setid", "set_id", "set id"]:
            default_setid_index = idx
            break
    setid_col = st.selectbox("Optional setID column", options=setid_candidates, index=default_setid_index, help="Use this if rows belong to groups like Liver, Brain, DNA, RNA, etc. It is also included in the inventory export.")
    setid_col = setid_col or None

    unique_candidates = [""] + list(input_df.columns)
    default_unique_index = 0
    for idx, col in enumerate(unique_candidates):
        if normalize_column_name(col) in ["uniqueid", "unique_id", "uid", "qr", "qr_code"]:
            default_unique_index = idx
            break
    unique_id_col = st.selectbox("Optional uniqueID column for QR codes", options=unique_candidates, index=default_unique_index, help="Optional. When present, UniqueID is always encoded first because it is the safest identifier for database workflows. If blank or not selected, QR codes can still be generated from the label text.")
    unique_id_col = unique_id_col or None

    ignore_cols_for_mapping = [c for c in [setid_col, unique_id_col] if c]
    base_mapping = default_mapping_for_columns([c for c in input_df.columns if c != unique_id_col], setid_col)
    expected_mapping_columns = base_mapping["source_column"].astype(str).tolist()
    set_mapping_keys = detected_set_mapping_keys(input_df, setid_col)
    mapping_signature = (tuple(expected_mapping_columns), str(setid_col or ""), tuple(set_mapping_keys))
    mapping_widget_token = str(abs(hash(mapping_signature)))

    if st.session_state.mapping_signature != mapping_signature:
        previous_maps = st.session_state.mapping_by_set if isinstance(st.session_state.mapping_by_set, dict) else {}
        refreshed_maps: Dict[str, pd.DataFrame] = {}
        for key in set_mapping_keys:
            previous = previous_maps.get(key)
            if isinstance(previous, pd.DataFrame) and mapping_columns_match(previous, expected_mapping_columns):
                refreshed_maps[key] = normalize_mapping(previous)
            else:
                refreshed_maps[key] = base_mapping.copy(deep=True)
        st.session_state.mapping_by_set = refreshed_maps
        st.session_state.mapping_signature = mapping_signature
        st.session_state.layout_df = pd.DataFrame()
        st.session_state.generated_docx = None
        st.session_state.generated_inventory_xlsx = None

    if setid_col:
        st.caption("Each detected SetID has its own formatting tab. All SetIDs start with the same default mapping, font sizes, colors, and alignment, then can be edited independently.")
    else:
        st.caption("No SetID column is selected, so one formatting tab applies to all labels.")

    tab_labels = [display_set_mapping_key(key) for key in set_mapping_keys]
    if len(tab_labels) > 15:
        st.warning(f"Detected {len(tab_labels)} SetIDs. All formatting tabs are available, but the tab bar may be wide.")
    formatting_tabs = st.tabs(tab_labels)

    for set_idx, (set_key, format_tab) in enumerate(zip(set_mapping_keys, formatting_tabs)):
        with format_tab:
            set_label = display_set_mapping_key(set_key)
            if setid_col:
                if set_key == UNASSIGNED_SET_KEY:
                    row_count = int(input_df[setid_col].apply(lambda v: set_mapping_key(v, setid_col) == UNASSIGNED_SET_KEY).sum())
                else:
                    row_count = int(input_df[setid_col].apply(lambda v: set_mapping_key(v, setid_col) == set_key).sum())
                st.caption(f"Formatting for SetID: {set_label} ({row_count} label row{'s' if row_count != 1 else ''}).")
            st.caption("For tabbed lines, map one column to Left/new line and another column to Right/tab on same line using the same Circle/Rectangle line number.")

            current_mapping = normalize_mapping(st.session_state.mapping_by_set[set_key])
            editor_key = f"spreadsheet_mapping_editor_{mapping_widget_token}_{set_idx}"
            edited_mapping = st.data_editor(
                current_mapping,
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                column_config={
                    "source_column": st.column_config.TextColumn("Excel/CSV column", disabled=True),
                    "print_part": st.column_config.SelectboxColumn("Label part", options=PRINT_PARTS),
                    "line": st.column_config.NumberColumn("Line", min_value=1, max_value=MAX_RECTANGLE_LINES, step=1),
                    "side": st.column_config.SelectboxColumn("Line or tab", options=SIDES),
                    "font_size": st.column_config.NumberColumn("Font", step=0.5, help="Defaults are tuned for this label, but font size is not restricted. Use a positive point size."),
                    "bold": st.column_config.CheckboxColumn("Bold"),
                    "align": st.column_config.SelectboxColumn("Align", options=DISPLAY_ALIGNMENTS),
                    "color": st.column_config.TextColumn("Color", disabled=True, help="Use the color pickers below. The HEX value is stored here."),
                    "tab_pos": st.column_config.NumberColumn("Tab pos", min_value=300, max_value=2200, step=50),
                },
                disabled=["source_column", "color"],
                key=editor_key,
            )
            edited_mapping = normalize_mapping(edited_mapping)

            st.markdown("**Text colors**")
            st.caption("Use the HEX color wheel for any column. Black remains the default.")
            color_values = []
            picker_cols = st.columns(min(4, max(1, len(edited_mapping))))
            for row_idx, (_, map_row) in enumerate(edited_mapping.iterrows()):
                picker_col = picker_cols[row_idx % len(picker_cols)]
                with picker_col:
                    picked_color = st.color_picker(
                        str(map_row["source_column"]),
                        value=normalize_hex_color(map_row.get("color", "#000000")),
                        key=f"mapping_color_{mapping_widget_token}_{set_idx}_{row_idx}",
                        help=f"Text color for {map_row['source_column']}."
                    )
                color_values.append(normalize_hex_color(picked_color))
            if color_values:
                edited_mapping.loc[:, "color"] = color_values

            st.session_state.mapping_by_set[set_key] = normalize_mapping(edited_mapping)
            if (st.session_state.mapping_by_set[set_key]["font_size"] <= 0).any():
                st.error("Font sizes must be greater than 0 pt.")
            for warning in mapping_warnings(st.session_state.mapping_by_set[set_key]):
                st.warning(warning)
            st.caption("Circle can use 3 text lines when lid QR is off. When lid QR is on, circle line 3 is reserved for the QR and is not printed as text. Rectangle can use lines 1 to 6. Line spacing in the DOCX output is fixed at 0.75.")

    if unique_id_col:
        st.caption(f"UniqueID column: {unique_id_col}. Non-empty IDs are always encoded first. Rows without an ID can still get a QR from their descriptive label text.")
    else:
        st.caption("No UniqueID column selected. QR codes can still be generated from label text. Add a UniqueID column when you need safe database-level item identification.")

    st.subheader("3. Build preview")
    if st.button("Build printable layout", type="primary"):
        try:
            st.session_state.layout_df = build_labels_from_input(
                input_df,
                st.session_state.mapping_by_set,
                setid_col,
                unique_id_col,
                existing_occupied,
                skip_occupied,
                int(start_sheet),
                int(start_row),
                int(start_col),
            )
            st.session_state.generated_docx = None
            st.session_state.generated_inventory_xlsx = None
            st.success("Printable layout generated.")
        except Exception as exc:
            st.error(str(exc))

    if not st.session_state.layout_df.empty:
        tab_preview, tab_grid, tab_advanced = st.tabs(["Editable layout", "Sheet map", "Advanced line data"])
        with tab_preview:
            display_cols = ["Use", "Input row", "Set ID", "Unique ID", "Sheet", "Row", "Label column", "Circle text", "Rectangle text"]
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
                disabled=["Input row", "Set ID", "Unique ID", "Circle text", "Rectangle text"],
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
            char_warnings = character_fit_warnings(st.session_state.layout_df, st.session_state.char_limits)
            if char_warnings:
                with st.expander(f"Text length warnings ({len(char_warnings)})", expanded=False):
                    for warning in char_warnings[:100]:
                        st.warning(warning)
                    if len(char_warnings) > 100:
                        st.caption("Only the first 100 warnings are shown.")

        with tab_grid:
            st.markdown(layout_grid_html(st.session_state.layout_df, existing_occupied), unsafe_allow_html=True)
            st.caption("Blue means planned labels. Red means existing text from the uploaded template. Orange means a conflict or duplicate.")

        with tab_advanced:
            st.caption("Advanced. You usually do not need to edit this. It stores the actual line formatting sent to Word.")
            json_cols = ["Input row", "Circle lines JSON", "Rectangle lines JSON"]
            edited_json = st.data_editor(
                st.session_state.layout_df[json_cols],
                use_container_width=True,
                hide_index=True,
                num_rows="fixed",
                disabled=["Input row"],
                key="json_layout_editor",
            )
            for col in ["Circle lines JSON", "Rectangle lines JSON"]:
                st.session_state.layout_df[col] = edited_json[col]

        st.subheader("4. Generate files")
        if st.button("Generate filled DOCX and inventory table", type="primary"):
            try:
                invalid_font_sets = [
                    display_set_mapping_key(key)
                    for key, mapping in st.session_state.mapping_by_set.items()
                    if (normalize_mapping(mapping)["font_size"] <= 0).any()
                ]
                if invalid_font_sets:
                    raise ValueError("Font sizes must be greater than 0 pt. Check: " + ", ".join(invalid_font_sets))
                if add_qr_codes and float(rectangle_qr_size_inches) <= 0:
                    raise ValueError("Rectangle QR size must be greater than 0 inches.")
                if add_qr_codes and add_circle_qr_codes and (float(circle_qr_small_inches) <= 0 or float(circle_qr_main_inches) <= 0):
                    raise ValueError("Circle/lid QR sizes must be greater than 0 inches.")
                output_bytes = fill_from_layout(
                    template_bytes,
                    st.session_state.layout_df,
                    allow_overwrite,
                    add_qr_codes=add_qr_codes,
                    rectangle_qr_size_inches=float(rectangle_qr_size_inches),
                    add_circle_qr_codes=add_circle_qr_codes,
                    circle_qr_small_inches=float(circle_qr_small_inches),
                    circle_qr_main_inches=float(circle_qr_main_inches),
                    max_qr_payload_chars=int(max_qr_payload_chars),
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


if __name__ == "__main__":
    main()
