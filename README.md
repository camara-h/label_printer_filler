# Spreadsheet Label Filler

A Streamlit app that takes an Excel or CSV table and formats it into several lab-label Word layouts while preserving the template geometry.

## Supported layouts

### CryoSTUCK - circle + rectangle
- 20 rows × 5 labels per sheet (100 labels)
- Existing CryoSTUCK workflow is preserved
- Circle: up to 3 text lines, or optional lid Data Matrix
- Rectangle: up to 6 mapped text lines and optional Data Matrix
- Bundled template: `CryoSTUCK_labels.docx`

### Tube lids - 0.49 in grid
- 21 rows × 16 lid labels
- Without Data Matrix: up to 2 text lines
- With Data Matrix: Data Matrix + 1 text line
- Default Data Matrix size: 0.15 in
- Bundled template: `Tube_Lids_049in.docx`

### 0.2 mL tube labels
- 4 labels across
- 28 printable rows per sheet
- The final 29th template row is intentionally blocked because the source template marks it as non-printing
- One short text line, with or without an inline Data Matrix
- Default Data Matrix size: 0.12 in
- Bundled template: `Microtube_02mL_Labels.docx`

### Avery 5267 - return address labels
- 20 rows × 4 labels per sheet (80 labels)
- Up to 2 text lines
- Optional Data Matrix can be appended to the second line
- Default Data Matrix size: 0.18 in
- Bundled template: `Avery_5267_Return_Address_Labels.docx`

## Workflow

1. Choose a label layout in the sidebar.
2. Use the bundled template for that layout or upload a compatible `.docx` template.
3. Upload an Excel or CSV file with one row per label.
4. Confirm `SetID`, `UniqueID`, and column formatting.
5. Each unique `SetID` receives its own formatting tab. All tabs start from the same layout-specific defaults and can be customized independently.
6. Build the editable placement preview.
7. Generate and download the filled Word template and inventory-style Excel table.

## Spreadsheet input

The app includes a **Download blank Excel input template** button. Recommended columns are:

- `CircleLine1`
- `CircleLine2MainInfo`
- `CircleLine3`
- `RectangleLine1MainInfo`
- `RectangleLine2`
- `RectangleLine3`
- `RectangleLine4`
- `RectangleLine5`
- `SetID`
- `UniqueID`

The same spreadsheet can be used for every layout. Each layout maps only the columns that make sense for that label by default. Users can change the mapping in the formatting table.

## Layout-specific default mapping

- CryoSTUCK: uses the existing circle and rectangle defaults.
- Tube lids: `CircleLine2MainInfo` is the main lid line; `CircleLine1` is the optional second line.
- 0.2 mL: `CircleLine2MainInfo` is the single default text line.
- Avery 5267: `RectangleLine1MainInfo` is line 1. `RectangleLine2` and `RectangleLine3` share line 2 by default using a tab.

## Data Matrix

The app generates ECC200 Data Matrix symbols with `pyStrich`.

For CryoSTUCK, the existing compact payload behavior is preserved. `UniqueID` is prioritized and lid information is used after it, capped at 20 characters by default.

For single-part layouts, `UniqueID` is also encoded first when present, followed by the mapped label text up to the configured character limit.

The matrix sizes are editable. The layout defaults are intended only as useful starting points.

## Text formatting

Formatting remains table-based. Font size, bold, alignment, line assignment, tab behavior, and label part can be changed per mapped column. Text colors use Streamlit HEX color pickers. Font size is not hard-limited beyond requiring a positive value.

DOCX line spacing is fixed at 0.75.

## Install

```bash
pip install -r requirements.txt
streamlit run app.py
```
