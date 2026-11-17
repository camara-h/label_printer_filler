# Spreadsheet to LabTAG Label Filler

A Streamlit app that takes an Excel or CSV table and fills a compatible LabTAG CryoSTUCK Word template.

## Main workflow

1. Upload or use the included Word template.
2. Upload an Excel or CSV file with one row per label.
3. Confirm which spreadsheet columns map to the circle and rectangle parts of the label.
4. Build the printable layout.
5. Download the filled `.docx` template and the inventory-style `.xlsx` export.

## Template preference

If `CryoSTUCK_labels.docx` exists in the app folder, the app uses it by default. Otherwise, it falls back to `Letter-125-NO0424.docx`.


## Per-SetID formatting

If a `SetID` column is selected, the app creates one formatting tab for every unique SetID found in the uploaded table. Each tab starts from the same default column mapping, font sizes, colors, bold settings, and alignment, but can then be customized independently. Rows with a blank SetID use an **Unassigned** formatting tab. If no SetID column is selected, a single **All labels** tab is used.

The main formatting controls remain table-based. Text color is controlled with a native Streamlit HEX color picker for each source column; black (`#000000`) remains the default.

## Data Matrix codes

The app can generate compact ECC200 Data Matrix symbols. This is optimized for short cryogenic-label identifiers and uses a one-module quiet zone.

`UniqueID` is optional, but when present it is always encoded first because it is the safest identifier for database workflows. The default payload hierarchy is `UniqueID` → `CircleLine2MainInfo` → `CircleLine1` → `CircleLine3`. Rectangle text is intentionally excluded so the symbol stays compact. The payload is capped at 20 characters by default, and later fields are truncated as needed.

The rectangle Data Matrix is controlled by **Create Data Matrix**. The additional circle/lid Data Matrix is **off by default** and can be enabled separately. When the lid Data Matrix is enabled, `CircleLine3` is reserved for the symbol and is not printed as text. Both symbol sizes default to 0.15 inches.

Fields inside the encoded payload are separated with `|` rather than a literal tab so a USB HID scanner returns one clean string instead of potentially moving focus between fields.

## Font sizing and character length warnings

Font size fields are not hard-limited. The default sizes are tuned for this label, but users can enter another positive point size if their layout requires it.


The app includes approximate text-fit warnings by label part and font size. These warnings do not block printing.

You can upload a custom JSON config using this structure:

```json
{
  "Circle": {"7": 8, "6": 10, "5": 13, "4": 16},
  "Rectangle": {"7": 18, "6": 24, "5": 30, "4": 38}
}
```

A starter file is included as `character_limits_example.json`.

## Install

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Downloadable input template

The app includes a **Download blank Excel input template** button before data upload. The starter columns are:

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

The default mapping is name-aware. The circle/lid can use `CircleLine1`, `CircleLine2MainInfo`, and `CircleLine3` as text lines. If lid Data Matrix is enabled, `CircleLine3` is reserved for the Data Matrix and is not printed as text. `CircleLine2MainInfo` and `RectangleLine1MainInfo` are mapped as the bold main-information lines. `SetID` and `UniqueID` are detected separately and are not printed as regular label text unless you manually map them.
