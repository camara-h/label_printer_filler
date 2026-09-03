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

## QR codes

If the spreadsheet has a column named `uniqueID`, `unique_id`, `uid`, `qr`, or `qr_code`, the app detects it automatically.

When QR codes are enabled, rows with a non-empty unique ID get a small QR code added to the bottom-right area of the rectangle label. Blank values are ignored.

QR code placement is intentionally small and optional. Always test print and scan before using it for a real experiment.

## Character length warnings

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
