# LabTAG LCS-125WH label filler

This Streamlit app fills the LabTAG LCS-125WH Word template while preserving the original table geometry, margins, rows, columns, and cell sizes from the uploaded `.docx` template.

## Run locally

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Use

1. Open the app.
2. Use the included `Letter-125-NO0424.docx` template or upload another copy of the official template.
3. Choose start row and label column.
4. Enter the number of labels to create.
5. Configure circle and rectangle lines.
6. Check `serialize trailing number` only for lines that should increment, for example `Tissue 1`.
7. Generate and download the filled `.docx`.

The app fills labels top to bottom, then left to right.

Logical label columns map to the Word table like this:

Label 1 uses table columns 1 and 2. Column 3 is spacer.
Label 2 uses table columns 4 and 5. Column 6 is spacer.
Label 3 uses table columns 7 and 8. Column 9 is spacer.
Label 4 uses table columns 10 and 11. Column 12 is spacer.
Label 5 uses table columns 13 and 14.

## Notes

Always test print on regular paper first and hold it behind the label sheet against a light source before printing on real labels.
# label_printer_filler
