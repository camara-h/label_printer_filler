# LabTAG LCS-125WH Label Filler

Streamlit app for filling the LabTAG LCS-125WH Word template while preserving the original `.docx` table geometry.
App can be accessed at https://labelprinterfiller.streamlit.app/ (it may need dot be awaken)

## Install

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Main workflow

1. Use the included LCS-125WH template or upload a `.docx` template.
2. Add one or more Label ID Sets, for example Liver and Brain.
3. For each set, choose start sheet, start row, start label column, and number of labels.
4. Add circle and rectangle lines with font size, bold, alignment, serialization, and text color.
5. Build the editable layout.
6. Review the editable layout and sheet map.
7. Manually adjust sheet, row, or label column if needed.
8. Generate and download the final `.docx`.

## Notes

- Text color uses Streamlit's built-in `st.color_picker`, with black as the default.
- The app fills labels top to bottom, then left to right.
- Existing text in the uploaded template is detected and can be skipped automatically.
- If more labels are needed than fit on the current page, the app adds additional blank template pages.
- The editable layout tab is for placement edits.
- The advanced text editing tab allows final line-level text edits using JSON lists while preserving the original formatting rules for each line.
- Always test print on plain paper first using actual size or 100 percent scaling.


## v4 printing safeguards

This version adds practical limits based on the physical label size:

- Circle labels are limited to 3 lines. This is a hard limit because additional lines usually do not print reliably.
- Rectangle labels are limited to 6 lines, with an in-app warning after 5 lines.
- Font sizes are limited to 4 to 7 pt. The app warns when 4 pt is used because it may be hard to read after printing.
- Text color is still controlled per line using the HEX color picker.
