# Purpose

Export a single PDF from an Excel `.xlsx` report using LibreOffice in headless mode, reliably targeting only the first worksheet.

# What we wanted to achieve

- Always export the first worksheet only (index 0).
- Avoid Microsoft Excel / COM; keep it cross-platform (macOS/Windows).
- Use LibreOffice (`soffice`) for conversion.
- Make the export deterministic by hiding all other sheets in a temporary copy.
- Keep the workflow CLI-friendly and scriptable.

# How it works (high level)

1) Pick the newest `.xlsx` from `../out` by default (or use `--input`).
2) Create a temporary copy where all other sheets are hidden.
3) Run `soffice --headless --convert-to pdf` on the temp file.
4) Move the produced PDF to the requested output path (or next to the input if `--output` is omitted).

# Notes

- Only `.xlsx` is supported.
- Output may differ from Excel due to LibreOffice rendering and fonts.
