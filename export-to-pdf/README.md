# Export to PDF (LibreOffice)

Cross-platform CLI utility to export PDF from .xlsx via LibreOffice in headless mode.

## Requirements

- Python 3.11+
- LibreOffice (uses `soffice`)

LibreOffice install:
- macOS: download and install from the official LibreOffice site.
- Windows: download and install from the official LibreOffice site.

macOS (Homebrew):

```bash
brew install --cask libreoffice
```

## Quick start

Mac/Linux:

```bash
./run_mac.sh
```

Windows (PowerShell):

```powershell
./run_windows.ps1
```

Default behavior:
- If `--input` is omitted, the tool picks the newest `.xlsx` from `../out`.
- If `--output` is omitted, the tool writes the `.pdf` next to the input `.xlsx`.

## Examples

```bash
python excel_to_pdf_lo.py
```

```bash
python excel_to_pdf_lo.py --soffice "/Applications/LibreOffice.app/Contents/MacOS/soffice"
```

## Finding the soffice path

- macOS (default path):
  `/Applications/LibreOffice.app/Contents/MacOS/soffice`
- Windows (default paths):
  `C:\Program Files\LibreOffice\program\soffice.exe`
  `C:\Program Files (x86)\LibreOffice\program\soffice.exe`

If LibreOffice is installed elsewhere, pass the path via `--soffice`.

## Options

- `--input` (optional): path to .xlsx (if omitted, the newest .xlsx from `--input-dir` is used)
- `--input-dir` (optional): folder to search .xlsx (default: `../out`)
- `--output` (optional): full path to .pdf (if omitted, next to input .xlsx)
- `--soffice` (optional): path to soffice
- `--sheet-index` (optional): sheet index (default: 0)
- `--keep-temp` (flag): keep temporary .xlsx
- `--timeout-seconds` (optional): conversion timeout (default: 60)

## Troubleshooting

- Fonts/layout differences: LibreOffice may render differently than Excel. Check installed fonts.
- Permissions: make sure you have write access to the output folder.
- Environment/sandbox limits: if LibreOffice does not start, check system security restrictions.

## Notes

- Only the selected sheet is exported; all other sheets are hidden temporarily.
- Only `.xlsx` is supported.
