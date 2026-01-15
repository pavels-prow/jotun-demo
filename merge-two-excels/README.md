# Merge Two Excels

Small Python CLI utility that merges data into an Excel template while preserving the template structure.

## Setup

Mac/Linux:

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Windows (PowerShell):

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Quick start

Mac/Linux:

```bash
./run_mac.sh --template ../template.xlsx --data ../data.xlsx --outdir ./out
```

Windows (PowerShell):

```powershell
./run_windows.ps1 --template ..\template.xlsx --data ..\data.xlsx --outdir .\out
```

## Usage

Mac/Linux:

```bash
python merge_excel.py --template ../template.xlsx --data ../data.xlsx --outdir ./out
```

Windows (PowerShell):

```powershell
python merge_excel.py --template ..\template.xlsx --data ..\data.xlsx --outdir .\out
```

Options:

- `--template` (required): path to template .xlsx
- `--data` (required): path to data .xlsx
- `--outdir` (optional): output directory (default: current directory)
- `--sheet` (optional): target sheet name in template (default: "data")
- `--prefix` (optional): output filename prefix (default: empty)
- `--dry-run` (optional flag): validate and print actions without writing output

## Notes

- Only `.xlsx` files are supported. `.xls` is not supported by openpyxl.
