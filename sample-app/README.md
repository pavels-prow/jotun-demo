# RPA Demo Desktop App

A simple PySide6 desktop app that generates fake data, displays it in a table, and exports to a raw `.xlsx` file using `openpyxl`.

## One-command run (Mac / Windows)

Mac:

```bash
./run_mac.sh
```

Windows (PowerShell):

```powershell
.\run_windows.ps1
```

These scripts create a `.venv`, install requirements, and launch the app.

## Mac development

```bash
python3.11 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
python app.py
```

## Windows build (PyInstaller)

```powershell
py -3.11 -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
pyinstaller --noconsole --onefile --name RpaDemoApp app.py
```
