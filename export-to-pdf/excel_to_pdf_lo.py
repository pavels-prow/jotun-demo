#!/usr/bin/env python3
import argparse
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional

from openpyxl import load_workbook

MAC_SOFFICE_CANDIDATES = [
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
    "/Applications/LibreOfficeDev.app/Contents/MacOS/soffice",
    str(Path.home() / "Applications/LibreOffice.app/Contents/MacOS/soffice"),
]
WINDOWS_SOFFICE_CANDIDATES = [
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]


def _is_executable_file(path: Path) -> bool:
    return path.exists() and path.is_file() and os.access(path, os.X_OK)


def _resolve_from_directory(directory: Path) -> Optional[str]:
    candidates = [
        directory / "soffice",
        directory / "soffice.bin",
        directory / "soffice.exe",
        directory / "program/soffice.exe",
        directory / "Contents/MacOS/soffice",
    ]
    for candidate in candidates:
        if _is_executable_file(candidate):
            return str(candidate)
    for candidate in candidates:
        if candidate.exists() and candidate.is_file():
            return str(candidate)
    return None


def resolve_soffice(user_path: Optional[str]) -> Optional[str]:
    if user_path:
        raw_path = Path(user_path)
        if raw_path.exists():
            if raw_path.is_dir():
                resolved = _resolve_from_directory(raw_path)
                return resolved
            return str(raw_path)
        from_path = shutil.which(user_path)
        return from_path

    for candidate in MAC_SOFFICE_CANDIDATES + WINDOWS_SOFFICE_CANDIDATES:
        if Path(candidate).exists():
            return candidate

    for name in ("soffice", "libreoffice", "soffice.bin"):
        from_path = shutil.which(name)
        if from_path:
            return from_path

    return None


def ensure_output_dir(output_path: Path) -> None:
    output_path.parent.mkdir(parents=True, exist_ok=True)


def prepare_temp_workbook(input_path: Path, sheet_index: int, temp_dir: Path) -> Path:
    try:
        wb = load_workbook(input_path)
    except Exception as exc:  # pragma: no cover - just pass through
        raise RuntimeError(f"Failed to open Excel file: {exc}") from exc

    if sheet_index < 0 or sheet_index >= len(wb.worksheets):
        raise RuntimeError(
            f"Invalid sheet index: {sheet_index}. Sheets available: {len(wb.worksheets)}"
        )

    wb.active = sheet_index
    for idx, ws in enumerate(wb.worksheets):
        if idx != sheet_index:
            ws.sheet_state = "hidden"

    temp_xlsx = temp_dir / f"temp_{input_path.stem}.xlsx"
    wb.save(temp_xlsx)
    return temp_xlsx


def run_soffice(
    soffice_path: str,
    input_xlsx: Path,
    temp_out_dir: Path,
    timeout_seconds: int,
) -> tuple[int, str, str]:
    command = [
        soffice_path,
        "--headless",
        "--nologo",
        "--nodefault",
        "--nolockcheck",
        "--norestore",
        "--convert-to",
        "pdf",
        "--outdir",
        str(temp_out_dir),
        str(input_xlsx),
    ]
    try:
        completed = subprocess.run(
            command,
            capture_output=True,
            text=True,
            timeout=timeout_seconds,
            check=False,
        )
    except subprocess.TimeoutExpired as exc:
        stdout = exc.stdout or ""
        stderr = exc.stderr or ""
        return 124, stdout, stderr

    return completed.returncode, completed.stdout, completed.stderr


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Export PDF from .xlsx via LibreOffice (headless)."
    )
    parser.add_argument(
        "--input",
        help="Path to .xlsx (if omitted, the newest .xlsx from --input-dir is used)",
    )
    parser.add_argument(
        "--input-dir",
        default="../out",
        help="Folder to search .xlsx files (default: ../out)",
    )
    parser.add_argument(
        "--output",
        help="Path to output .pdf (if omitted, next to input .xlsx)",
    )
    parser.add_argument(
        "--soffice",
        help="Path to soffice binary (auto-detect if omitted)",
    )
    parser.add_argument(
        "--sheet-index",
        type=int,
        default=0,
        help="Sheet index to export (default: 0)",
    )
    parser.add_argument(
        "--keep-temp",
        action="store_true",
        help="Keep temporary .xlsx for debugging",
    )
    parser.add_argument(
        "--timeout-seconds",
        type=int,
        default=60,
        help="Conversion timeout in seconds (default: 60)",
    )
    return parser.parse_args(argv)


def main(argv: list[str]) -> int:
    args = parse_args(argv)

    if args.input:
        input_path = Path(args.input)
    else:
        input_dir = Path(args.input_dir)
        if not input_dir.exists():
            print("Error: input search folder not found", file=sys.stderr)
            return 2
        candidates = list(input_dir.glob("*.xlsx"))
        if not candidates:
            print("Error: no .xlsx files found in the folder", file=sys.stderr)
            return 2
        candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
        input_path = candidates[0]

    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix(".pdf")

    print(f"Input file: {input_path}")
    print(f"Output file: {output_path}")

    if not input_path.exists():
        print("Error: input file not found", file=sys.stderr)
        return 2

    soffice_path = resolve_soffice(args.soffice)
    if not soffice_path:
        print(
            "Error: LibreOffice (soffice) not found. Install LibreOffice and "
            "provide the path via --soffice.",
            file=sys.stderr,
        )
        print("Tried common locations and PATH. Examples:", file=sys.stderr)
        print(
            '  macOS: --soffice "/Applications/LibreOffice.app/Contents/MacOS/soffice"',
            file=sys.stderr,
        )
        print(
            r'  Windows: --soffice "C:\Program Files\LibreOffice\program\soffice.exe"',
            file=sys.stderr,
        )
        return 3

    print(f"soffice: {soffice_path}")

    try:
        ensure_output_dir(output_path)
    except Exception as exc:
        print(f"Error: failed to create output directory: {exc}", file=sys.stderr)
        return 4

    temp_dir_path: Path
    temp_dir_handle = None
    if args.keep_temp:
        temp_dir_path = Path(tempfile.mkdtemp(prefix="xlsx_to_pdf_"))
    else:
        temp_dir_handle = tempfile.TemporaryDirectory(prefix="xlsx_to_pdf_")
        temp_dir_path = Path(temp_dir_handle.name)

    try:
        temp_xlsx = prepare_temp_workbook(input_path, args.sheet_index, temp_dir_path)
    except Exception as exc:
        print(f"Error: failed to prepare temporary file: {exc}", file=sys.stderr)
        if temp_dir_handle:
            temp_dir_handle.cleanup()
        return 5

    if args.keep_temp:
        print(f"Temporary file: {temp_xlsx}")

    temp_out_dir = temp_dir_path / "out"
    temp_out_dir.mkdir(parents=True, exist_ok=True)

    code, stdout, stderr = run_soffice(
        soffice_path,
        temp_xlsx,
        temp_out_dir,
        args.timeout_seconds,
    )

    if code == 124:
        print("Error: conversion timed out", file=sys.stderr)
        if args.keep_temp and stdout:
            print(stdout)
        if args.keep_temp and stderr:
            print(stderr, file=sys.stderr)
        if temp_dir_handle:
            temp_dir_handle.cleanup()
        return 4

    if code != 0:
        print(f"Error: soffice exited with code {code}", file=sys.stderr)
        if args.keep_temp and stdout:
            print(stdout)
        if args.keep_temp and stderr:
            print(stderr, file=sys.stderr)
        if temp_dir_handle:
            temp_dir_handle.cleanup()
        return 4

    produced_pdf = temp_out_dir / f"{temp_xlsx.stem}.pdf"
    if not produced_pdf.exists():
        print("Error: PDF was not created", file=sys.stderr)
        if temp_dir_handle:
            temp_dir_handle.cleanup()
        return 4

    try:
        shutil.move(str(produced_pdf), str(output_path))
    except Exception as exc:
        print(f"Error: failed to move PDF: {exc}", file=sys.stderr)
        if temp_dir_handle:
            temp_dir_handle.cleanup()
        return 4

    if temp_dir_handle:
        temp_dir_handle.cleanup()

    print("Done: PDF created successfully")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
