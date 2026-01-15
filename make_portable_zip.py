#!/usr/bin/env python3
import argparse
import fnmatch
import os
import sys
import zipfile
from datetime import datetime
from pathlib import Path


DEFAULT_EXCLUDE_GLOBS = [
    ".git/**",
    "**/.git/**",
    ".venv/**",
    "**/.venv/**",
    "__pycache__/**",
    "**/__pycache__/**",
    ".pytest_cache/**",
    "**/.pytest_cache/**",
    ".mypy_cache/**",
    "**/.mypy_cache/**",
    ".ruff_cache/**",
    "**/.ruff_cache/**",
    ".idea/**",
    "**/.idea/**",
    ".vscode/**",
    "**/.vscode/**",
    "node_modules/**",
    "**/node_modules/**",
    "dist/**",
    "**/dist/**",
    "build/**",
    "**/build/**",
    "*.pyc",
    "*.pyo",
    "*.DS_Store",
    "*.zip",
]


def parse_args(argv: list[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Create a portable .zip of the repo without virtualenvs/caches."
    )
    parser.add_argument(
        "--src",
        default=".",
        help="Source directory to zip (default: current directory)",
    )
    parser.add_argument(
        "--output",
        help="Output .zip path (default: <repo>_portable_<timestamp>.zip in current dir)",
    )
    parser.add_argument(
        "--exclude",
        action="append",
        default=[],
        help="Additional exclude glob (can be repeated)",
    )
    parser.add_argument(
        "--overwrite",
        action="store_true",
        help="Overwrite output zip if it already exists",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Do not write zip; only print what would be included",
    )
    return parser.parse_args(argv)


def should_exclude(rel_posix: str, exclude_globs: list[str]) -> bool:
    for pattern in exclude_globs:
        if fnmatch.fnmatch(rel_posix, pattern):
            return True
    return False


def default_output_path(src_root: Path) -> Path:
    repo_name = src_root.name or "repo"
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
    return Path.cwd() / f"{repo_name}_portable_{timestamp}.zip"


def main(argv: list[str]) -> int:
    args = parse_args(argv)

    src_root = Path(args.src).resolve()
    if not src_root.exists():
        print(f"Error: source directory not found: {src_root}", file=sys.stderr)
        return 2
    if not src_root.is_dir():
        print(f"Error: source is not a directory: {src_root}", file=sys.stderr)
        return 2

    output_path = Path(args.output).resolve() if args.output else default_output_path(src_root)
    if output_path.exists() and not args.overwrite and not args.dry_run:
        print(f"Error: output zip already exists: {output_path}", file=sys.stderr)
        print("Use --overwrite to replace it.", file=sys.stderr)
        return 3

    exclude_globs = DEFAULT_EXCLUDE_GLOBS + list(args.exclude)

    files_to_add: list[tuple[Path, str]] = []
    total_bytes = 0

    for path in src_root.rglob("*"):
        if path.is_dir():
            continue

        try:
            rel = path.relative_to(src_root).as_posix()
        except ValueError:
            continue

        if output_path.is_relative_to(src_root) and path.resolve() == output_path:
            continue

        if should_exclude(rel, exclude_globs):
            continue

        try:
            st = path.stat()
        except OSError:
            continue

        files_to_add.append((path, rel))
        total_bytes += st.st_size

    print(f"Source:  {src_root}")
    print(f"Output:  {output_path}")
    print(f"Files:   {len(files_to_add)}")
    print(f"Size:    {total_bytes / (1024 * 1024):.2f} MiB (uncompressed)")

    if args.dry_run:
        return 0

    output_path.parent.mkdir(parents=True, exist_ok=True)
    mode = "w"
    compression = zipfile.ZIP_DEFLATED
    compresslevel = 9

    with zipfile.ZipFile(
        output_path,
        mode=mode,
        compression=compression,
        compresslevel=compresslevel,
    ) as zf:
        for path, rel in files_to_add:
            zf.write(path, arcname=rel)

    print("Done.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
