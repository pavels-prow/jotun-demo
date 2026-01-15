#!/usr/bin/env bash
set -euo pipefail

if [[ $# -eq 0 ]]; then
  python3 bootstrap.py --template ../template.xlsx --data ../data.xlsx --outdir ../out
else
  python3 bootstrap.py "$@"
fi
