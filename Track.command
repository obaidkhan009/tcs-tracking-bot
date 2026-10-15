#!/bin/bash
# Double-click launcher for the tracking bot (macOS).
cd "$(dirname "$0")"
if [ ! -x .venv/bin/python ]; then
  echo "Setting up for the first time..."
  python3 -m venv .venv && .venv/bin/pip install -q -r requirements.txt
fi
echo "Drag your Excel file here and press Enter (or press Enter to use shipments_sample.xlsx):"
read -r FILE
FILE="${FILE:-shipments_sample.xlsx}"
FILE="${FILE%\'}"; FILE="${FILE#\'}"   # strip quotes Finder adds
.venv/bin/python tracker.py "$FILE"
