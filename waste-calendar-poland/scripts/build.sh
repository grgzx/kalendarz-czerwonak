#!/usr/bin/env bash
# Rebuild the assessment workbook from the GUS tables and the verified sample.
set -euo pipefail
cd "$(dirname "$0")/.."

python3 scripts/extract_gus.py
python3 scripts/build_workbook.py

# Recalculate formulas so the file also shows values in viewers that do not
# evaluate formulas (Preview, some web previews). Optional: needs LibreOffice.
if command -v soffice >/dev/null 2>&1; then
    tmp="$(mktemp -d)"
    if soffice --headless --calc --convert-to xlsx:"Calc MS Excel 2007 XML" \
        --outdir "$tmp" waste-calendar-poland-assessment.xlsx >/dev/null 2>&1 \
        && [ -f "$tmp/waste-calendar-poland-assessment.xlsx" ]; then
        cp "$tmp/waste-calendar-poland-assessment.xlsx" waste-calendar-poland-assessment.xlsx
        echo "przeliczono formuly (LibreOffice)"
    fi
    rm -rf "$tmp"
else
    echo "uwaga: brak soffice, formuly nie sa przeliczone (Excel i tak je policzy)"
fi
