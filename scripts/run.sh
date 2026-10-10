#!/usr/bin/env bash
# Regenera el entregable de conciliación 60/61/2x desde el .xlsb de Odoo.
# Requiere: python3, pandas, openpyxl, pyxlsb
set -euo pipefail
cd "$(dirname "$0")/.."
WORK=$(mktemp -d)
python3 scripts/prep.py "Analisis_Cta61_Conciliacion_Compras_Ene-Set_2026_La_Iberica - copia.xlsb" "$WORK"
python3 scripts/build_data.py "$WORK"
python3 scripts/build_ac.py "$WORK"
python3 scripts/write_xlsx.py "$WORK" "Conciliacion_Compras_60_61_2x_Ene-Set_2026_La_Iberica.xlsx"
