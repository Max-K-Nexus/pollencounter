#!/bin/bash

# Script per avviare il sistema di conta pollinica (CLI)

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# Controlla se openpyxl e' installato
if ! python3 -c "import openpyxl" 2>/dev/null; then
    echo "ERRORE: openpyxl non e' installato"
    echo ""
    echo "Installa con:"
    echo "  pip3 install openpyxl"
    echo ""
    exit 1
fi

# Avvia lo script Python (gestisce tutto internamente)
python3 "$SCRIPT_DIR/../codice/polline_counter.py"
