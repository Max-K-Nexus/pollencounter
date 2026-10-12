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

# Opzionale: bollettini Word. Il programma parte comunque.
if ! python3 -c "import docx" 2>/dev/null; then
    echo "Nota: python-docx non e' installato: i bollettini Word non saranno disponibili."
    echo "  sudo zypper install python3-python-docx   (openSUSE)"
    echo "  sudo apt install python3-docx             (Debian/Ubuntu)"
    echo ""
fi

# Avvia lo script Python (gestisce tutto internamente)
python3 "$SCRIPT_DIR/../codice/polline_counter.py"
