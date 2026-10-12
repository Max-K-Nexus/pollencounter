#!/bin/bash

# Script per avviare la conta pollinica da terminale (macOS)

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

# Controlla che python3 sia disponibile
if ! command -v python3 &>/dev/null; then
    echo "ERRORE: Python 3 non trovato."
    echo ""
    echo "Installa Python 3 da:"
    echo "  https://www.python.org/downloads/"
    echo ""
    echo "Oppure, se hai Homebrew:"
    echo "  brew install python3"
    echo ""
    exit 1
fi

# Controlla openpyxl
if ! python3 -c "import openpyxl" 2>/dev/null; then
    echo "ERRORE: openpyxl non e' installato."
    echo ""
    echo "Installa con:"
    echo "  pip3 install openpyxl"
    echo ""
    exit 1
fi

# Opzionale: bollettini Word. Il programma parte comunque.
if ! python3 -c "import docx" 2>/dev/null; then
    echo "Nota: python-docx non e' installato: i bollettini Word non saranno disponibili."
    echo "  pip3 install python-docx"
    echo ""
fi

python3 "$SCRIPT_DIR/../codice/polline_counter.py"
