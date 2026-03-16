#!/bin/bash

# Script per avviare l'interfaccia grafica della conta pollinica

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
GUI="$SCRIPT_DIR/../codice/polline_counter_gui.py"

# Controlla tkinter
if ! python3 -c "import tkinter" 2>/dev/null; then
    echo "ERRORE: tkinter non e' installato"
    echo ""
    echo "Installa con:"
    echo "  sudo apt install python3-tk"
    echo ""
    exit 1
fi

# Controlla openpyxl
if ! python3 -c "import openpyxl" 2>/dev/null; then
    echo "ERRORE: openpyxl non e' installato"
    echo ""
    echo "Installa con:"
    echo "  pip3 install openpyxl"
    echo ""
    exit 1
fi

python3 "$GUI"
