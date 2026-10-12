#!/bin/bash

# Script per avviare l'interfaccia grafica della conta pollinica (macOS)

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
GUI="$SCRIPT_DIR/../codice/polline_counter_gui.py"

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

# Controlla tkinter
if ! python3 -c "import tkinter" 2>/dev/null; then
    echo "ERRORE: tkinter non e' installato."
    echo ""
    echo "Se hai installato Python da python.org, tkinter e' gia' incluso."
    echo "Se hai usato Homebrew, installa:"
    echo "  brew install python-tk"
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

# Opzionale: lettura vocale. Si controlla solo se il modello vocale e' stato
# messo in codice/modelli (cioe' se l'utente vuole la voce).
if ls "$SCRIPT_DIR"/../codice/modelli/vosk-model* >/dev/null 2>&1; then
    if ! python3 -c "import vosk, sounddevice, pyttsx3" 2>/dev/null; then
        echo "Nota: per la lettura vocale mancano delle librerie. Installa con:"
        echo "  pip3 install vosk sounddevice pyttsx3"
        echo "  (l'installazione puo' richiedere qualche minuto: scarica anche pyobjc)"
        echo ""
    fi
fi

python3 "$GUI"
