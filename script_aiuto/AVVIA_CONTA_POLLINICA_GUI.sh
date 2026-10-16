#!/bin/bash

# Script per avviare l'interfaccia grafica della conta pollinica

SCRIPT_DIR="$(cd "$(dirname "$(readlink -f "$0")")" && pwd)"
GUI="$SCRIPT_DIR/../codice/polline_counter_gui.py"

# Controlla tkinter
if ! python3 -c "import tkinter" 2>/dev/null; then
    echo "ERRORE: tkinter non e' installato"
    echo ""
    echo "Installa con:"
    echo "  sudo apt install python3-tk        (Debian/Ubuntu)"
    echo "  sudo zypper install python3-tk     (openSUSE)"
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

# Opzionale: bollettini Word. Il programma parte comunque.
if ! python3 -c "import docx" 2>/dev/null; then
    echo "Nota: python-docx non e' installato: i bollettini Word non saranno disponibili."
    echo "  sudo zypper install python3-python-docx   (openSUSE)"
    echo "  sudo apt install python3-docx             (Debian/Ubuntu)"
    echo ""
fi

# Opzionale: aspetto moderno (stile Windows 11). Il programma parte comunque.
if ! python3 -c "import sv_ttk" 2>/dev/null; then
    echo "Nota: per l'aspetto moderno (stile Windows 11) installa sv-ttk:"
    echo "  pip3 install --user sv-ttk"
    echo "  (se pip rifiuta con 'externally-managed-environment' aggiungi --break-system-packages)"
    echo ""
fi

# Opzionale: lettura vocale. Si controlla solo se il modello vocale e' stato
# messo in codice/modelli (cioe' se l'utente vuole la voce).
if ls "$SCRIPT_DIR"/../codice/modelli/vosk-model* >/dev/null 2>&1; then
    if ! python3 -c "import vosk, sounddevice, pyttsx3" 2>/dev/null; then
        echo "Nota: per la lettura vocale mancano delle librerie. Installa con:"
        echo "  pip3 install --user vosk sounddevice pyttsx3"
        echo "  (se pip rifiuta con 'externally-managed-environment' aggiungi --break-system-packages)"
        echo ""
    fi
    if ! command -v espeak-ng >/dev/null 2>&1; then
        echo "Nota: per la conferma parlata serve espeak-ng:"
        echo "  sudo zypper install espeak-ng   (openSUSE)   |   sudo apt install espeak-ng   (Debian/Ubuntu)"
        echo ""
    fi
    if ! command -v paplay >/dev/null 2>&1 && ! command -v pw-play >/dev/null 2>&1 \
            && ! command -v aplay >/dev/null 2>&1; then
        echo "Nota: non trovo un programma per riprodurre l'audio (beep e voce resteranno muti)."
        echo "  sudo zypper install alsa-utils   (openSUSE)   |   sudo apt install alsa-utils   (Debian/Ubuntu)"
        echo ""
    fi
fi

python3 "$GUI"
