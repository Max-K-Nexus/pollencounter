#!/bin/bash

# Script per creare l'applicazione macOS (.app) con PyInstaller
# DA ESEGUIRE SU UN MAC con Python 3 installato

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
CODICE="$SCRIPT_DIR/../codice"

echo "============================================"
echo "  CREAZIONE APP - CONTA POLLINICA (macOS)"
echo "============================================"
echo ""

# Verifica python3
if ! command -v python3 &>/dev/null; then
    echo "ERRORE: Python 3 non trovato."
    echo "Installa da: https://www.python.org/downloads/"
    exit 1
fi

echo "[1/3] Installazione dipendenze..."
pip3 install openpyxl pyinstaller sv-ttk python-docx
if [ $? -ne 0 ]; then
    echo "ERRORE durante l'installazione delle dipendenze."
    exit 1
fi

echo ""
echo "[2/3] Creazione applicazione .app..."
cd "$SCRIPT_DIR"

# Lettura vocale: inclusa solo se il modello vocale e' in codice/modelli
# (vedi ISTRUZIONI_MAC.txt). Senza modello l'app funziona, ma senza voce.
# ATTENZIONE: la parte vocale su macOS NON e' stata verificata (sviluppata
# senza un Mac): provare il pulsante "Voce" prima di distribuire l'app.
VOCE_OPTS=()
if ls "$CODICE"/modelli/vosk-model* >/dev/null 2>&1; then
    echo "Modello vocale trovato: includo la lettura vocale."
    if pip3 install vosk sounddevice pyttsx3; then
        VOCE_OPTS=(--add-data "$CODICE/modelli:modelli" \
                   --collect-all vosk --collect-all sounddevice --collect-all _sounddevice_data \
                   --hidden-import pyttsx3.drivers --hidden-import pyttsx3.drivers.nsss)
    else
        echo "ATTENZIONE: librerie vocali non installate, creo l'app senza voce."
    fi
else
    echo "Modello vocale non trovato in codice/modelli: creo l'app senza voce."
fi

pyinstaller --windowed \
  --add-data "$CODICE/Polline_Template_Settimanale.xlsx:." \
  --add-data "$CODICE/concentrazioni_polliniche.xlsx:." \
  --add-data "$CODICE/ITA_Template_Bollettino_pubblicazione.docx:." \
  --add-data "$CODICE/ENG_Template_Bollettino_pubblicazione.docx:." \
  --hidden-import dominio \
  --hidden-import sessione \
  --hidden-import esportatori \
  --hidden-import percorsi \
  --hidden-import voce \
  --hidden-import docx \
  --hidden-import sv_ttk \
  "${VOCE_OPTS[@]}" \
  --name "Conta_Pollinica" \
  "$CODICE/polline_counter_gui.py"
if [ $? -ne 0 ]; then
    echo "ERRORE durante la creazione dell'applicazione."
    exit 1
fi

echo ""
if [ ${#VOCE_OPTS[@]} -gt 0 ]; then
    # macOS chiude le app che usano il microfono senza questa descrizione.
    /usr/libexec/PlistBuddy -c "Add :NSMicrophoneUsageDescription string Conta Pollinica usa il microfono per la lettura vocale dei granuli." \
        dist/Conta_Pollinica.app/Contents/Info.plist
    # La modifica invalida la firma: ri-firma "ad hoc" (senza certificato Apple).
    codesign --force --deep -s - dist/Conta_Pollinica.app
fi

echo ""
echo "[3/3] Pulizia file temporanei..."
mv dist/Conta_Pollinica.app "$SCRIPT_DIR/Conta_Pollinica.app"
rm -rf dist build Conta_Pollinica.spec

echo ""
echo "============================================"
echo "  FATTO!"
echo "============================================"
echo ""
echo "L'applicazione si trova in:"
echo "  mac/Conta_Pollinica.app"
echo ""
echo "Per distribuirla, comprimi la cartella Conta_Pollinica.app"
echo "in un archivio .zip (clic destro > Comprimi)."
echo ""
echo "NOTA: alla prima apertura su altri Mac, fai clic destro"
echo "sull'icona e scegli 'Apri' per aggirare il blocco di"
echo "macOS Gatekeeper (richiesto per le app non firmate)."
echo ""
