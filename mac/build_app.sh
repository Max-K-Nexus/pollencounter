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
pip3 install openpyxl pyinstaller sv-ttk
if [ $? -ne 0 ]; then
    echo "ERRORE durante l'installazione delle dipendenze."
    exit 1
fi

echo ""
echo "[2/3] Creazione applicazione .app..."
cd "$SCRIPT_DIR"
pyinstaller --windowed \
  --add-data "$CODICE/Polline_Template_Settimanale.xlsx:." \
  --add-data "$CODICE/concentrazioni_polliniche.xlsx:." \
  --hidden-import polline_counter \
  --hidden-import sv_ttk \
  --name "Conta_Pollinica" \
  "$CODICE/polline_counter_gui.py"
if [ $? -ne 0 ]; then
    echo "ERRORE durante la creazione dell'applicazione."
    exit 1
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
