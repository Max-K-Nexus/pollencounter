#!/usr/bin/env python3
"""
Risoluzione dei percorsi (cartella sorgente/bundle, cartella accanto
all'eseguibile, file di configurazione).

Isola in un solo punto la differenza fra esecuzione da sorgente e da
eseguibile PyInstaller (`sys.frozen`), cosi' CLI e GUI non la duplicano.
"""

import sys
from pathlib import Path


def _frozen():
    return getattr(sys, "frozen", False)


# Cartella che contiene i file inclusi nel bundle (template, soglie, .docx):
# in sorgente e' la cartella di questo file (codice/); da frozen e' la
# cartella temporanea di estrazione di PyInstaller (sys._MEIPASS).
BUNDLE_DIR = Path(sys._MEIPASS) if _frozen() else Path(__file__).parent

# Cartella dei sorgenti .py (coincide con BUNDLE_DIR sia in sorgente che frozen).
SCRIPT_DIR = BUNDLE_DIR

# Cartella accanto all'eseguibile (frozen) o allo script (sorgente): e' la
# cartella di output di default prima che l'utente ne scelga una propria,
# ed e' dove risiede pollencounter.cfg.
EXE_DIR = Path(sys.executable).parent if _frozen() else Path(__file__).parent

CONFIG_FILE = EXE_DIR / "pollencounter.cfg"
