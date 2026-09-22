---
name: build-windows-exe
description: Compila Conta_Pollinica.exe per Windows usando Wine e PyInstaller Windows. Usare quando l'utente chiede di ricompilare, aggiornare o generare l'eseguibile Windows del Conta Pollinica.
---

# Build dell'eseguibile Windows (.exe)

L'exe si compila su **Linux via Wine**, con Python 3.11 Windows installato nel
prefix Wine dell'utente corrente. NON usare PyInstaller Linux nativo
(produrrebbe un eseguibile ELF, non un .exe).

## Prerequisiti

- **Wine:** `/usr/bin/wine` (su openSUSE: `sudo zypper install wine`, oppure il
  repository OBS `Emulators:Wine` se la versione nei repo base e' troppo vecchia)
- **Python Windows 3.11:** installato nel prefix Wine dell'utente
  (`~/.wine/drive_c/users/$USER/AppData/Local/Programs/Python/Python311/python.exe`),
  scaricando l'installer ufficiale (`python-3.11.9-amd64.exe` da python.org) e
  lanciandolo con `wine python-3.11.9-amd64.exe /quiet InstallAllUsers=0 PrependPath=1`
- **PyInstaller Windows:** `wine python -m pip install pyinstaller`

Verificare con:
```bash
wine python --version        # deve stampare Python 3.11.x
wine python -m PyInstaller --version
```

## Comando di build

Eseguire dalla directory `windows/` del repository
(`/home/sbettella/Progetti_Claude/pollencounter/windows`):

```bash
cd /home/sbettella/Progetti_Claude/pollencounter/windows

# Aggiorna dipendenze Python Windows (solo se necessario)
WINEDEBUG=-all wine python -m pip install --quiet openpyxl sv-ttk python-docx pyinstaller

# Compila l'exe (i sorgenti .py, il template .xlsx, il file soglie e i due
# template .docx dei bollettini sono tutti in codice/ e vanno tutti inclusi:
# dimenticare i .docx fa fallire silenziosamente la generazione dei bollettini
# nell'exe compilato)
WINEDEBUG=-all wine python -m PyInstaller --onefile --windowed \
  --add-data "../codice/Polline_Template_Settimanale.xlsx;." \
  --add-data "../codice/concentrazioni_polliniche.xlsx;." \
  --add-data "../codice/ITA_Template_Bollettino_pubblicazione.docx;." \
  --add-data "../codice/ENG_Template_Bollettino_pubblicazione.docx;." \
  --hidden-import dominio \
  --hidden-import sessione \
  --hidden-import esportatori \
  --hidden-import percorsi \
  --hidden-import docx \
  --hidden-import sv_ttk \
  --name "Conta_Pollinica" \
  ../codice/polline_counter_gui.py

# Sposta e pulisci
mv dist/Conta_Pollinica.exe ./Conta_Pollinica.exe
rm -rf dist build Conta_Pollinica.spec
ls -lh Conta_Pollinica.exe   # verifica ~15-20MB (piu' grande di prima: include python-docx)
```

**Note critiche:**
- Il separatore in `--add-data` è `;` (stile Windows), non `:` (Linux).
- `WINEDEBUG=-all` sopprime i messaggi di debug di Wine (molto verbosi altrimenti).
- Il build richiede ~2-3 minuti.
- Se PyInstaller non è trovato: `WINEDEBUG=-all wine python -m pip install pyinstaller`
- **Non** usare `pip3 install pyinstaller` né `pipx install pyinstaller`:
  producono eseguibili Linux, non Windows.
- L'exe è costruito solo da `polline_counter_gui.py` (GUI): dalla riscrittura
  del guscio la GUI non lancia più `polline_counter.py` come sottoprocesso, e
  l'eseguibile non ha più una modalità `--cli` di rilancio interno. Per usare
  la CLI su Windows serve un'installazione Python separata (vedi
  `windows/ISTRUZIONI_WINDOWS.txt`).
- `windows/Conta_Pollinica.exe` non è più tracciato in git (vedi `.gitignore`):
  ogni build resta solo locale, non finisce nella storia del repository.

## Flusso completo post-modifica

```bash
# 1. Compila (i sorgenti sono in codice/, windows/ contiene solo i .bat e l'exe)
cd windows && WINEDEBUG=-all wine python -m PyInstaller ...

# 2. Verifica con Wine
WINEDEBUG=-all wine Conta_Pollinica.exe
```
