# Conta Pollinica — Guida per Claude Code

## Scopo del progetto

Sistema di conta pollinica settimanale usato in aerobiologia e sanità pubblica.
Gli operatori contano i granuli di polline da vetrini campionatori e inseriscono
un codice numerico per ogni granulo osservato. Il software registra i dati in un
file Excel strutturato e genera un bollettino pollinico con livelli di
concentrazione (assente / bassa / media / alta).

**Popolazione utente:** specializzandi, dottorandi e docenti di biologia senza
esperienza di programmazione. L'interfaccia è in italiano. Le istruzioni nei
messaggi di errore devono essere chiare e non presupporre conoscenze informatiche.

---

## Struttura dei file

```
pollencounter/
  codice/                             ← script principali e file di riferimento del codice
    polline_counter.py                ← script CLI, cross-platform (Linux + Windows)
    polline_counter_gui.py            ← GUI tkinter, cross-platform
    Polline_Template_Settimanale.xlsx ← template Excel master (NON modificare)
    concentrazioni_polliniche.xlsx    ← soglie per il bollettino
    pollencounter.cfg                 ← configurazione cartella di lavoro per anno
  script_aiuto/                       ← avviatori e utility di manutenzione
    AVVIA_CONTA_POLLINICA_GUI.sh      ← avvio GUI da terminale (Linux)
    AVVIA_CONTA_POLLINICA.sh          ← avvio CLI da terminale (Linux)
    applica_formattazione.py          ← utility manutenzione template (eseguire manualmente)
  letture_settimanali/                ← file .xlsx di output delle sessioni reali
  riferimenti/                        ← dati storici e file di riferimento
    Tabelle monitoraggi settimanali_2022_2023.xlsx
    Tabelle monitoraggi ultima settimana gennaio.xlsx
  windows/                            ← file specifici per Windows
    AVVIA_CONTA_POLLINICA.bat         ← avvio GUI su Windows (con Python)
    build_exe.bat                     ← crea Conta_Pollinica.exe con PyInstaller
    Conta_Pollinica.exe               ← eseguibile precompilato
    ISTRUZIONI_WINDOWS.txt            ← istruzioni per utenti Windows
  CLAUDE.md                           ← questo file (guida per Claude Code)
  CHANGELOG.md                        ← log delle modifiche (vedi sezione dedicata)
  ISTRUZIONI.txt                      ← istruzioni d'uso
  DISTRIBUZIONE_OPZIONI.md            ← opzioni di distribuzione centralizzata
  PROMPT_WEBAPP_CLAUDE.md
  esempio di bolletino.pdf
```

---

## Architettura

### `polline_counter.py` (CLI)

Il flusso principale è in `main()`:

1. Cerca file `.xlsx` esistenti in `OUTPUT_DIR` (`cerca_file_esistenti`)
2. Mostra menu: riprendi file / nuovo / importa da altra cartella (`chiedi_ripresa_o_nuovo`)
3. Chiede la settimana di riferimento e il giorno di lavoro
4. Loop di inserimento (`sessione_giorno`): l'operatore digita codici 01–59
5. Al termine: menu uscita con scelta della cartella di salvataggio

**Costanti chiave:**
- `AUTOSAVE_INTERVAL = 5` — autosave ogni N inserimenti in `~autosave_*.xlsx`
- `CODICI_SPECIE` — dict `"01"–"59"` → nome specie (01–47 pollini, 48–59 spore)
- `SOGLIE_MAPPING` — dict codice → nome famiglia nel file soglie
- `OUTPUT_DIR = Path(__file__).parent` quando non frozen; `Path(sys.executable).parent` quando frozen (PyInstaller)

**Funzioni principali:**

| Funzione | Ruolo |
|----------|-------|
| `sessione_giorno()` | Loop di inserimento per un giorno; ritorna `"continue"` o `"quit"` |
| `menu_uscita_salvataggio()` | Menu fine sessione: salva nuovo / sovrascrivi / esci senza salvare |
| `chiedi_cartella_salvataggio()` | Chiede la cartella; stampa `__GUI_ASKDIR__` (intercettato dalla GUI) |
| `chiedi_ripresa_o_nuovo()` | Menu iniziale; opzione `i` stampa `__GUI_ASKOPENFILE__` |
| `autosave()` | Salva in `~autosave_{lunedi_str}.xlsx` silenziosamente |
| `pulisci_file_temporanei()` | Post-sessione: cancella o rinomina l'autosave in `incompleto_` |
| `genera_bollettino()` | Crea il foglio "bollettino" con colori concentrazione |
| `carica_soglie()` | Legge `concentrazioni_polliniche.xlsx` |

**SIGTERM handler** (solo Linux): registrato in `main()` dopo che `wb` e
`lunedi_str` sono noti. Chiama `autosave()` e `sys.exit(0)`. Permette alla GUI
di chiudere il processo con `process.terminate()` salvando i dati.

### `polline_counter_gui.py` (GUI)

Finestra tkinter divisa in due pannelli:
- **Sinistra:** terminale emulato (widget `Text` + `Entry`)
- **Destra:** notebook con tre tab live (Settimanale, Giornaliero, Bollettino)

Il tab Bollettino usa `carica_soglie()` e colori `_BOLL_COLORS` per visualizzare
i livelli di concentrazione in tempo reale.

**Subprocess:**
- **Linux:** pty (`openpty`) — il processo vede un terminale reale, ANSI e input interattivo funzionano nativamente
- **Windows:** `subprocess.Popen` con `stdin=PIPE, stdout=PIPE` + thread lettore + `queue.Queue`

**Protocollo marker GUI** (pattern critico):

Lo script stampa tag speciali che la GUI intercetta in `_handle_gui_markers()`,
li rimuove dal testo visualizzato e apre un dialogo nativo:

| Marker nello stdout dello script | Dialogo aperto dalla GUI |
|----------------------------------|--------------------------|
| `__GUI_ASKDIR__` | `filedialog.askdirectory()` — scegli cartella di salvataggio |
| `__GUI_ASKOPENFILE__` | `filedialog.askopenfilename()` — scegli file da importare |

La GUI invia il percorso scelto (o stringa vuota se annullato) via stdin.
Lo script riceve il percorso come risposta alla `input()` successiva.
In modalità CLI i marker sono visibili ma innocui.

**Tracking del file corrente** (`_detect_tracked_file`):
La GUI legge i path completi stampati dallo script tramite regex:
- `Ripreso: (.+\.xlsx)` — file ripreso
- `File salvato: (.+\.xlsx)` — salvataggio definitivo
- `[auto-salvato]` — cerca il più recente `~autosave_*.xlsx` in `SCRIPT_DIR`

Queste stampe usano path completi (non `.name`) appositamente per questo tracking.

**Font:** `_MONO_FONT = "Courier New"` su Windows, `"Monospace"` su Linux.

**`sv_ttk`:** tema opzionale (`try/except`). Attivato solo su Windows in `main()`.

---

## Convenzioni da rispettare

- **Lingua:** tutto il testo mostrato all'utente è in italiano.
- **Stampe path completi:** `print(f"... {path}")` non `{path.name}` — serve al tracking della GUI.
- **Marker GUI:** se si aggiunge una nuova `input()` che nella GUI dovrebbe aprire un dialogo, seguire il pattern `print("__GUI_MARKER__", flush=True)` + gestione in `_handle_gui_markers()`. Il `flush=True` è obbligatorio per Windows (pipe bufferizzate).
- **Caratteri ASCII only nelle stampe:** non usare caratteri Unicode fuori cp1252 (es. `✓`, `─`, emoji) nelle stringhe stampate a stdout. Windows con cp1252 va in crash. Usare alternative ASCII (es. `[OK]` al posto di `✓`).
- **Template Excel:** non modificare la struttura dei fogli `riepilogo_settimana` e `dati_grezzi`. Le righe sono fisse: pollini in righe 6–52, spore in 58–69 (funzioni `codice_to_row`, `giorno_to_col`).
- **`OUTPUT_DIR` vs `SCRIPT_DIR`:** i file di output (xlsx) vanno in `OUTPUT_DIR`; il template e le soglie si cercano anche in `BUNDLE_DIR` (frozen) e `SCRIPT_DIR`. Entrambi gli script risiedono in `codice/` e usano `Path(__file__).parent` — nessun percorso hardcoded.
- **Autosave:** il file `~autosave_*.xlsx` viene trovato da `cerca_file_esistenti()` al prossimo avvio ed è presentato all'utente come file recuperabile.
- **Cartella `windows/` non autocontenuta:** contiene solo i `.bat`, l'exe e le istruzioni. I sorgenti `.py` e i file `.xlsx` risiedono in `codice/`; `build_exe.bat` e `AVVIA_CONTA_POLLINICA.bat` li referenziano con path `..\codice\`. Non copiare i sorgenti in `windows/`.

---

## Dipendenze

```
openpyxl      ← obbligatorio (lettura/scrittura Excel)
tkinter       ← obbligatorio per la GUI (su Debian: sudo apt install python3-tk)
sv_ttk        ← opzionale (tema Windows; pip install sv-ttk)
pyinstaller   ← solo per build Windows exe (vedi sezione sotto)
winsound      ← solo Windows, incluso nella stdlib
```

---

## Build dell'eseguibile Windows (.exe)

L'exe si compila su **Linux via Wine**, con Python 3.11 Windows già installato
nel prefix Wine di root. NON usare PyInstaller Linux nativo (produrrebbe un
eseguibile ELF, non un .exe).

### Prerequisiti (già presenti sul sistema)

- **Wine:** `/usr/bin/wine` (wine-10.0)
- **Python Windows 3.11:** `/root/.wine/drive_c/users/root/AppData/Local/Programs/Python/Python311/python.exe`
- **PyInstaller Windows:** `/root/.wine/drive_c/.../Python311/Scripts/pyinstaller.exe`

Verificare con:
```bash
wine python --version        # deve stampare Python 3.11.x
wine python -m PyInstaller --version
```

### Comando di build

Eseguire dalla directory `windows/`:

```bash
cd /home/Simone/Documenti/Spec_Igiene/pollencounter/windows

# Aggiorna dipendenze Python Windows (solo se necessario)
WINEDEBUG=-all wine python -m pip install --quiet openpyxl sv-ttk

# Compila l'exe (i file sorgente e xlsx sono in codice/)
WINEDEBUG=-all wine python -m PyInstaller --onefile --windowed \
  --add-data "../codice/Polline_Template_Settimanale.xlsx;." \
  --add-data "../codice/concentrazioni_polliniche.xlsx;." \
  --hidden-import polline_counter \
  --hidden-import sv_ttk \
  --name "Conta_Pollinica" \
  ../codice/polline_counter_gui.py

# Sposta e pulisci
mv dist/Conta_Pollinica.exe ./Conta_Pollinica.exe
rm -rf dist build Conta_Pollinica.spec
ls -lh Conta_Pollinica.exe   # verifica ~12MB
```

**Note critiche:**
- Il separatore in `--add-data` è `;` (stile Windows), non `:` (Linux).
- `WINEDEBUG=-all` sopprime i messaggi di debug di Wine (molto verbosi altrimenti).
- Il build richiede ~2-3 minuti.
- Se PyInstaller non è trovato: `WINEDEBUG=-all wine python -m pip install pyinstaller`
- **Non** usare `pip3 install pyinstaller` né `pipx install pyinstaller`:
  producono eseguibili Linux, non Windows.

### Flusso completo post-modifica

```bash
# 1. Compila (i sorgenti sono in root, windows/ contiene solo i bat e l'exe)
cd windows && WINEDEBUG=-all wine python -m PyInstaller ...

# 2. Verifica con Wine
WINEDEBUG=-all wine Conta_Pollinica.exe
```

---

## Changelog (`CHANGELOG.md`) — OBBLIGATORIO

Il file `CHANGELOG.md` nella root del progetto contiene il log cronologico
di tutte le modifiche apportate al codice.

**REGOLA: aggiornare `CHANGELOG.md` al termine di OGNI modifica, non solo
a fine sessione.** Ogni singolo task che modifica codice o file di progetto
deve avere la sua entry nel changelog prima di passare al task successivo.

**Regole di compilazione per Claude Code:**

1. **Quando aggiornare:** subito dopo aver completato ogni singola modifica
   funzionale. NON aspettare la fine della sessione: aggiornare il changelog
   e' l'ultimo passo di ogni task.
2. **Formato:** aggiungere una nuova sezione con data (`## YYYY-MM-DD`),
   titolo del fix/feature (`### Titolo`), e i campi **Problema**,
   **Causa** e **Correzione** con i file e le righe coinvolte.
3. **Dove aggiungere:** in cima al file, subito dopo l'intestazione, in modo
   che la modifica piu' recente sia sempre la prima visibile.
4. **Cosa includere:** ogni modifica funzionale (bug fix, nuova feature,
   cambiamento di comportamento). Non includere refactoring cosmetici o
   modifiche ai soli commenti.

---

## Utility di manutenzione template

### `applica_formattazione.py`

Script standalone da eseguire **manualmente** quando si vuole aggiornare la
formattazione visiva del template Excel. Non e' importato ne' chiamato dagli
script principali.

**Eseguire con:**
```bash
python3 script_aiuto/applica_formattazione.py
```

**Cosa fa:**
- Imposta altezze righe (15pt per righe 5–53 e 57–70)
- Applica sfondo verde tenue (`C5E0B4`) alle righe delle specie principali
- Applica bordi thin completi sulle sezioni dati
- Applica grassetto selettivo su nomi specie importanti e intestazioni
- Centra il contenuto delle celle dati
- Aggiorna in parallelo le tre copie del template (root, `linux/`, `windows/`)
- Crea automaticamente un backup prima di modificare il template principale

**Nota:** non inserisce formule Excel. Opera esclusivamente sulla formattazione.

---

## Note operative per sessioni future

- Prima di modificare qualsiasi funzione, leggere il file per intero: i due script sono grandi (~1400 e ~670 righe) ma strettamente accoppiati.
- Il codice è usato in produzione da utenti non tecnici: privilegiare stabilità e messaggi chiari rispetto a refactoring.
- Le modifiche al protocollo marker o alle stampe dei path rompono il tracking della GUI: verificare sempre entrambi i file insieme.
- I sorgenti `.py` sono unici (in `codice/`). Non creare copie in `windows/`.
- **Dopo ogni modifica:** aggiornare `CHANGELOG.md` (vedi sezione sopra). Questo e' obbligatorio.
- **Dopo ogni spostamento di script o cambio di percorsi:** aggiornare il launcher Desktop
  (`/home/Simone/Scrivania/ContaPollinica.desktop`), campo `Exec=`. Percorso attuale:
  `bash .../pollencounter/script_aiuto/AVVIA_CONTA_POLLINICA_GUI.sh`.
- Per testare: `python3 codice/polline_counter.py` (CLI) e `python3 codice/polline_counter_gui.py` (GUI) dalla directory `pollencounter/`, oppure direttamente dalla cartella `codice/`.
- **Per ricompilare l'exe:** usare Wine + Python Windows (vedi sezione "Build dell'eseguibile Windows"). Non usare PyInstaller Linux nativo.
