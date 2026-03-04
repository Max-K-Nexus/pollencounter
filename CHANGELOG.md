# Changelog — Conta Pollinica

Log delle modifiche apportate al progetto, compilato al termine di ogni task.

---

## 2026-03-03

### Riorganizzazione struttura cartelle del progetto

**Motivazione:** separare il codice sorgente, gli script di utilità, gli output
e i dati storici in cartelle dedicate per maggiore chiarezza.

**Nuova struttura:**
- `codice/` — script Python principali (`polline_counter.py`, `polline_counter_gui.py`),
  template Excel, soglie concentrazioni, `pollencounter.cfg`
- `script_aiuto/` — script di avvio Linux (`.sh`) e `applica_formattazione.py`
- `letture_settimanali/` — invariata (output sessioni)
- `riferimenti/` — file Excel storici di riferimento
- `windows/` — invariata (`.bat`, `.exe`)

**File modificati per aggiornare i percorsi:**
- `script_aiuto/applica_formattazione.py`: `BASE_DIR` aggiornato a `../codice/`
- `script_aiuto/AVVIA_CONTA_POLLINICA_GUI.sh`: percorso GUI → `../codice/`
- `script_aiuto/AVVIA_CONTA_POLLINICA.sh`: percorso script → `../codice/`
- `windows/AVVIA_CONTA_POLLINICA.bat`: percorso GUI → `..\codice\`
- `windows/build_exe.bat`: tutti i percorsi sorgente → `..\codice\`
- `CLAUDE.md`: struttura e comandi di build aggiornati
- `ISTRUZIONI.txt`: comandi di avvio aggiornati

**Note:** `polline_counter.py` e `polline_counter_gui.py` non richiedono modifiche
(usano `Path(__file__).parent` e si adattano automaticamente alla nuova posizione).

---

## 2026-03-03

### Ripristino cfg stabile: rimossa auto-aggiornamento cartella a fine sessione

**Problema:** la modifica precedente aggiornava automaticamente il cfg con la
cartella usata per l'ultimo salvataggio, sovrascrivendo la cartella di default
dell'anno ad ogni sessione.

**Correzione:** rimossa `_aggiorna_config()` e la relativa chiamata in `main()`.
Il cfg ora rimane stabile sulla cartella di default configurata dall'utente
(una volta per anno). Se lo script lavora con file da un'altra cartella
(tramite importa / chiedi_cartella_salvataggio), questo avviene senza toccare
il cfg. File: `polline_counter.py`.

---

## 2026-03-03

### Fix lag input e aggiornamento live tab GUI

**Problema 1 — Lag durante l'inserimento pollini:**
`autosave()` chiamava `wb.save()` nel thread principale del subprocess,
bloccando stdin per 0.5-2s ogni 5 inserimenti. L'utente vedeva il cursore
"congelato" e i codici digitati apparivano tutti insieme dopo il save.

**Causa:** serializzazione xlsx (XML + ZIP) sincrona nel main thread.

**Correzione:** `autosave()` ora lancia `wb.save()` in un thread daemon
(`_do_save`). Aggiunto flag `_autosave_running` per saltare se il save
precedente e' ancora in corso (evita accumulo). Aggiunto `_attendi_autosave()`
chiamato prima di ogni salvataggio definitivo in `menu_uscita_salvataggio()`
(`polline_counter.py`, funzione `autosave` e `menu_uscita_salvataggio`).

**Problema 2 — Tab GUI non aggiornati durante la sessione:**
- La GUI tracciava il file importato (non cambia in memoria) invece
  dell'autosave live. Il primo aggiornamento arrivava solo al 5° inserimento.
- `_refresh_running` veniva resettato a `False` all'inizio di `_refresh_summary`
  prima che il thread finisse: se arrivava un `[auto-salvato]` durante la
  lettura xlsx, veniva schedulato un secondo timer concorrente.

**Correzione:**
- `sessione_giorno()` ora fa un autosave immediato all'avvio (prima del loop),
  cosi' la GUI riceve `[auto-salvato]:` subito dopo `Giorno:` e inizia a
  leggere il file live da zero inserimenti (`polline_counter.py`).
- `_refresh_running` non viene piu' resettato all'inizio di `_refresh_summary`:
  rimane `True` per tutta la durata del timer + thread, viene azzerato solo
  in `_schedula_prossimo_refresh` quando la sessione termina
  (`polline_counter_gui.py`).
- `_applica_dati` usa `tree.delete(*children)` (singola chiamata Tcl) invece
  di un loop per ogni riga, riducendo il tempo nel main thread
  (`polline_counter_gui.py`).

---

## 2026-03-03

### cfg aggiornato a fine sessione con la cartella effettivamente usata

**Problema:** se l'utente salvava il file in una cartella diversa da quella
configurata nel cfg, alla sessione successiva `cerca_file_esistenti()` cercava
nell'old `OUTPUT_DIR` e non trovava i file appena creati.

**Causa:** `carica_o_crea_config()` scriveva il cfg solo alla prima configurazione
dell'anno; i salvataggi successivi in cartelle diverse non aggiornavano il cfg.

**Correzione:** aggiunto `_aggiorna_config(cartella)` (`polline_counter.py`,
prima di `carica_o_crea_config`). Chiamato in `main()` dopo ogni salvataggio
riuscito con `cartella_salvata`, in modo che il cfg rifletta sempre l'ultima
cartella usata e la prossima sessione parta da lì.

---

## 2026-03-03

### Fix refresh live GUI: autosave non trovato quando OUTPUT_DIR != SCRIPT_DIR

**Problema:** il pannello di riepilogo live nella GUI non si aggiornava durante
l'inserimento dati. Il refresh veniva schedulato solo quando la GUI riusciva a
individuare il file autosave, ma la ricerca usava `SCRIPT_DIR.glob("~autosave_*.xlsx")`
(sempre la cartella root dello script), mentre l'autosave veniva scritto in
`OUTPUT_DIR` (la cartella configurata da `pollencounter.cfg`). Se le due
cartelle erano diverse il file non veniva mai trovato e `_tracked_file` restava None.

**Causa:** introdotto con la feature `pollencounter.cfg` (2026-02-27): prima
`OUTPUT_DIR` coincideva sempre con `SCRIPT_DIR`, dopo poteva divergere.

**Correzione:**
- `polline_counter.py` (riga ~2081): cambiato `print("  [auto-salvato]")` in
  `print(f"  [auto-salvato]: {OUTPUT_DIR / f'~autosave_{lunedi_str}.xlsx'}")`.
  Il path completo del file e' ora incluso nel messaggio.
- `polline_counter_gui.py` (`_detect_tracked_file`): sostituita la ricerca glob
  in `SCRIPT_DIR` con parsing regex del path direttamente dal messaggio
  `[auto-salvato]: /percorso/file.xlsx`. Se il file esiste, viene impostato come
  `_tracked_file` e il refresh parte.
- `pollencounter.cfg`: aggiornato il path 2026 da `linux/letture_settimanali`
  (non piu' esistente) a `letture_settimanali` (root del progetto).

---

### Riorganizzazione struttura directory

**Motivazione:** ridurre la ridondanza delle copie di sorgenti e dati presenti
in piu' cartelle (`linux/`, `windows/`, root). I file `.py` e `.xlsx` avevano
tre copie che andavano sincronizzate manualmente dopo ogni modifica.

**Correzione:**

- **Sorgenti unici in root:** eliminati `windows/polline_counter.py`,
  `windows/polline_counter_gui.py`, `windows/concentrazioni_polliniche.xlsx`,
  `windows/Polline_Template_Settimanale.xlsx` (tutti identici al master in root).
  La cartella `windows/` ora contiene solo i `.bat`, l'exe e le istruzioni.

- **`build_exe.bat` e `AVVIA_CONTA_POLLINICA.bat`** aggiornati con path `..`
  per referenziare i sorgenti nella root padre.

- **Script shell spostati in root:** `AVVIA_CONTA_POLLINICA_GUI.sh` e
  `AVVIA_CONTA_POLLINICA.sh` spostati da `linux/` alla root e corretti
  (il CLI puntava erroneamente a `$SCRIPT_DIR/polline_counter.py`
  che non esisteva in `linux/`).

- **`linux/` svuotata:** rimossi `.sh`, `concentrazioni_polliniche.xlsx`,
  `Polline_Template_Settimanale.xlsx` (versione obsoleta, 15K, priva del
  foglio "soglie"). Restano solo file root-owned (`__pycache__`, `.claude/`)
  eliminabili con `sudo rm -rf linux/`.

- **`letture_settimanali/` creata in root:** spostati i file di output
  di sessioni reali (`Conta_Pollinica_16-02-2026.xlsx`,
  `Conta_Pollinica_23-02-2026.xlsx`) che erano dispersi in `linux/`.

- **Duplicati eliminati in root:** `concentrazioni_polliniche (1).xlsx`
  (duplicato accidentale) e `Polline_Template_Settimanale_BACKUP.xlsx`.

- **`ISTRUZIONI.txt`** spostato da `linux/` a root; corretto riferimento
  a `aggiorna_template.py` (eliminato) con `applica_formattazione.py`.

- **`CLAUDE.md`** aggiornato: struttura file, comando build Wine con path
  `../`, convenzioni, note operative.

---

## 2026-03-02

### Rebuild eseguibile Windows (Conta_Pollinica.exe)

**Motivazione:** l'exe precedente (build 26 feb ore 20:26) non includeva le
modifiche funzionali del 27 feb (riga "G. anno", foglio W##, config
`pollencounter.cfg`).

**Correzione:** ricompilato con Wine + Python Windows 3.11 + PyInstaller
`--onefile --windowed`. Nuovo exe: 12 MB, timestamp 2 mar.

---

### Uniformazione file Python e pulizia utility template

**Problema:**
1. `windows/polline_counter.py` era rimasto indietro di una modifica rispetto al
   master (3 righe di commento divergenti alle righe 157–160, nessuna differenza
   funzionale).
2. Coesistevano due script di manutenzione template (`linux/aggiorna_template.py`
   e `applica_formattazione.py`) con logiche parzialmente sovrapposte e
   incompatibili (colori diversi, il primo inseriva formule AVERAGE non necessarie,
   il secondo aggiornava tutte e tre le copie del template mentre il primo solo
   `linux/`).

**Causa:** la copia Windows non era stata aggiornata dopo l'ultima modifica del
27-02 ore 17:54. Lo script `linux/aggiorna_template.py` era uno script più
vecchio mai consolidato nel workflow ufficiale.

**Correzione:**
- `windows/polline_counter.py`: aggiornato copiando il master (ora identici
  byte-per-byte).
- `linux/aggiorna_template.py`: eliminato. La utility canonica di manutenzione
  template è `applica_formattazione.py` in root.
- `CLAUDE.md`: rimosso riferimento a `aggiorna_template.py` dalla struttura file;
  aggiunta sezione "Utility di manutenzione template" che documenta
  `applica_formattazione.py`.

---

## 2026-02-27

### Riga "giorno dell'anno" (G. anno) nei fogli annuali

**Problema:** le intestazioni a data nel Calendario e nei fogli W## mostravano
solo la data nel formato dd/mm, senza indicare a quale giorno dell'anno
(1-366) corrispondesse.

**Correzione:**
- `polline_counter.py`:
  - Costanti: aggiunto `_CAL_DOY_ROW = 4`; `_CAL_DATA_START_ROW` spostato
    da 4 a 5; `_CAL_SEP_ROW` aggiornato di conseguenza (52).
  - `crea_intestazione_calendario()`: aggiunta riga 4 con label "G. anno"
    in colonna A; `freeze_panes` aggiornato a "B5".
  - `scrivi_colonna_calendario()`: dopo l'intestazione data, scrive il
    numero di giorno dell'anno nella riga `_CAL_DOY_ROW` per ogni colonna.
  - `crea_foglio_settimana_annuale()`: aggiunta funzione interna
    `_scrivi_doy()` che scrive la riga "G. anno" (giorno 1-366) subito
    sotto l'intestazione in entrambe le sezioni (CONTA GREZZA e
    CONCENTRAZIONI), con colore coerente (giallo / blu).
- Copiato `windows/polline_counter.py`.

### Foglio settimanale nel riepilogo annuale (W##)

**Problema:** il riepilogo annuale non aveva una vista organizzata per settimana;
era possibile vedere i dati solo per giorno (foglio Dati) o per specie (Calendario).

**Correzione:**
- `polline_counter.py`: aggiunte funzioni `_nome_foglio_settimana()`,
  `_posizione_foglio_settimana()`, `crea_foglio_settimana_annuale()`.
  Il foglio e' denominato `W##` (numero ISO settimana, es. `W05`), creato o
  sovrascritto ad ogni esportazione. Layout: due sezioni verticali —
  "CONTA GREZZA" (Specie x 7 giorni + Totale settimana) e
  "CONCENTRAZIONI p/m3" (stessa struttura + Media settimana).
  Pollini e spore separati da una riga separatore. Colori e stili
  coerenti con il resto del file annuale (verde/verde chiaro per le
  specie monitorate, giallo per gli header grezzi, blu per le conc.).
  I fogli W## sono inseriti in ordine crescente dopo i fogli fissi
  ("Dati Anno", "Calendario").
- `esporta_riepilogo_annuale()`: chiama `crea_foglio_settimana_annuale()`
  dopo il loop dei giorni, prima del salvataggio.
- Copiato `windows/polline_counter.py`.

### Configurazione cartella di lavoro per anno (pollencounter.cfg)

**Problema:** `OUTPUT_DIR` era fisso alla cartella dello script. Se l'operatore
salvava i file settimanali in cartelle diverse, il riepilogo annuale si
frammentava in piu' posizioni e i file non venivano trovati al riavvio.

**Causa:** Assenza di memoria persistente della cartella di lavoro.

**Correzione:**
- `polline_counter.py` (righe 14-168, 1869-1928): aggiunto `import json`,
  costante `CONFIG_FILE` (`pollencounter.cfg` accanto allo script/exe),
  funzione `carica_o_crea_config()` che legge/crea il cfg per anno e
  all'avvio di `main()` riassegna `OUTPUT_DIR` (con `global`) prima di
  qualsiasi operazione su file.
- Il file cfg e' in JSON, una chiave per anno (es. `"2026": "/path/..."`)
  cosi' cambia automaticamente al cambio anno.
- Supporto marker GUI `__GUI_ASKDIR__` per la scelta della cartella.
- Copiato `windows/polline_counter.py`.

---

## 2026-02-26

### Riepilogo annuale: layout e foglio Calendario

**Feature:**
- Formato data nel riepilogo annuale cambiato da "gio 01/gen/2026" a "23/02/2026"
  (più compatto, applicato a tutte le iterazioni del codice).
- Foglio "Dati {anno}": larghezze colonne ottimizzate (Data=11, dati=4, sep=2,
  gap=3), altezza riga intestazioni=80pt (testo ruotato visibile), blocco
  intestazioni (`freeze_panes="B4"`), auto-filter sulla riga 3.
- Nuovo foglio "Calendario" nello stesso file: layout trasposto con specie in
  righe (pollini 01-47 + separatore + spore 48-59) e date in colonne.
  Mostra solo concentrazioni (p/m3). Larghezza colonna specie=26, colonne
  data=5, `freeze_panes="B4"`, altezza riga 3=60pt.
- `esporta_riepilogo_annuale()` aggiorna entrambi i fogli sincronizzando la
  stessa scelta duplicati (sovrascrivi/aggiungi/somma).

**Modifiche** (`polline_counter.py`, `windows/polline_counter.py`):
- Aggiunto import `get_column_letter` da `openpyxl.utils`.
- Rimosse costanti `GIORNI_ABBREV_ANN` e `MESI_ABBREV_ANN` (non più usate).
- Aggiunte costanti `_CAL_*` per il layout del foglio Calendario.
- `formatta_data_annuale()`: semplificato a `dt.strftime("%d/%m/%Y")`.
- `crea_intestazione_annuale()`: aggiunta formattazione larghezze/freeze/filter.
- Nuove funzioni: `_cal_row_for_codice()`, `crea_intestazione_calendario()`,
  `trova_colonna_per_data_calendario()`, `_prossima_colonna_calendario()`,
  `scrivi_colonna_calendario()`.
- Foglio rinominato da "Riepilogo {anno}" a "Dati {anno}".
- Nota: l'exe Windows va ricompilato su Windows con `build_exe.bat`.

---

### Riepilogo annuale Excel

**Feature:** dopo il salvataggio del file settimanale, l'utente puo' esportare
i dati in un file `Riepilogo_Annuale_{anno}.xlsx` nella stessa cartella. Il file
raccoglie i dati giornalieri di tutte le settimane: conta grezza e concentrazioni
(valore × fattore).

**Modifiche** (`polline_counter.py`, `windows/polline_counter.py`):
- Aggiunte costanti: `GIORNI_ABBREV_ANN`, `MESI_ABBREV_ANN`, `POLLINI_CODICI`,
  `SPORE_CODICI`, set di formattazione (`ANNUALE_VERDE_CODICI`, ecc.), layout
  colonne riepilogo annuale.
- Nuove funzioni: `formatta_data_annuale()`, `raccogli_dati_giornalieri()`,
  `crea_intestazione_annuale()`, `trova_riga_per_data()`, `scrivi_riga_annuale()`,
  `esporta_riepilogo_annuale()` e helper `_ann_col_grezzo()`, `_ann_col_conc()`,
  `_prossima_riga_annuale()`.
- `menu_uscita_salvataggio()`: ritorna `(bool, Path|None)` con la cartella di
  salvataggio invece del solo bool.
- `sessione_giorno()` e `main()`: aggiornati per gestire la tupla e proporre
  l'esportazione annuale dopo il salvataggio (`"Aggiornare il riepilogo annuale?"`).
- Gestione duplicati: se un giorno e' gia' presente nel file annuale, chiede
  all'utente se sovrascrivere, aggiungere riga o sommare (scelta valida per
  tutta la sessione).
- Formattazione: verde (FF92D050) per specie chiave, verde chiaro (FFC5E0B4) per
  Alternaria, intestazioni gialle (FFE699), bordi thin, formato "0.0" per
  concentrazioni, testo intestazioni ruotato 90 gradi.

---

### Ricompilazione exe con template formattato e soglie

**Problema:** l'exe Windows usava un template `Polline_Template_Settimanale.xlsx`
privo della formattazione visiva (colori intestazioni gialli, fill rosso, ecc.)
e del foglio "soglie" integrato. 99 celle avevano fill diverso rispetto al
template root.

**Correzione:**
- Copiato il template root (con formattazione completa e foglio "soglie") in
  `windows/Polline_Template_Settimanale.xlsx`.
- Ricompilato `Conta_Pollinica.exe` con il template aggiornato.

---

### Fix PermissionError su Windows alla pulizia file temporanei

**Problema:** su Windows (e Wine), alla chiusura della sessione
`pulisci_file_temporanei()` crashava con `PermissionError: [WinError 32]
Violazione di condivisione` tentando di eliminare o rinominare il file
`~autosave_*.xlsx` ancora aperto dalla GUI (thread di refresh).

**Causa:** la GUI legge periodicamente il file autosave con
`openpyxl.load_workbook()` per aggiornare le tabelle live. Su Windows il file
resta lockato brevemente durante la lettura, e se `unlink()` viene chiamato
in quel momento fallisce.

**Correzione** (`polline_counter.py`):
- Aggiunti helper `_safe_remove()` e `_safe_rename()` con retry (3 tentativi,
  pausa 0.5s). Se il file resta lockato, stampa un avviso senza crashare.
- `pulisci_file_temporanei()`: usa i nuovi helper al posto di `unlink()`
  e `rename()` diretti.
- Copiato `polline_counter.py` aggiornato in `windows/`.

---

### Soglie integrate nel template Excel

**Problema:** le soglie di concentrazione pollinica erano in un file separato
(`concentrazioni_polliniche.xlsx`). L'utente doveva assicurarsi di avere il
file nella cartella corretta, creando una dipendenza esterna fragile.

**Correzione:**
- Aggiunto foglio "soglie" a `Polline_Template_Settimanale.xlsx` (root e windows/)
  con tutti i dati copiati da `concentrazioni_polliniche.xlsx`.
- `carica_soglie()` (`polline_counter.py`): aggiunto parametro opzionale `wb=None`.
  Se il workbook contiene un foglio "soglie", legge da li'; altrimenti fallback
  al file esterno (che resta come riferimento).
- Estratta funzione `_parse_soglie_da_foglio()` per riuso del parsing.
- `genera_bollettino()`: passa `wb` a `carica_soglie(wb)`.
- `polline_counter_gui.py`: in `_leggi_dati_thread()`, ricarica le soglie dal
  workbook aperto con `carica_soglie(wb)` e aggiorna `self._soglie`. Cosi'
  funziona anche senza il file esterno.
- Il file `concentrazioni_polliniche.xlsx` non viene eliminato (fallback).
- Aggiunto `OUTPUT_DIR` ai candidati di ricerca del file soglie esterno: quando
  l'exe PyInstaller gira, `SCRIPT_DIR`/`BUNDLE_DIR` puntano alla cartella
  temporanea `_MEI*`, non alla cartella dell'exe. Ora cerca anche li'.
- Copiati `polline_counter.py` e `polline_counter_gui.py` aggiornati in `windows/`.

---

### Bollettino integrato nel foglio riepilogo_settimana

**Problema:** il bollettino pollinico veniva generato come foglio Excel
separato ("bollettino"), scomodo da consultare e non aggiornato
automaticamente durante l'inserimento dati.

**Correzione** (`polline_counter.py`):
- `genera_bollettino()` riscritta: scrive direttamente su `riepilogo_settimana`
  da riga 73 in poi (costante `BOLL_START_ROW = 73`), con pulizia preventiva
  dell'area (righe 72-115, colonne D-Y). Layout: separatore grigio (riga 72),
  titolo (riga 73), intestazioni gialle/blu (riga 75), dati (righe 76+),
  legenda colori in fondo. Non crea piu' il foglio separato.
- `autosave()`: aggiunti parametri opzionali `ws_riepilogo` e `lunedi`.
  Se forniti, chiama `genera_bollettino()` prima del salvataggio, cosi'
  il bollettino si aggiorna ad ogni autosave (ogni 5 inserimenti).
- Chiamate ad `autosave()` nel ciclo di inserimento e nel KeyboardInterrupt
  aggiornate per passare `ws_riepilogo` e `lunedi`.
- Comando "g": messaggio cambiato da "foglio 'bollettino'" a "nel riepilogo".
- `main()`: retrocompatibilita' — se nel workbook caricato esiste un foglio
  "bollettino", viene rimosso automaticamente.
- Copiato `polline_counter.py` aggiornato in `windows/`.

---

## 2025-09-03

### Riscrittura architettura rilevamento marker GUI

**Problema:** il popup di scelta cartella continuava a non apparire su
Windows. Il buffer anti-spezzamento nei poll function non era sufficiente:
il marker `__GUI_ASKDIR__` veniva comunque mostrato come testo nel
terminale integrato.

**Correzione** (`polline_counter_gui.py`):
- Riscritta `_handle_gui_markers()` con buffer di accumulo **interno
  persistente**: il testo viene accodato ad ogni poll e il buffer viene
  scansionato per marker completi. Il testo sicuro (prima del marker o
  senza marker parziali alla fine) viene ritornato per la visualizzazione.
- Rimossa tutta la logica di buffering da `_poll_output_win32()` e
  `_poll_output_unix()`: ora passano il testo crudo a `_handle_gui_markers`.
- Aggiunto parametro `flush` per svuotare il buffer alla chiusura del
  processo, e metodo `_flush_marker_buf()` come safety net.
- I dialoghi usano `parent=self.root`, `self.root.lift()` e
  `focus_force()` per apparire in primo piano.

### Applicata formattazione ai template Excel

**Problema:** i file `Polline_Template_Settimanale.xlsx` in tutte e tre
le cartelle (root, linux/, windows/) non avevano formattazione visiva:
nessun colore di sfondo, nessun bordo, nessun grassetto.

**Correzione:**
- Eseguito `applica_formattazione.py` che applica: sfondo verde sulle
  righe di specie rilevanti, bordi thin su tutte le celle dati,
  grassetto su intestazioni e specie principali, allineamento centrato.
- Formattazione applicata a tutti e tre i template (pollencounter/,
  linux/, windows/).

### Fix buffer marker spezzati (popup salvataggio non appariva)

**Problema:** nella GUI, il popup di scelta cartella (`askdirectory`) non
appariva. Il marker `__GUI_ASKDIR__` veniva mostrato come testo nel
terminale integrato anziche' essere intercettato.

**Causa:** il buffer anti-spezzamento controllava solo i primi 6 caratteri
del marker (`__GUI_`), ma la spezzatura tra due poll poteva avvenire a
qualsiasi punto dei 14-19 caratteri. Es. se un poll riceveva `__GUI_ASKD`
e il successivo `IR__`, nessuno dei due conteneva il marker completo.

**Correzione** (`polline_counter_gui.py`):
- Riscritto il buffer in `_poll_output_win32()`: ora controlla TUTTE le
  possibili lunghezze di coda (da 1 a 18 caratteri) contro i prefissi
  di entrambi i marker.
- Applicato lo stesso buffer anche a `_poll_output_unix()` (il pty puo'
  spezzare i dati allo stesso modo).
- Aggiunto `parent=self.root`, `self.root.lift()` e `focus_force()` ai
  dialoghi `askdirectory` e `askopenfilename` per assicurare che il
  popup appaia in primo piano.

### Fix nome file raddoppiato da autosave

**Problema:** riprendendo un file `Conta_Pollinica_~autosave_02-02-2026.xlsx`
e scegliendo "Salva su file nuovo", il nome suggerito diventava
`Conta_Pollinica_Conta_Pollinica_~autosave_02-02-2026.xlsx` (prefisso
raddoppiato).

**Causa:** `prima_data` veniva impostato a `file_ripreso.stem`, che per
un file gia' rinominato conteneva il prefisso `Conta_Pollinica_`. Poi
`chiedi_nome_file()` aggiungeva un secondo `Conta_Pollinica_`.

**Correzione** (`polline_counter.py`, `chiedi_nome_file()`):
- Aggiunto `re.sub()` che rimuove i prefissi noti (`Conta_Pollinica_`,
  `~autosave_`, `incompleto_`) da `data_str` prima di comporre il nome.

### Fix compatibilita' Windows (UnicodeEncodeError cp1252)

**Problema:** l'eseguibile Windows andava in crash con
`UnicodeEncodeError: 'charmap' codec can't encode character '\u2713'`
quando l'utente tentava di salvare il file alla chiusura.

**Causa:** il carattere Unicode `✓` (U+2713) non esiste nella codepage
cp1252 usata di default dalla console Windows.

**Correzione** (`polline_counter.py`, righe 417, 452, 461):
- Sostituito `✓` con `[OK]` (puro ASCII) in 3 messaggi di conferma
  dentro `menu_uscita_salvataggio()`.

### Fix finestra di salvataggio che non si apre su Windows

**Problema:** alla chiusura della GUI su Windows, la finestra nativa di
scelta cartella (`askdirectory`) non appariva. L'utente non poteva
scegliere dove salvare il file.

**Causa (1 — principale):** i `print()` dei marker `__GUI_ASKDIR__` e
`__GUI_ASKOPENFILE__` non facevano `flush`. Su Windows con pipe
bufferizzate il marker restava nel buffer di stdout e la GUI non lo
riceveva mai prima che `input()` bloccasse il processo.

**Causa (2 — secondaria):** su Windows il thread lettore legge 1 byte
alla volta e il poll ogni 50 ms assembla i byte disponibili. Il marker
(14+ caratteri) poteva arrivare spezzato su due poll consecutivi, non
venendo riconosciuto da `_handle_gui_markers()`.

**Correzione:**
- `polline_counter.py`: aggiunto `flush=True` ai `print` dei marker
  (righe 497 e 532).
- `polline_counter_gui.py`: aggiunto buffer di accumulo `_marker_buf`
  in `_poll_output_win32()` che trattiene i frammenti di marker
  incompleti e li ricompone al poll successivo.

### Cartella windows/ resa autocontenuta

**Problema:** la cartella `windows/` non conteneva i file `.py`, e i
file `.bat` puntavano alla cartella padre con `%~dp0..`.

**Correzione:**
- Copiati `polline_counter.py` e `polline_counter_gui.py` aggiornati
  dentro `windows/`.
- `build_exe.bat`: cambiato `cd /d "%~dp0.."` in `cd /d "%~dp0"`.
- `AVVIA_CONTA_POLLINICA.bat`: cambiato `"%~dp0..\polline_counter_gui.py"`
  in `"%~dp0polline_counter_gui.py"`.
- La cartella `windows/` ora contiene tutto il necessario per eseguire
  e compilare il programma autonomamente.
