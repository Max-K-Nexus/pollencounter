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

## Architettura

Riscrittura del guscio completata il 2026-09-22 (vedi CHANGELOG). Un solo
modello di dominio, condiviso da CLI e GUI, che girano nello stesso processo
dell'interfaccia: **non esiste più un sottoprocesso CLI pilotato dalla GUI**,
né un protocollo di marker nello stdout, né un autosave periodico dell'intero
workbook con rilettura da file. Sei moduli in `codice/`:

| Modulo | Ruolo | I/O |
|---|---|---|
| `dominio.py` | Codici specie, date, soglie, calcolo concentrazione/livello, struttura del bollettino (`BOLLETTINO_RIGHE`), interpretazione dei comandi da tastiera (`interpreta_comando`) | **Nessuno** — puro, interamente coperto da `tests/test_dominio.py` |
| `sessione.py` | Modello in memoria `Settimana` (conteggi, log, storico, undo) e `Journal` (persistenza incrementale JSONL) | Legge/scrive il journal e i file `.xlsx` in import |
| `esportatori.py` | `esporta_xlsx`, `esporta_riepilogo_annuale`, `genera_bollettini_word`, `carica_soglie` | Scrive `.xlsx`/`.docx`, legge il template e `concentrazioni_polliniche.xlsx` |
| `voce.py` | Lettura vocale (opzionale): `Ascoltatore` (microfono -> riconoscitore -> coda di frasi), `Sintesi` (talkback), `crea_voce`. Il motore (Vosk) e' dietro un'interfaccia sostituibile | Microfono, altoparlanti, modello Vosk in `modelli/` |
| `percorsi.py` | Risoluzione `BUNDLE_DIR`/`SCRIPT_DIR`/`EXE_DIR`/`CONFIG_FILE` (frozen vs sorgente) | — |
| `polline_counter.py` (CLI) / `polline_counter_gui.py` (GUI) | Interfacce, entrambe sopra lo stesso `sessione.Settimana` | Input utente (terminale o tkinter) |

### Modello e persistenza (`sessione.py`)

- **`Settimana`**: `conteggi` (`{codice: [7 interi]}`), `log` (righe per il
  foglio `dati_grezzi`), `storico`, un solo undo-stack per il giorno
  correntemente attivo (`attiva_giorno()` lo azzera — stesso comportamento
  dell'originale: l'annullo funziona solo nella sessione del giorno in corso).
  Notifica i cambiamenti con `on_change(callback)`: la GUI vi si aggancia per
  ridisegnare le tab **in modo sincrono, nello stesso processo** — niente
  polling, niente race condition fra "ultimo dato letto da file" e "ultimo
  delta ricevuto" (era il difetto principale della vecchia architettura).
- **`Journal`**: ogni operazione (`inserisci`, `annulla`, `correggi`,
  `aggiungi_nota`, `attiva_giorno`) scrive subito una riga JSONL su
  `~sessione_<lunedi>.jsonl`, con `flush()` + `os.fsync()`. La prima riga
  (`"t": "inizio"`) è uno snapshot completo dello stato di partenza; il resto
  sono eventi incrementali. `sessione.ripristina_da_journal()` ricostruisce lo
  stato rigiocando le stesse chiamate di modello (non serializza/deserializza
  a mano): è così che CLI e GUI recuperano una sessione dopo un crash, offerta
  all'avvio insieme ai file `.xlsx` già salvati. Il salvataggio su `.xlsx`
  resta un'azione esplicita dell'utente (comando/pulsante `s` o `q`); dopo un
  salvataggio riuscito il journal viene eliminato.

### Lettura vocale (`dominio.interpreta_vocale` + `voce.py`)

La voce e' una **seconda sorgente di comandi** sopra lo stesso percorso della
tastiera: `dominio.interpreta_vocale()` produce gli stessi oggetti di
`interpreta_comando()` (+ `Totale`) e la GUI li esegue con
`_esegui_comando(cmd, fonte="voce")`, la stessa funzione usata da `_invia()`:
stesso `Settimana.inserisci(..., journal=...)`, quindi journal, undo e tab live
funzionano senza modifiche. Idee prese da EcoCount (Allen & Sewell 2014, SAGE
Open 4(2)): parola di attivazione, talkback di cio' che e' stato capito,
annullo a voce, dizionario dei taxa personalizzabile.

- **Parola di attivazione "conta"** obbligatoria su ogni frase (anche
  "conta annulla"): la sola grammatica chiusa non basta, "due/tre/sei" sono
  parole comuni. Disattivabile con `"parola_attivazione": ""` in
  `pollencounter.cfg`.
- **Solo azioni reversibili a voce** (`dominio.COMANDI_VOCALI`: annulla,
  ripeti, ancora, totale). Salva/chiudi giornata/esci restano da tastiera: un
  falso riconoscimento non deve poter chiudere o perdere nulla (test dedicato).
- **Grammatica chiusa** (`costruisci_grammatica`) passata a Vosk con `[unk]`;
  frasi sotto soglia di confidenza o con `[unk]` vengono scartate in silenzio.
  Quantita' a voce solo 2-20 (grammatica piu' piccola); da tastiera resta 1-100.
- **Anti-eco:** durante la sintesi l'audio in ingresso e' scartato
  (`Ascoltatore.silenzia`), altrimenti il PC registrerebbe la propria voce.
- **Dialogo aperto = voce ignorata** (`root.grab_current()` in
  `_gestisci_frase_vocale`). tkinter non e' thread-safe: il thread di ascolto
  riempie solo una coda, svuotata da `root.after(100, ...)`.
- Sinonimi personali: chiave `sinonimi_vocali` di `pollencounter.cfg`
  (`sessione.leggi_sinonimi_vocali`). Dipendenze `vosk`/`sounddevice`/
  `pyttsx3` e modello in `codice/modelli/` (non in git) sono **opzionali**.
- Il modello piccolo italiano **non conosce molti nomi latini/famiglie**:
  `crea_voce` li individua con `vosk_model_find_word`, toglie le frasi
  corrispondenti dalla grammatica e li segnala; per quelle specie valgono
  numero e nomi comuni (`SINONIMI_VOCALI_DEFAULT`). La voce di sintesi viene
  scelta italiana (`scegli_voce_italiana`).
- Provato con motore reale su audio sintetico (espeak-ng) e, il 2026-10-09,
  dall'utente al microfono su Linux ("conta acero", "conta ontano", "conta
  acero per tre": OK). L'exe Windows con la voce e' compilato e si avvia sotto Wine, ma
  non e' provato su Windows vero (microfono, SAPI5), ne' in ambienti
  rumorosi o con altri accenti.

### Bollettino: un'unica fonte di colore (`dominio.righe_bollettino`)

`dominio.righe_bollettino(conteggi, fattore, soglie)` calcola, per ogni specie
di `BOLLETTINO_RIGHE`, sia le concentrazioni dei 7 giorni sia il **livello per
ciascun giorno** (`livelli_giorno`, non solo un livello aggregato sulla
media). Sia l'anteprima nella tab Bollettino della GUI sia
`esportatori.genera_bollettini_word()` chiamano questa stessa funzione: non
possono più divergere (era il difetto #3 della revisione — l'anteprima
colorava per media settimanale, il Word per singolo giorno).

### `polline_counter_gui.py` (GUI)

Finestra tkinter divisa in due pannelli, **stesso processo, nessun
sottoprocesso**:
- **Sinistra:** barra pulsanti giorno (LUN…DOM), riquadro di log (sostituisce
  il vecchio terminale emulato) e casella di inserimento che accetta gli
  stessi comandi della CLI via `dominio.interpreta_comando()`
- **Destra:** notebook con tre tab live (Settimanale, Giornaliero,
  Bollettino), aggiornate dal callback `on_change` del modello

La schermata iniziale (scelta cartella, sessioni journal recuperabili, file
`.xlsx`, nuovo file, importa) usa dialoghi nativi tkinter
(`filedialog`/`simpledialog`/`messagebox`) direttamente — non c'è più bisogno
di un protocollo di marker perché non c'è più un processo figlio a cui
chiedere di aprirli.

**Font:** `_MONO_FONT = "Courier New"` su Windows, `"Monospace"` su Linux.

**`sv_ttk`:** tema opzionale (`try/except`). Attivato solo su Windows in `main()`.

### Test (`codice/tests/`)

`unittest` della stdlib (nessuna dipendenza aggiuntiva). Eseguire da `codice/`:
```
python3 -m unittest discover -s tests
```
`test_dominio.py` copre l'intero modulo puro; `test_sessione.py` copre modello
e journal (incluso il replay dopo crash); `test_esportatori.py` copre
l'esportazione `.xlsx`/annuale/`.docx` (i test sul bollettino Word si
saltano automaticamente se `python-docx` non è installato). `test_voce.py`
copre `voce.py` con un riconoscitore finto (anti-eco, soglia di confidenza,
errori del microfono, scelta della voce italiana, ricerca del modello).

---

## Convenzioni da rispettare

- **Lingua:** tutto il testo mostrato all'utente è in italiano.
- **Caratteri ASCII only nelle stampe della CLI:** non usare caratteri Unicode
  fuori cp1252 (es. `✓`, `─`, emoji) nelle stringhe stampate a stdout in
  `polline_counter.py`. Windows con cp1252 va in crash. Usare alternative
  ASCII (es. `[OK]` al posto di `✓`). Non si applica ai widget tkinter della
  GUI (Tk gestisce Unicode correttamente su tutte le piattaforme): lì i testi
  possono usare accenti veri.
- **Template Excel:** non modificare la struttura dei fogli `riepilogo_settimana`
  e `dati_grezzi`. Le righe sono fisse: pollini in righe 6–52, spore in 58–69
  (funzioni `dominio.codice_to_row`, `dominio.giorno_to_col`).
- **Percorsi (`percorsi.py`):** `BUNDLE_DIR`/`SCRIPT_DIR` per i file inclusi
  nel bundle (template, soglie, `.docx`); `EXE_DIR` per la cartella accanto
  all'eseguibile/script (dove risiede `pollencounter.cfg` e, di default, dove
  si salva); la cartella di lavoro effettiva è quella scelta dall'utente e
  salvata in `pollencounter.cfg` per anno (`sessione.leggi_cartella_anno` /
  `salva_cartella_anno`).
- **Journal di sessione:** il file `~sessione_<lunedi>.jsonl` viene trovato da
  `sessione.recupera_sessioni()` al prossimo avvio e presentato come sessione
  recuperabile (sostituisce il vecchio `~autosave_*.xlsx`). `cerca_file_ripresa()`
  esclude template e `Riepilogo_Annuale_*.xlsx` e verifica la presenza del
  foglio `riepilogo_settimana` prima di proporre un file (vedi CHANGELOG,
  difetto #2 della revisione).
- **Soglie di concentrazione:** un'unica fonte, `concentrazioni_polliniche.xlsx`
  (`esportatori.carica_soglie`). Non leggere più un foglio `soglie` dal
  workbook di sessione né usare tabelle di soglie imbustate nel codice, se non
  come fallback estremo (`dominio.SOGLIE_FALLBACK`, usato solo se il file
  esterno non si trova).
- **Cartella `windows/` non autocontenuta:** contiene solo i `.bat`, l'exe e le istruzioni. I sorgenti `.py` e i file `.xlsx` risiedono in `codice/`; `build_exe.bat` e `AVVIA_CONTA_POLLINICA.bat` li referenziano con path `..\codice\`. Non copiare i sorgenti in `windows/`.

---

## Dipendenze

```
openpyxl      ← obbligatorio (lettura/scrittura Excel)
tkinter       ← obbligatorio per la GUI (su Debian: sudo apt install python3-tk)
python-docx   ← opzionale (bollettini Word; sudo apt install python3-docx)
sv_ttk        ← opzionale (tema Windows; pip install sv-ttk)
pyinstaller   ← solo per build Windows exe (vedi sezione sotto)
winsound      ← solo Windows, incluso nella stdlib
```

---

## Build dell'eseguibile Windows (.exe)

Per ricompilare `Conta_Pollinica.exe` (Wine + PyInstaller Windows), vedi la skill
`build-windows-exe` (`.claude/skills/build-windows-exe/SKILL.md`).
`windows/build_exe.bat` e `mac/build_app.sh` includono la lettura vocale solo se
`codice/modelli/vosk-model*` esiste (altrimenti build senza voce, come prima);
quello per macOS non e' verificato.

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

Per aggiornare la formattazione visiva del template Excel (`applica_formattazione.py`),
vedi la skill `aggiorna-formattazione-template`
(`.claude/skills/aggiorna-formattazione-template/SKILL.md`).

---

## Note operative per sessioni future

- Prima di modificare qualsiasi funzione, leggere il file per intero: `dominio.py`,
  `sessione.py` ed `esportatori.py` sono strettamente accoppiati fra loro e con
  entrambe le interfacce (CLI e GUI).
- Il codice è usato in produzione da utenti non tecnici: privilegiare stabilità e messaggi chiari rispetto a refactoring.
- Le modifiche alla struttura di `Settimana` o del `Journal` vanno verificate
  su **entrambe** le interfacce (CLI e GUI) e sui test in `codice/tests/`:
  condividono lo stesso modello, quindi una modifica incompatibile rompe
  entrambe silenziosamente. Eseguire `python3 -m unittest discover -s tests`
  prima di considerare finita una modifica al dominio.
- I sorgenti `.py` sono unici (in `codice/`). Non creare copie in `windows/`.
- **Dopo ogni modifica:** aggiornare `CHANGELOG.md` (vedi sezione sopra). Questo e' obbligatorio.
- **Dopo ogni spostamento di script o cambio di percorsi:** aggiornare il launcher Desktop
  (`/home/Simone/Scrivania/ContaPollinica.desktop` sulla macchina di produzione — non
  presente su questo ambiente di sviluppo, verificarlo manualmente), campo `Exec=`.
  Percorso atteso: `bash .../pollencounter/script_aiuto/AVVIA_CONTA_POLLINICA_GUI.sh`.
- Per testare: `python3 codice/polline_counter.py` (CLI) e `python3 codice/polline_counter_gui.py` (GUI) dalla directory `pollencounter/`, oppure direttamente dalla cartella `codice/`.
- **Per ricompilare l'exe:** usare Wine + Python Windows (vedi sezione "Build dell'eseguibile Windows"). Non usare PyInstaller Linux nativo.
- `script_aiuto/setup_bollettino_template.py` (rotto: importava
  `BOLL_START_ROW`, mai definito) è stato eliminato il 2026-10-03; se serve
  va rifatto da zero.
- I documenti tecnici non destinati agli utenti (revisione del 22/09/2026,
  prompt web app, opzioni di distribuzione) sono in `docs/`.
- Gli script di avvio macOS sono `mac/*.command` (doppio clic dal Finder),
  non più `.sh`.
