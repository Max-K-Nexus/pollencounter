# Opzioni di distribuzione centralizzata — Conta Pollinica

Documento di riferimento per scegliere una strategia di aggiornamento automatico.
Redatto: 2026-02-27

---

## Contesto del problema

Attualmente ogni aggiornamento richiede:
- Modificare i `.py` sul PC di sviluppo
- Copiare i file in `windows/`
- Ricompilare `Conta_Pollinica.exe` via Wine (~2-3 min)
- Distribuire manualmente il nuovo exe sui PC degli utenti

L'obiettivo è che un aggiornamento applicato una volta raggiunga tutti i sistemi
automaticamente o con il minimo intervento manuale.

---

## Opzione A — Cartella condivisa di rete (basso sforzo, subito)

**Principio:** i file `.py` stanno su una cartella di rete condivisa (NAS,
OneDrive universitario, Samba). I launcher locali puntano a quel percorso.
Aggiornamento: modifichi i file sul server, tutti usano la versione nuova
al prossimo avvio. Zero compilazione.

**Requisiti:**
- Python installato su ogni PC (già necessario per eseguire i .py)
- Rete condivisa accessibile (OneDrive, cartella di laboratorio, NAS)

**Launcher Linux** (`AVVIA_CONTA_POLLINICA_GUI.sh`):
```bash
#!/bin/bash
python3 /percorso/condiviso/pollencounter/polline_counter_gui.py
```

**Launcher Windows** (`AVVIA.bat`):
```bat
@echo off
python "%USERPROFILE%\OneDrive - Universita\pollencounter\polline_counter_gui.py"
```

**Limiti:**
- Non funziona offline
- Richiede Python su ogni macchina Windows (non risolve il problema exe)
- Non adatto se le macchine non hanno rete comune

**Adatto se:** tutti i PC sono in LAN universitaria o usano OneDrive condiviso.

---

## Opzione B — Repository Git + script di aggiornamento (basso sforzo)

**Principio:** il codice sta su GitHub/GitLab privato. Ogni macchina ha
una copia clonata. Un script `aggiorna.sh` / `aggiorna.bat` fa `git pull`.

**Script Linux** (`aggiorna.sh`):
```bash
#!/bin/bash
cd /home/utente/pollencounter
git pull origin main
echo "Aggiornamento completato."
```

**Script Windows** (`aggiorna.bat`):
```bat
@echo off
cd /d "%~dp0"
git pull origin main
pause
```

**Limiti:**
- Richiede `git` installato su ogni macchina
- L'exe Windows va ancora ricompilato dopo ogni modifica
- Gli utenti devono ricordarsi di eseguire lo script di aggiornamento

**Adatto se:** le macchine sono gestite da una persona tecnica (tu),
e vuoi anche storico versioni e rollback.

---

## Opzione C — Auto-update integrato nell'applicazione (medio sforzo)

**Principio:** all'avvio la GUI controlla un file `version.txt` su GitHub
o un server. Se c'è una versione più recente, avvisa l'utente e offre
un pulsante "Aggiorna ora" che scarica e sostituisce i file.

**Componenti da sviluppare:**
1. File `version.txt` su GitHub Releases o server web (es. `1.4.2`)
2. Funzione di controllo versione nella GUI (chiamata HTTP all'avvio)
3. Downloader in background con progress bar
4. Su Windows: launcher separato `updater.exe` che sostituisce
   il `.exe` principale (non puoi sovrascrivere un exe in esecuzione)

**Sforzo stimato:** 1-2 giorni di sviluppo + gestione GitHub Releases.

**Limiti:**
- Dipende da connessione internet affidabile
- Complessità aggiuntiva (gestione errori, versioning, firma del codice)
- Su Windows il meccanismo di sostituzione exe richiede un launcher separato

**Adatto se:** i PC potrebbero non avere rete condivisa ma hanno internet,
e vuoi un'esperienza utente completamente autonoma.

---

## Opzione D — Applicazione web (alto sforzo, massima centralizzazione)

**Principio:** il software diventa un sito web (Flask backend + HTML frontend).
Gli utenti aprono il browser, inseriscono i dati, scaricano l'Excel generato.
Zero installazione sui client. Aggiornamento istantaneo: modifichi il server,
tutti vedono subito la versione nuova.

**Stack minimo:**
```
Flask (o FastAPI)   ← server web Python
openpyxl            ← generazione Excel (codice esistente riusabile)
HTML/JS             ← interfaccia (sostituisce tkinter)
```

**Riuso del codice esistente:**
- Tutta la logica di `polline_counter.py` è riusabile quasi integralmente
- `sessione_giorno()`, `genera_bollettino()`, `esporta_riepilogo_annuale()`
  diventano funzioni chiamate da endpoint Flask
- La GUI tkinter viene sostituita da una pagina HTML

**Infrastruttura necessaria:**
- Un server sempre acceso in LAN (anche un Raspberry Pi a ~50€)
  oppure un server universitario o un cloud economico (es. PythonAnywhere)

**Limiti:**
- Riscrittura parziale dell'interfaccia (qualche giorno di lavoro)
- Richiede connessione di rete (anche solo LAN)
- Gestione sessioni utente multiple (se più persone usano il sistema insieme)

**Adatto se:** il progetto cresce, si aggiungono più laboratori/sedi,
o si vuole eliminare definitivamente il problema multi-OS.

---

## Matrice di confronto

| Criterio              | A (Rete) | B (Git) | C (Auto-update) | D (Web app) |
|-----------------------|:--------:|:-------:|:---------------:|:-----------:|
| Sforzo implementazione| Basso    | Basso   | Medio           | Alto        |
| Funziona offline      | No       | Si      | Si              | No          |
| Aggiornamento auto    | Si       | No      | Si              | Si          |
| Richiede Python client| Si       | Si      | No (exe)        | No          |
| Richiede server       | NAS/LAN  | No      | No              | Si          |
| Scalabile (N sedi)    | No       | Medio   | Si              | Si          |
| Adatto utenti non tecnici | Si   | No      | Si              | Si          |

---

## Raccomandazione

**Subito:** Opzione A (cartella condivisa) se esiste gia' una rete o
OneDrive condiviso. Costo: ~30 minuti per modificare i launcher.

**Medio termine:** Opzione C (auto-update) se si vuole mantenere il formato
exe per Windows ma avere aggiornamenti automatici.

**Lungo termine:** Opzione D (web app) se il progetto si espande a piu'
laboratori o si vuole eliminare tutti i problemi di distribuzione OS.
