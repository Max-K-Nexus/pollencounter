🌿 Pollencounter
Pollencounter è uno strumento per il conteggio pollinico, progettato per supportare il monitoraggio aerobiologico e la produzione di bollettini ufficiali.
L'obiettivo principale è ridurre il lavoro manuale sui file Excel, standardizzare i calcoli e semplificare la gestione delle letture settimanali e dei report. Non è necessario essere programmatori per usare la versione Windows o la GUI.

✨ Funzionalità principali

✅ Automazione — Riduce drasticamente l'inserimento manuale e i calcoli ripetitivi.
✅ Standardizzazione — Calcoli uniformi per garantire la qualità dei dati aerobiologici.
✅ Versatilità — Utilizzabile tramite interfaccia grafica (GUI) o script Python da riga di comando (CLI).
✅ Multipiattaforma — Supporto nativo per Windows (eseguibile .exe), Linux e macOS.
✅ Bollettini Word — Genera automaticamente i bollettini pollinici in italiano e inglese (.docx).
✅ Riepilogo annuale — Esporta i dati di tutte le settimane in un file Excel annuale con foglio Calendario e fogli settimanali W##.
✅ Autosave — Salvataggio automatico ogni 5 inserimenti; nessun dato viene mai perso.


📂 Struttura della repository
pollencounter/
├── codice/                          # Script principali e file di riferimento
│   ├── polline_counter.py           # Logica di elaborazione (CLI, cross-platform)
│   ├── polline_counter_gui.py       # Versione con Interfaccia Grafica (GUI)
│   ├── Polline_Template_Settimanale.xlsx          # Template base per i calcoli
│   ├── concentrazioni_polliniche.xlsx             # Soglie concentrazioni polliniche (fallback)
│   ├── ITA_Template_Bollettino_pubblicazione.docx # Template bollettino italiano
│   ├── ENG_Template_Bollettino_pubblicazione.docx # Template bollettino inglese
│   └── pollencounter.cfg            # Configurazione cartella di lavoro per anno
├── script_aiuto/                    # Utility e script di avvio (Linux)
│   ├── AVVIA_CONTA_POLLINICA_GUI.sh # Avvio GUI da terminale (Linux)
│   ├── AVVIA_CONTA_POLLINICA.sh     # Avvio CLI da terminale (Linux)
│   ├── applica_formattazione.py     # Utility manutenzione template (eseguire manualmente)
│   └── setup_bollettino_template.py # Utility per rigenerare la sezione bollettino nel template
├── letture_settimanali/             # Cartella INPUT (file di conteggio settimanali)
├── riferimenti/                     # Tabelle storiche e dati di riferimento
├── mac/                             # Risorse specifiche per utenti macOS
│   ├── AVVIA_CONTA_POLLINICA_GUI.sh # Avvio GUI da Terminale (macOS)
│   ├── AVVIA_CONTA_POLLINICA.sh     # Avvio CLI da Terminale (macOS)
│   ├── build_app.sh                 # Script per creare Conta_Pollinica.app (PyInstaller)
│   └── ISTRUZIONI_MAC.txt           # Guida specifica per utenti macOS
├── windows/                         # Risorse specifiche per utenti Windows
│   ├── AVVIA_CONTA_POLLINICA.bat    # Script di avvio rapido
│   ├── build_exe.bat                # Script per compilare l'eseguibile (sviluppo)
│   └── ISTRUZIONI_WINDOWS.txt       # Guida specifica per utenti Windows
├── esempio di bolletino.pdf         # Esempio di output finale
├── CHANGELOG.md                     # Cronologia delle modifiche
├── ISTRUZIONI.txt                   # Documentazione generale
└── CLAUDE.md                        # Guida tecnica per sviluppatori

Nota per Windows: la cartella windows/ non include l'eseguibile precompilato Conta_Pollinica.exe perché il file supera i limiti di dimensione per la condivisione su GitHub (~12 MB). Per ottenerlo, compilarlo localmente con windows/build_exe.bat seguendo le istruzioni in windows/ISTRUZIONI_WINDOWS.txt.


🔄 Flusso di lavoro tipico

Raccolta dati — L'operatore compila i file Excel di conteggio settimanale nella cartella letture_settimanali/, seguendo la convenzione Conta_Pollinica_GG-MM-AAAA.xlsx.
Elaborazione — L'utente avvia l'applicazione tramite la GUI o gli script Python. Il programma legge i file di input, i template e la configurazione nella cartella codice/.
Output — Lo strumento genera file Excel aggiornati, bollettini Word (ITA/ENG) pronti per la pubblicazione e un riepilogo annuale.


🚀 Guida all'uso
🪟 Utenti Windows (non tecnici)
Questa modalità non richiede l'installazione di Python.

Scaricare il progetto da GitHub (Code → Download ZIP) ed estrarlo.
Aprire la cartella windows/.
Leggere il file ISTRUZIONI_WINDOWS.txt.
Avviare l'applicazione facendo doppio clic su AVVIA_CONTA_POLLINICA.bat.

🍎 Utenti macOS

Aprire la cartella mac/ e leggere ISTRUZIONI_MAC.txt.
Avviare la GUI con:

bash   ./mac/AVVIA_CONTA_POLLINICA_GUI.sh

Per creare l'app bundle nativa (.app) su un Mac, eseguire mac/build_app.sh.

🐍 Utenti Python — Linux / sviluppatori
Clonare la repository:
bashgit clone https://github.com/Max-K-Nexus/pollencounter.git
cd pollencounter
Installare le dipendenze:
bashpip install openpyxl
pip install python-docx   # opzionale — richiesto per la generazione dei bollettini Word
Avviare la GUI:
bashpython3 codice/polline_counter_gui.py
Oppure la versione CLI:
bashpython3 codice/polline_counter.py
In alternativa, usare gli script nella cartella script_aiuto/:
bash./script_aiuto/AVVIA_CONTA_POLLINICA_GUI.sh   # GUI
./script_aiuto/AVVIA_CONTA_POLLINICA.sh        # CLI

⚙️ Configurazione e convenzioni

File .cfg — codice/pollencounter.cfg permette di impostare la cartella di lavoro per anno senza modificare il codice. Viene creato automaticamente al primo avvio.
Convenzione nomi file — I file in letture_settimanali/ devono seguire il formato Conta_Pollinica_GG-MM-AAAA.xlsx.
Template Excel — Polline_Template_Settimanale.xlsx non va modificato manualmente. Per aggiornare la formattazione usare script_aiuto/applica_formattazione.py.


🛠 Risoluzione problemi (FAQ)
L'eseguibile non si avvia?
Assicurarsi di aver estratto correttamente l'archivio ZIP e che l'antivirus non stia bloccando il file.
Errori con i file Excel?
Verificare di non aver modificato o spostato la struttura delle colonne nei template nella cartella codice/.
Bollettini Word non generati?
È richiesta la libreria python-docx. Installarla con:
bashpip install python-docx
# oppure su Debian/Ubuntu:
sudo apt install python3-docx
Uso Mac o Linux?
Il file .exe è esclusivamente per Windows. Su altri sistemi usare gli script Python in codice/.

👥 Autori e contributi
Il progetto Pollencounter è stato sviluppato da:

Simone Bettella — Concept, sviluppo originale e logica di calcolo.
Massimiliano Iotti — Manutenzione, automazione, documentazione e supporto multipiattaforma.


📜 Licenza
Il progetto è distribuito sotto la GNU General Public License v3.0 (GPL-3.0).
Se utilizzi Pollencounter in un progetto o report, si prega di citare gli autori:

Pollencounter – developed by Simone Bettella and Massimiliano Iotti
https://github.com/Max-K-Nexus/pollencounter
